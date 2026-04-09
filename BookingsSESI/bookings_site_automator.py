"""
Automação do Microsoft Bookings com Playwright
=============================================

Objetivo
--------
Automatizar o preenchimento do site do Bookings já autenticado no navegador,
para criar agendamentos recorrentes em lote para várias turmas.

Fluxo
-----
1. Primeira execução:
   - abre o navegador
   - o usuário faz login manualmente na Microsoft
   - o programa salva a sessão autenticada

2. Próximas execuções:
   - reusa a sessão salva
   - abre o site do Bookings
   - navega pelo formulário
   - preenche os campos
   - envia os agendamentos em lote

Importante
----------
- Este script usa Playwright com perfil persistente separado.
- NÃO usa coordenadas de mouse fixas como estratégia principal.
- Os seletores do site podem variar; por isso, ficam em um JSON externo.
- Se a estrutura do Bookings mudar, ajuste o arquivo selectors_bookings_sesi.json.

Instalação
----------
pip install playwright
playwright install chromium

Execução
--------
python bookings_site_automator.py

Modo de uso
-----------
Ao iniciar, o programa pergunta:
1) Quer salvar/atualizar o login?
2) Quer rodar uma automação de teste?
3) Quer executar os agendamentos do arquivo JSON?

Arquivos esperados
------------------
- selectors_bookings_sesi.json
- bookings_jobs_exemplo.json  (você pode duplicar e criar o seu)

Base técnica
------------
Playwright permite reutilizar estado autenticado com storage_state e também
trabalhar com contexto persistente, o que é adequado para reutilizar a sessão
entre execuções. citeturn523490view0turn523490view1
"""

from __future__ import annotations

import json
import re
import shutil
import sys
import time
import datetime as dt
import threading
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import BooleanVar, StringVar, ttk, messagebox
from typing import Any, Dict, List, Optional
from urllib.parse import urlparse

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright, Page, BrowserContext

import msvcrt


DEFAULT_BOOKINGS_URL = ""
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
APP_NAME = "SESI Reservas Recorrentes"
APP_SUBTITLE = "Agendamentos recorrentes com automação assistida"
MAX_RECURRENCE_DAYS = 31
AUTH_DIR = BASE_DIR / "playwright_auth"
AUTH_DIR.mkdir(exist_ok=True)
STORAGE_STATE_PATH = AUTH_DIR / "storage_state.json"
PROFILE_DIR = AUTH_DIR / "chromium_profile"
LOGGED_USER_CACHE_PATH = AUTH_DIR / "logged_user.txt"
SETTINGS_PATH = BASE_DIR / "bookings_app_settings.json"

SELECTORS_PATH = BASE_DIR / "selectors_bookings_sesi.json"
JOBS_PATH = BASE_DIR / "bookings_jobs_exemplo.json"
LOGO_CANDIDATES = [
    BASE_DIR / "sesi_logo_app.png",
    BASE_DIR / "imagens" / "sesi_logo_app.png",
    BASE_DIR / "imagens" / "sesi_logo_vermelha.png",
    BASE_DIR / "sesi_logo_recortada.png",
    BASE_DIR / "sesi_logo_vermelha.png",
    BASE_DIR / "imagens" / "sesi_logo_recortada.png",
]


def load_app_settings() -> Dict[str, Any]:
    if not SETTINGS_PATH.exists():
        return {"bookings_url": DEFAULT_BOOKINGS_URL}
    try:
        raw = json.loads(SETTINGS_PATH.read_text(encoding="utf-8-sig"))
        if isinstance(raw, dict):
            raw.setdefault("bookings_url", DEFAULT_BOOKINGS_URL)
            return raw
    except Exception:
        pass
    return {"bookings_url": DEFAULT_BOOKINGS_URL}


def save_app_settings(settings: Dict[str, Any]) -> None:
    payload = {"bookings_url": str(settings.get("bookings_url", "")).strip()}
    SETTINGS_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def normalize_bookings_url(value: str) -> str:
    return str(value or "").strip()


def is_valid_bookings_url(value: str) -> bool:
    url = normalize_bookings_url(value)
    if not url:
        return False
    try:
        parsed = urlparse(url)
    except Exception:
        return False
    return (
        parsed.scheme in {"http", "https"}
        and "outlook.office.com" in parsed.netloc.lower()
        and "/book/" in parsed.path.lower()
    )


def get_bookings_url() -> str:
    return normalize_bookings_url(load_app_settings().get("bookings_url", DEFAULT_BOOKINGS_URL))


def get_bookings_resource_email(bookings_url: Optional[str] = None) -> str:
    url = normalize_bookings_url(bookings_url or get_bookings_url())
    if not url:
        return ""
    try:
        return (urlparse(url).path.split("/")[-1] or "").strip().lower()
    except Exception:
        return ""


def require_bookings_url() -> str:
    url = get_bookings_url()
    if not is_valid_bookings_url(url):
        raise ValueError("Informe e salve o link do Bookings da unidade antes de continuar.")
    return url


@dataclass
class BookingJob:
    escolha_reserva: str
    dias_semana: list[str]
    equipe: list[str]
    horario: str
    notas: str
    componente: str
    publico: str
    turmas: list[dict]  # [{"turma": str, "dias_semana": list[str], "horario": str}, ...]
    principal_recurso: str
    tipo_atividade: str
    data_inicio: str
    data_fim: str
    confirmar_envio_real: bool = False


log_text_widget = None
results_text_widget = None

def log(msg: str) -> None:
    print(f"[LOG] {msg}")
    if log_text_widget:
        log_text_widget.config(state="normal")
        log_text_widget.insert(tk.END, f"[LOG] {msg}\n")
        log_text_widget.see(tk.END)
        log_text_widget.config(state="disabled")

def warn(msg: str) -> None:
    print(f"[WARN] {msg}")
    if log_text_widget:
        log_text_widget.config(state="normal")
        log_text_widget.insert(tk.END, f"[WARN] {msg}\n")
        log_text_widget.see(tk.END)
        log_text_widget.config(state="disabled")

def fail(msg: str) -> None:
    print(f"[ERRO] {msg}")
    if log_text_widget:
        log_text_widget.config(state="normal")
        log_text_widget.insert(tk.END, f"[ERRO] {msg}\n")
        log_text_widget.see(tk.END)
        log_text_widget.config(state="disabled")

def safe_add_result(msg: str, root: tk.Tk) -> None:
    if results_text_widget:
        root.after(0, lambda: results_text_widget.config(state="normal"))
        root.after(0, lambda: results_text_widget.insert(tk.END, f"{msg}\n"))
        root.after(0, lambda: results_text_widget.see(tk.END))
        root.after(0, lambda: results_text_widget.config(state="disabled"))


def load_json(path: Path) -> Dict[str, Any]:
    if not path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {path}")
    return json.loads(path.read_text(encoding="utf-8-sig"))


def load_jobs(path: Path) -> List[BookingJob]:
    raw = load_json(path)
    items = raw.get("jobs", [])
    jobs: List[BookingJob] = []
    for item in items:
        validate_recurrence_period(item["data_inicio"], item["data_fim"])
        jobs.append(
            BookingJob(
                escolha_reserva=item["escolha_reserva"],
                dias_semana=item["dias_semana"],
                equipe=item["equipe"],
                horario=item["horario"],
                notas=item.get("notas", ""),
                componente=item.get("componente", "Outros"),
                publico=item["publico"],
                turmas=item["turmas"],  # Now list of dicts
                principal_recurso=item["principal_recurso"],
                tipo_atividade=item["tipo_atividade"],
                data_inicio=item["data_inicio"],
                data_fim=item["data_fim"],
                confirmar_envio_real=bool(item.get("confirmar_envio_real", False)),
            )
        )
    return jobs


def parse_br_date(date_value: str) -> dt.date:
    return dt.datetime.strptime(date_value, "%d/%m/%Y").date()


def validate_recurrence_period(start_date: str, end_date: str) -> None:
    start = parse_br_date(start_date)
    end = parse_br_date(end_date)
    if end < start:
        raise ValueError("A data fim deve ser igual ou posterior à data início.")
    if (end - start).days > MAX_RECURRENCE_DAYS:
        raise ValueError("A recorrência pode ter no máximo 1 mês entre a data início e a data fim.")


def load_logo_image(subsample: int = 8) -> Optional[tk.PhotoImage]:
    for path in LOGO_CANDIDATES:
        if path.exists():
            try:
                image = tk.PhotoImage(file=str(path))
                effective_subsample = subsample
                try:
                    if image.width() <= 400:
                        effective_subsample = max(1, min(subsample, 4))
                except Exception:
                    effective_subsample = subsample
                if effective_subsample > 1:
                    image = image.subsample(effective_subsample, effective_subsample)
                return image
            except Exception:
                continue
    return None


def wait_for_any_selector(page: Page, selectors: list[str], timeout: int = 15000) -> str:
    deadline = time.time() + (timeout / 1000.0)
    while time.time() < deadline:
        for sel in selectors:
            locator = page.locator(sel)
            try:
                if locator.count() > 0 and locator.first.is_visible():
                    return sel
            except Exception:
                pass
        time.sleep(0.2)
    raise TimeoutError(f"Nenhum seletor encontrado: {selectors}")


def click_any(page: Page, selectors: list[str], timeout: int = 15000) -> None:
    if not selectors:
        log("  [SKIP] Seletores vazios, pulando click")
        return
    log(f"  [CLICK] Buscando seletores: {selectors}")
    sel = wait_for_any_selector(page, selectors, timeout=timeout)
    log(f"  [CLICK] Encontrado seletor: '{sel}'")
    page.locator(sel).first.click()
    log(f"  [CLICK] Clicado com sucesso")


def fill_any(page: Page, selectors: list[str], value: str, timeout: int = 15000) -> None:
    if not selectors:
        log("  [SKIP] Seletores vazios, pulando preenchimento")
        return
    log(f"  [FILL] Buscando seletores: {selectors}")
    sel = wait_for_any_selector(page, selectors, timeout=timeout)
    log(f"  [FILL] Encontrado seletor: '{sel}', preenchendo com: '{value}'")
    target = page.locator(sel).first
    target.click()
    target.fill(value)
    log(f"  [FILL] Preenchido com sucesso")


def click_by_text(page: Page, text_value: str, timeout: int = 10000) -> None:
    locator = page.get_by_text(text_value, exact=True)
    locator.first.wait_for(state="visible", timeout=timeout)
    locator.first.click()


def choose_option_from_open_dropdown(page: Page, option_text: str, selectors_cfg: Dict[str, Any]) -> None:
    log(f"  [BUSCAR] Procurando opção: '{option_text}'")
    option_selectors = selectors_cfg["generic"]["dropdown_option_candidates"]
    
    for sel in option_selectors:
        try:
            locator = page.locator(sel).filter(has_text=option_text)
            count = locator.count()
            if count > 0:
                log(f"  [ENCONTRADO] Seletor '{sel}' com texto '{option_text}' ({count} elementos)")
                locator.first.click()
                log(f"  [CLICADO] Opção '{option_text}'")
                return
        except Exception as e:
            pass
    
    # Also check aria-label
    for sel in option_selectors:
        try:
            locator = page.locator(sel).locator(f"[aria-label*='{option_text}']")
            count = locator.count()
            if count > 0:
                log(f"  [ENCONTRADO] Seletor '{sel}' com aria-label '{option_text}' ({count} elementos)")
                locator.first.click()
                log(f"  [CLICADO] Opção '{option_text}'")
                return
        except Exception as e:
            pass

    try:
        page.get_by_text(option_text, exact=True).first.click()
        log(f"  [CLICADO] Opção '{option_text}' (get_by_text)")
        return
    except Exception as e:
        # Log detalhado de erro
        log(f"  [ERRO] Não conseguiu encontrar opção '{option_text}'")
        log(f"  [DEBUG] Seletores testados: {option_selectors}")
        raise RuntimeError(f"Não foi possível selecionar a opção '{option_text}'.") from e


def try_close_popups(page: Page, selectors_cfg: Dict[str, Any]) -> None:
    popup_selectors = selectors_cfg.get("popup_close_candidates", [])
    for sel in popup_selectors:
        try:
            locator = page.locator(sel)
            if locator.count() > 0 and locator.first.is_visible():
                locator.first.click()
                time.sleep(0.4)
        except Exception:
            continue


def clear_login_session() -> None:
    if STORAGE_STATE_PATH.exists():
        STORAGE_STATE_PATH.unlink()
        log(f"Sessão antiga removida: {STORAGE_STATE_PATH}")
    if LOGGED_USER_CACHE_PATH.exists():
        LOGGED_USER_CACHE_PATH.unlink()
        log(f"Cache de usuário removido: {LOGGED_USER_CACHE_PATH}")
    if PROFILE_DIR.exists():
        try:
            shutil.rmtree(PROFILE_DIR)
            log(f"Perfil persistente removido: {PROFILE_DIR}")
        except Exception as e:
            log(f"Falha ao remover perfil persistente: {e}")
    PROFILE_DIR.mkdir(exist_ok=True)


def _is_email(value: str) -> bool:
    email_pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    try:
        return bool(re.match(email_pattern, str(value).strip())) and len(value) < 100
    except Exception:
        return False


def _is_disallowed_logged_user_email(value: str) -> bool:
    email = str(value).strip().lower()
    if not email:
        return True
    bookings_resource_email = get_bookings_resource_email()
    if bookings_resource_email and email == bookings_resource_email:
        return True
    local_part = email.split("@")[0]
    if "agendamento" in local_part and "sesisenaisp.onmicrosoft.com" in email:
        return True
    return False


def _extract_first_email(text: str) -> Optional[str]:
    if not text:
        return None
    matches = re.findall(r'[\w\.-]+@[\w\.-]+\.\w+', str(text))
    for match in matches:
        if _is_email(match) and not _is_disallowed_logged_user_email(match):
            return match.strip()
    return None


def _extract_logged_user_from_state(state: Dict[str, Any]) -> str:
    # Buscar informações explícitas em localStorage
    for origin in state.get("origins", []):
        for item in origin.get("localStorage", []):
            name = str(item.get("name", "")).lower()
            value = str(item.get("value", ""))
            if any(key in name for key in ["email", "login_hint", "preferred_username", "upn"]):
                if _is_email(value) and not _is_disallowed_logged_user_email(value):
                    return value.strip()

    # Fallback em cookies
    for cookie in state.get("cookies", []):
        name = str(cookie.get("name", "")).lower()
        value = str(cookie.get("value", ""))
        if any(key in name for key in ["email", "login", "upn"]):
            if _is_email(value) and not _is_disallowed_logged_user_email(value):
                return value.strip()

    # Fallback final: procurar qualquer email no localStorage
    for origin in state.get("origins", []):
        for item in origin.get("localStorage", []):
            value = str(item.get("value", ""))
            if _is_email(value) and not _is_disallowed_logged_user_email(value):
                return value.strip()
    return "Usuário desconhecido"


def _extract_logged_user_from_page(page: Page) -> str:
    candidates: List[str] = []

    try:
        candidates.append(page.url)
    except Exception:
        pass

    try:
        candidates.append(page.title())
    except Exception:
        pass

    try:
        body_text = page.locator("body").inner_text(timeout=3000)
        candidates.append(body_text)
    except Exception:
        pass

    try:
        page_data = page.evaluate(
            """() => {
                const attrs = [];
                for (const el of Array.from(document.querySelectorAll('*')).slice(0, 400)) {
                    for (const attr of ['aria-label', 'title', 'value', 'placeholder', 'data-testid']) {
                        const val = el.getAttribute && el.getAttribute(attr);
                        if (val) attrs.push(val);
                    }
                    if (el.textContent && el.textContent.includes('@')) attrs.push(el.textContent);
                }
                const local = [];
                for (let i = 0; i < localStorage.length; i++) {
                    const key = localStorage.key(i);
                    local.push(`${key}=${localStorage.getItem(key)}`);
                }
                const session = [];
                for (let i = 0; i < sessionStorage.length; i++) {
                    const key = sessionStorage.key(i);
                    session.push(`${key}=${sessionStorage.getItem(key)}`);
                }
                return {
                    attrs: attrs.join('\\n'),
                    local: local.join('\\n'),
                    session: session.join('\\n'),
                    html: document.documentElement ? document.documentElement.innerHTML : ''
                };
            }"""
        )
        candidates.extend([
            str(page_data.get("attrs", "")),
            str(page_data.get("local", "")),
            str(page_data.get("session", "")),
            str(page_data.get("html", "")),
        ])
    except Exception:
        pass

    for text in candidates:
        email = _extract_first_email(text)
        if email:
            return email

    return "Usuário desconhecido"


def cache_logged_user(user: str) -> bool:
    if user in {"", "Nenhum usuário logado", "Usuário desconhecido"}:
        return False
    if _is_disallowed_logged_user_email(user):
        return False
    try:
        LOGGED_USER_CACHE_PATH.write_text(user.strip(), encoding="utf-8")
        return True
    except Exception:
        return False


def detect_logged_user(page: Optional[Page] = None, state: Optional[Dict[str, Any]] = None) -> str:
    user = "Usuário desconhecido"
    if page is not None:
        try:
            user = _extract_logged_user_from_page(page)
        except Exception:
            user = "Usuário desconhecido"
    if user in {"Nenhum usuário logado", "Usuário desconhecido"} and state is not None:
        try:
            user = _extract_logged_user_from_state(state)
        except Exception:
            user = "Usuário desconhecido"
    cache_logged_user(user)
    return user


def get_logged_user_from_storage() -> str:
    if LOGGED_USER_CACHE_PATH.exists():
        try:
            cached_user = LOGGED_USER_CACHE_PATH.read_text(encoding="utf-8").strip()
            if cached_user and not _is_disallowed_logged_user_email(cached_user):
                return cached_user
        except Exception:
            pass

    if not STORAGE_STATE_PATH.exists():
        return "Nenhum usuário logado"
    try:
        state = load_json(STORAGE_STATE_PATH)
    except Exception:
        return "Usuário desconhecido"

    user = _extract_logged_user_from_state(state)
    if user not in {"Nenhum usuário logado", "Usuário desconhecido"} and not _is_disallowed_logged_user_email(user):
        try:
            LOGGED_USER_CACHE_PATH.write_text(user, encoding="utf-8")
        except Exception:
            pass
    elif LOGGED_USER_CACHE_PATH.exists():
        try:
            LOGGED_USER_CACHE_PATH.unlink()
        except Exception:
            pass
    return user


def has_authenticated_session() -> bool:
    if STORAGE_STATE_PATH.exists():
        try:
            state = load_json(STORAGE_STATE_PATH)
            if state.get("cookies") or state.get("origins"):
                return True
        except Exception:
            pass
    try:
        if PROFILE_DIR.exists():
            for item in PROFILE_DIR.iterdir():
                return True
    except Exception:
        pass
    return False


def get_logged_user_display() -> str:
    user = get_logged_user_from_storage()
    if user not in {"Nenhum usuário logado", "Usuário desconhecido", ""}:
        return user
    if has_authenticated_session():
        return "Sessão ativa (email não identificado)"
    return "Nenhum usuário logado"


def _get_active_page(context) -> Optional[Page]:
    try:
        if context.pages:
            for candidate in reversed(context.pages):
                if not candidate.is_closed():
                    return candidate
    except Exception:
        pass
    return None


def save_login_session(selectors_cfg: Dict[str, Any], confirm_login_callback=None, status_callback=None) -> None:
    bookings_url = require_bookings_url()
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            viewport={"width": 1400, "height": 980},
        )
        page = context.new_page()
        log("Abrindo Bookings para login manual...")
        page.goto(bookings_url, wait_until="domcontentloaded")
        log("Faça login manualmente no navegador.")
        log("Quando voltar para a página do Bookings, confirme na janela do programa.")

        if status_callback:
            status_callback("Aguardando conclusão do login...")

        login_confirmed = False
        wait_deadline = time.time() + 300
        stable_ready_count = 0
        confirmation_requested = False

        while time.time() < wait_deadline:
            active_page = _get_active_page(context)
            detected_user = "Usuário desconhecido"
            current_url = ""
            if active_page is not None:
                try:
                    current_url = active_page.url
                except Exception:
                    current_url = ""
                detected_user = detect_logged_user(page=active_page)

            back_on_bookings = "outlook.office.com/book/" in current_url.lower()
            valid_user = detected_user not in {"Nenhum usuário logado", "Usuário desconhecido"}

            if back_on_bookings or valid_user:
                stable_ready_count += 1
                if status_callback and stable_ready_count == 1:
                    status_callback("Login detectado. Aguardando estabilização da página...")
            else:
                stable_ready_count = 0

            if stable_ready_count >= 3 and not confirmation_requested:
                confirmation_requested = True
                if confirm_login_callback:
                    display_user = detected_user if valid_user else "usuário autenticado"
                    if status_callback:
                        status_callback("Confirme o login na janela do programa.")
                    login_confirmed = bool(confirm_login_callback(display_user))
                else:
                    login_confirmed = True
                if login_confirmed:
                    break
                confirmation_requested = False
                stable_ready_count = 0
            time.sleep(1)

        if not login_confirmed:
            context.close()
            raise TimeoutError("Tempo esgotado aguardando a confirmação do login.")

        try:
            state = context.storage_state(path=str(STORAGE_STATE_PATH))
            log(f"Sessão salva em: {STORAGE_STATE_PATH}")
            active_page = _get_active_page(context)
            user = detect_logged_user(page=active_page, state=state)
            if user not in {"Nenhum usuário logado", "Usuário desconhecido"}:
                log(f"Usuário logado identificado: {user}")
            else:
                warn("Não foi possível identificar o email do usuário logado automaticamente.")
        finally:
            context.close()


def new_authenticated_context(p, headless: bool = False) -> BrowserContext:
    if STORAGE_STATE_PATH.exists():
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(
            storage_state=str(STORAGE_STATE_PATH),
            viewport={"width": 1400, "height": 980},
        )
        return context

    log("storage_state.json não encontrado. Usando perfil persistente salvo.")
    context = p.chromium.launch_persistent_context(
        user_data_dir=str(PROFILE_DIR),
        headless=headless,
        viewport={"width": 1400, "height": 980},
    )
    return context


def open_bookings_page(page: Page) -> None:
    bookings_url = require_bookings_url()
    log("Abrindo página do Bookings...")
    page.goto(bookings_url, wait_until="domcontentloaded")
    page.wait_for_timeout(2000)


def fill_step_escolha_reserva(page: Page, job: BookingJob, cfg: Dict[str, Any]) -> None:
    log(f"Preenchendo reserva: {job.escolha_reserva}")
    # Tentar clicar se houver seletores
    if cfg["fields"]["escolha_reserva_open"]:
        click_any(page, cfg["fields"]["escolha_reserva_open"])
    
    page.wait_for_timeout(1000)
    
    # Procurar pela opção "50min" diretamente
    try:
        # Procurar em li elementos com radio buttons
        li_elements = page.locator("li").all()
        found = False
        
        if len(li_elements) > 0:
            log(f"  [DEBUG] Encontrados {len(li_elements)} elementos <li>")
            log(f"  [DEBUG] Listando primeiros 10 <li> com textos:")
            
            for idx, li in enumerate(li_elements[:10]):
                try:
                    text = li.text_content().strip()
                    if idx < 10:
                        log(f"    [{idx}] '{text}'")
                    
                    # Procurar por "50min" ou o valor procurado
                    if job.escolha_reserva in text:
                        log(f"  [ENCONTRADO] '{job.escolha_reserva}' no elemento {idx}")
                        li.click()
                        log(f"  [CLICADO] Opção '{job.escolha_reserva}'")
                        found = True
                        break
                except:
                    pass
        
        if not found:
            # Tentar método genérico
            log(f"  [FALLBACK] Usando choose_option_from_open_dropdown")
            choose_option_from_open_dropdown(page, job.escolha_reserva, cfg)
    except Exception as e:
        log(f"  [ERRO] {e}")
        raise
    
    page.wait_for_timeout(2000)


def fill_step_dia_semana(page: Page, job: BookingJob, cfg: Dict[str, Any]) -> None:
    if not cfg["fields"]["dia_semana_open"]:
        log("Pulando seleção de dia da semana (não disponível na página)")
        return
    log(f"Preenchendo dia(s) da semana: {job.dias_semana}")
    click_any(page, cfg["fields"]["dia_semana_open"])
    for dia in job.dias_semana:
        choose_option_from_open_dropdown(page, dia, cfg)
    if cfg["fields"].get("dia_semana_close_after_select"):
        click_any(page, cfg["fields"]["dia_semana_close_after_select"])


def fill_step_equipe(page: Page, job: BookingJob, cfg: Dict[str, Any]) -> None:
    if not cfg["fields"]["equipe_open"]:
        log("Pulando seleção de equipe (não disponível na página)")
        return
    log(f"Selecionando equipe: {job.equipe}")
    for pessoa in job.equipe:
        log(f"  Abrindo dropdown de equipe...")
        click_any(page, cfg["fields"]["equipe_open"])
        page.wait_for_timeout(1000)
        log(f"  Selecionando '{pessoa}'...")
        choose_option_from_open_dropdown(page, pessoa, cfg)
        page.wait_for_timeout(1500)
    if cfg["fields"].get("equipe_close_after_select"):
        click_any(page, cfg["fields"]["equipe_close_after_select"])


def fill_step_horario(page: Page, job: BookingJob, cfg: Dict[str, Any]) -> None:
    log(f"Selecionando horário: {job.horario}")
    page.wait_for_timeout(2000)
    # Procurar no div do selecionador de horário
    log(f"  Aguardando que horários estejam disponíveis...")
    page.wait_for_selector("div[role='group'] ul[role='list'] li label span", timeout=10000)
    log(f"  Horários carregados, procurando por {job.horario}...")
    try:
        # Procurar por todos os spans de horário
        all_time_spans = page.locator("div[role='group'] ul[role='list'] li label span").all()
        found_span = None
        
        # Normalizar horário para comparação (remover leading zeros)
        horario_normalizado = job.horario.lstrip('0') if job.horario.startswith('0') else job.horario
        
        for span in all_time_spans:
            text = span.text_content().strip()
            # Comparar com e sem leading zeros
            if text == job.horario or text == horario_normalizado:
                found_span = span
                break
        
        if found_span:
            log(f"  [ENCONTRADO] Horário {job.horario}")
            # Clicar no span para selecionar
            found_span.click()
            page.wait_for_timeout(500)
            log(f"  [CLICADO] Horário {job.horario}")
        else:
            log(f"  [ERRO] Horário {job.horario} não encontrado com match exato")
            log(f"  Horários disponíveis na página:")
            # Log dos horários disponíveis para debug
            for idx, t in enumerate(all_time_spans):
                try:
                    text = t.text_content().strip()
                    if idx < 25:  # Mostrar os primeiros 25
                        log(f"    - '{text}'")
                except:
                    pass
            raise RuntimeError(f"Horário '{job.horario}' não encontrado com correspondência exata")
    except Exception as e:
        raise RuntimeError(f"Erro ao selecionar horário '{job.horario}': {e}") from e


def fill_step_notas(page: Page, job: BookingJob, cfg: Dict[str, Any]) -> None:
    log("Preenchendo notas")
    fill_any(page, cfg["fields"]["notas_input"], job.notas)


def fill_step_publico(page: Page, job: BookingJob, cfg: Dict[str, Any]) -> None:
    log(f"Selecionando público: {job.publico}")
    _select_valor_by_label(page, "PÚBLICO", job.publico, cfg)


def _clean_option_text(value: str) -> str:
    text = str(value or "").replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text).strip(" \t\r\n,;")
    return text


def _dedupe_option_values(values: list[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    ignored = {
        "",
        "selecione",
        "selecione...",
        "select",
        "select...",
        "--selecione uma opção--",
        "-- selecione uma opção --",
        "selecione uma opção",
    }
    for value in values:
        cleaned = _clean_option_text(value)
        if not cleaned or cleaned.casefold() in ignored:
            continue
        key = cleaned.casefold()
        if key not in seen:
            seen.add(key)
            result.append(cleaned)
    return result


def _extract_option_text_from_locator(locator) -> str:
    for getter in (
        lambda: locator.inner_text(timeout=1000),
        lambda: locator.text_content(timeout=1000),
        lambda: locator.get_attribute("aria-label"),
        lambda: locator.get_attribute("value"),
        lambda: locator.get_attribute("title"),
    ):
        try:
            value = getter()
        except Exception:
            continue
        cleaned = _clean_option_text(value or "")
        if cleaned:
            return cleaned
    return ""


def _collect_visible_texts(page: Page, selectors: list[str], limit_per_selector: int = 120) -> list[str]:
    collected: list[str] = []
    for sel in selectors:
        try:
            locator = page.locator(sel)
            count = min(locator.count(), limit_per_selector)
        except Exception:
            continue
        for idx in range(count):
            try:
                item = locator.nth(idx)
                if not item.is_visible():
                    continue
                text = _extract_option_text_from_locator(item)
                if text:
                    collected.append(text)
            except Exception:
                continue
    return _dedupe_option_values(collected)


def _click_first_matching_visible_option(page: Page, option_text: str, selectors: list[str]) -> bool:
    target = _clean_option_text(option_text)
    if not target:
        return False
    target_lower = target.casefold()
    for sel in selectors:
        try:
            locator = page.locator(sel)
            count = min(locator.count(), 120)
        except Exception:
            continue
        for idx in range(count):
            try:
                item = locator.nth(idx)
                if not item.is_visible():
                    continue
                text = _extract_option_text_from_locator(item)
                if not text:
                    continue
                if text.casefold() == target_lower or target_lower in text.casefold():
                    item.click()
                    return True
            except Exception:
                continue
    try:
        page.get_by_text(option_text, exact=True).first.click(timeout=2000)
        return True
    except Exception:
        return False


def _extract_reservation_options(page: Page) -> list[str]:
    source_texts: list[str] = []
    try:
        source_texts.append(page.locator("body").inner_text(timeout=3000))
    except Exception:
        pass
    source_texts.extend(_collect_visible_texts(page, ["li", "button", "label", "span"], limit_per_selector=80))

    matches: list[str] = []
    for text in source_texts:
        matches.extend(re.findall(r"\b\d+h\d+min\b|\b\d+min\b", str(text), flags=re.IGNORECASE))
    return _dedupe_option_values(matches)


def _extract_time_options(page: Page) -> list[str]:
    time_selectors = [
        "div[role='group'] ul[role='list'] li label span",
        "ul[role='list'] li label span",
        "ul[role='list'] li span",
    ]
    candidates = _collect_visible_texts(page, time_selectors, limit_per_selector=240)
    times = [value for value in candidates if re.fullmatch(r"\d{1,2}:\d{2}", value)]
    return _dedupe_option_values(times)


def _merge_horario_options(existing: list[str], imported: list[str]) -> list[str]:
    combined = _dedupe_option_values(list(existing or []) + list(imported or []))

    def time_key(value: str) -> tuple[int, int]:
        try:
            hour, minute = value.split(":")
            return int(hour), int(minute)
        except Exception:
            return (99, 99)

    return sorted(combined, key=time_key)


def _extract_equipe_options(page: Page, selectors_cfg: Dict[str, Any]) -> list[str]:
    open_selectors = selectors_cfg.get("fields", {}).get("equipe_open", [])
    if not open_selectors:
        return []
    option_selectors = list(selectors_cfg.get("generic", {}).get("dropdown_option_candidates", []))
    option_selectors.extend([
        "[role='listbox'] [role='option']",
        "[role='option']",
        "li",
    ])
    click_any(page, open_selectors)
    page.wait_for_timeout(1200)
    try:
        options = _collect_visible_texts(page, option_selectors, limit_per_selector=180)
    finally:
        try:
            page.keyboard.press("Escape")
            page.wait_for_timeout(300)
        except Exception:
            pass
    filtered = [
        value
        for value in options
        if "membro da equipe" not in value.casefold()
        and "selecionar equipe" not in value.casefold()
    ]
    return _dedupe_option_values(filtered)


def import_field_options_from_bookings(selectors_cfg: Dict[str, Any]) -> Dict[str, list[str]]:
    imported: Dict[str, list[str]] = {}
    with sync_playwright() as p:
        context = new_authenticated_context(p, headless=False)
        try:
            page = context.new_page()
            open_bookings_page(page)
            try_close_popups(page, selectors_cfg)
            page.wait_for_timeout(1800)

            reserva_options = _extract_reservation_options(page)
            if reserva_options:
                imported["escolha_reserva"] = reserva_options
                log(f"Opções importadas para escolha_reserva: {reserva_options}")

            selected_reserva = reserva_options[0] if reserva_options else ""
            if selected_reserva:
                clicked = _click_first_matching_visible_option(page, selected_reserva, ["li", "button", "label", "span", "div"])
                if clicked:
                    log(f"Reserva de referência selecionada para leitura dos campos: {selected_reserva}")
                    page.wait_for_timeout(2000)
                    try_close_popups(page, selectors_cfg)

            equipe_options = []
            try:
                equipe_options = _extract_equipe_options(page, selectors_cfg)
            except Exception as exc:
                warn(f"Não foi possível importar as opções de equipe automaticamente: {exc}")
            if equipe_options:
                imported["equipe"] = equipe_options
                log(f"Opções importadas para equipe: {len(equipe_options)} itens")

            select_field_labels = [
                ("COMPONENTE", "componente"),
                ("PÚBLICO", "publico"),
                ("TURMA", "turma"),
                ("PRINCIPAL RECURSO", "principal_recurso"),
                ("TIPO DE ATIVIDADE", "tipo_atividade"),
            ]
            for label_keyword, field_key in select_field_labels:
                values = get_select_options(page, label_keyword)
                if values:
                    imported[field_key] = values
                    log(f"Opções importadas para {field_key}: {len(values)} itens")

            horario_options = _extract_time_options(page)
            if horario_options:
                imported["horario"] = horario_options
                log(f"Opções importadas para horario: {horario_options}")

            try:
                context.storage_state(path=str(STORAGE_STATE_PATH))
            except Exception:
                pass
        finally:
            if hasattr(context, "browser") and context.browser:
                try:
                    context.browser.close()
                except Exception:
                    pass
            else:
                try:
                    context.close()
                except Exception:
                    pass

    imported = {field: _dedupe_option_values(values) for field, values in imported.items() if values}
    if not imported:
        raise RuntimeError(
            "NÃ£o foi possÃ­vel ler os campos do Bookings automaticamente. Verifique se o link da unidade estÃ¡ correto e se o usuÃ¡rio estÃ¡ logado."
        )
    return imported


def get_select_options(page: Page, label_keyword: str) -> list[str]:
    """Extrai opções de um select pelo label."""
    select_elements = page.locator("select")
    total_selects = select_elements.count()
    
    for idx in range(total_selects):
        try:
            select = select_elements.nth(idx)
            select_id = select.get_attribute("id")
            
            if select_id:
                label = page.locator(f"label[for='{select_id}']")
                if label.count() > 0:
                    label_text = label.text_content().upper()
                    if label_keyword.upper() in label_text:
                        options = select.locator("option").all()
                        raw_values = []
                        for opt in options:
                            try:
                                text = (opt.text_content() or "").strip()
                            except Exception:
                                text = ""
                            if text:
                                raw_values.append(text)
                        return _dedupe_option_values(raw_values)
        except:
            pass
    return []


def _select_valor_by_label(page: Page, label_keyword: str, valor: str, cfg: Dict[str, Any]) -> None:
    """Função auxiliar para selecionar valor em SELECT pelo label."""
    select_elements = page.locator("select")
    total_selects = select_elements.count()
    
    def normalize(text: str) -> str:
        return text.strip().lower() if text else ""

    def option_matches(option_text: str, option_value: str, target: str) -> bool:
        option_text_n = normalize(option_text)
        option_value_n = normalize(option_value)
        target_n = normalize(target)
        if option_text_n == target_n or option_value_n == target_n:
            return True
        if target_n in option_text_n or option_text_n in target_n:
            return True
        return False

    found_select = None
    for idx in range(total_selects):
        try:
            select = select_elements.nth(idx)
            select_id = select.get_attribute("id")
            
            if select_id:
                label = page.locator(f"label[for='{select_id}']")
                if label.count() > 0:
                    label_text = label.text_content().upper()
                    if label_keyword.upper() in label_text:
                        found_select = select
                        log(f"  [ENCONTRADO] SELECT com label '{label_keyword}' no índice {idx}")
                        break
            # fallback por aria-label direto no select
            aria_label = select.get_attribute("aria-label")
            if aria_label and label_keyword.upper() in aria_label.upper():
                found_select = select
                log(f"  [ENCONTRADO] SELECT com aria-label '{aria_label}' no índice {idx}")
                break
        except Exception:
            pass
    
    if not found_select:
        log(f"  [AVISO] SELECT com label '{label_keyword}' não encontrado")
        return

    try:
        found_select.wait_for(state="visible", timeout=10000)
        page.wait_for_timeout(500)  # Aguardar renderização completa
        
        options = found_select.locator("option").all()
        
        # Logar todas as opções disponíveis para debug
        log(f"  [DEBUG] Opções do SELECT '{label_keyword}':")
        all_option_info = []
        for i, opt in enumerate(options):
            text = opt.text_content() or ""
            value = opt.get_attribute("value") or ""
            all_option_info.append((text, value))
            if i < 10:
                log(f"    [{i}] text='{text}' value='{value}'")
        
        # Procurar pela opção
        candidate_by_value = None
        candidate_by_label = None
        
        for option in options:
            text = option.text_content() or ""
            value = option.get_attribute("value") or ""
            
            if option_matches(text, value, valor):
                if normalize(value) == normalize(valor):
                    candidate_by_value = value
                    break
                elif candidate_by_label is None:
                    candidate_by_label = text or value
        
        # Primeiro tenta por valor (mais confiável)
        if candidate_by_value:
            log(f"  [SELEÇÃO] Tentando por valor: '{candidate_by_value}'")
            try:
                found_select.select_option(candidate_by_value)
                log(f"  [SELECIONADO] '{valor}' com sucesso por valor")
                page.wait_for_timeout(500)
                return
            except Exception as e:
                log(f"  [RETRY] Falha ao selecionar por valor: {e}")
        
        # Depois tenta por label
        if candidate_by_label:
            log(f"  [SELEÇÃO] Tentando por label: '{candidate_by_label}'")
            try:
                found_select.select_option(candidate_by_label)
                log(f"  [SELECIONADO] '{valor}' com sucesso por label")
                page.wait_for_timeout(500)
                return
            except Exception as e:
                log(f"  [RETRY] Falha ao selecionar por label: {e}")
        
        # Fallback: tentar abrir dropdown e clicar
        if candidate_by_value or candidate_by_label:
            log(f"  [FALLBACK] Tentando abrir dropdown e selecionar manualmente")
            try:
                found_select.click()
                page.wait_for_timeout(300)
                choose_option_from_open_dropdown(page, valor, cfg)
                return
            except Exception as e:
                log(f"  [RETRY] Falha no fallback: {e}")
        
        log(f"  [ERRO] Opção '{valor}' não encontrada")
        log(f"  [ERRO] Opções disponíveis: {[(t, v) for t, v in all_option_info[:15]]}")
        raise RuntimeError(f"Opção '{valor}' não encontrada no select '{label_keyword}'.")
        
    except Exception as e:
        log(f"  [ERRO] Ao selecionar '{valor}': {e}")
        raise


def fill_step_componente(page: Page, job: BookingJob, cfg: Dict[str, Any]) -> None:
    log(f"Selecionando componente: {job.componente}")
    _select_valor_by_label(page, "COMPONENTE", job.componente, cfg)


def fill_step_turma(page: Page, turma: str, cfg: Dict[str, Any]) -> None:
    log(f"Selecionando turma: {turma}")
    _select_valor_by_label(page, "TURMA", turma, cfg)


def fill_step_principal_recurso(page: Page, job: BookingJob, cfg: Dict[str, Any]) -> None:
    log(f"Selecionando principal recurso: {job.principal_recurso}")
    _select_valor_by_label(page, "PRINCIPAL RECURSO", job.principal_recurso, cfg)


def fill_step_tipo_atividade(page: Page, job: BookingJob, cfg: Dict[str, Any]) -> None:
    log(f"Selecionando tipo de atividade: {job.tipo_atividade}")
    _select_valor_by_label(page, "TIPO DE ATIVIDADE", job.tipo_atividade, cfg)


def fill_step_data(page: Page, date_value: str, cfg: Dict[str, Any]) -> None:
    log(f"Definindo data: {date_value}")
    import datetime as dt
    date_obj = dt.datetime.strptime(date_value, "%d/%m/%Y").date()
    iso = date_obj.isoformat() + "T00:00:00.000Z"
    log(f"  Data em formato ISO: {iso}")
    selector = f"div[data-value='{iso}']"
    log(f"  Procurando seletor: {selector}")
    click_any(page, [selector])
    page.wait_for_timeout(5000)
    log(f"  Data selecionada com sucesso")


def click_enviar(page: Page, cfg: Dict[str, Any], envio_real: bool) -> None:
    if not envio_real:
        warn("Modo teste: envio real desativado. O formulário será preenchido, mas não será enviado.")
        return
    log("Clicando em enviar...")
    click_any(page, cfg["fields"]["botao_enviar"])


def click_novo_agendamento(page: Page, cfg: Dict[str, Any]) -> None:
    selectors = cfg["fields"].get("novo_agendamento")
    if not selectors:
        return
    log("Abrindo novo agendamento...")
    click_any(page, selectors)
    page.wait_for_timeout(1500)


def wait_after_submit(page: Page, cfg: Dict[str, Any]) -> None:
    timeout_ms = cfg["generic"].get("after_submit_wait_ms", 3000)
    page.wait_for_timeout(timeout_ms)


def fill_single_booking(page: Page, job: BookingJob, turma_dict: dict, data_ref: str, cfg: Dict[str, Any], root: tk.Tk, stop_event: Optional[threading.Event] = None) -> None:
    turma = turma_dict["turma"]
    horario = turma_dict["horario"]
    dias_semana = turma_dict["dias_semana"]
    log(f"Preenchendo formulário para turma {turma} em {data_ref}")
    
    if stop_event and stop_event.is_set():
        raise InterruptedError("Parada solicitada antes de iniciar preenchimento")
    
    try_close_popups(page, cfg)
    
    if stop_event and stop_event.is_set():
        raise InterruptedError("Parada solicitada após fechar popups")
    
    log("1. Preenchendo escolha de reserva...")
    fill_step_escolha_reserva(page, job, cfg)
    page.wait_for_timeout(500)
    
    if stop_event and stop_event.is_set():
        raise InterruptedError("Parada solicitada durante preenchimento")
    
    log("2. Preenchendo equipe...")
    fill_step_equipe(page, job, cfg)
    page.wait_for_timeout(1000)
    
    if stop_event and stop_event.is_set():
        raise InterruptedError("Parada solicitada durante preenchimento")
    
    log("3. Preenchendo data...")
    fill_step_data(page, data_ref, cfg)
    
    if stop_event and stop_event.is_set():
        raise InterruptedError("Parada solicitada durante preenchimento")
    
    log("4. Preenchendo horário...")
    # Temporarily modify job.horario for fill_step_horario
    original_horario = job.horario
    job.horario = horario
    try:
        fill_step_horario(page, job, cfg)
    finally:
        job.horario = original_horario
    
    if stop_event and stop_event.is_set():
        raise InterruptedError("Parada solicitada durante preenchimento")
    
    log("5. Preenchendo dia da semana...")
    # Temporarily modify job.dias_semana for fill_step_dia_semana
    original_dias = job.dias_semana
    job.dias_semana = dias_semana
    try:
        fill_step_dia_semana(page, job, cfg)
    finally:
        job.dias_semana = original_dias
    
    if stop_event and stop_event.is_set():
        raise InterruptedError("Parada solicitada durante preenchimento")
    
    log("6. Preenchendo notas...")
    fill_step_notas(page, job, cfg)
    
    if stop_event and stop_event.is_set():
        raise InterruptedError("Parada solicitada durante preenchimento")
    
    log("7. Preenchendo componente...")
    fill_step_componente(page, job, cfg)
    
    if stop_event and stop_event.is_set():
        raise InterruptedError("Parada solicitada durante preenchimento")
    
    log("8. Preenchendo público...")
    fill_step_publico(page, job, cfg)
    
    if stop_event and stop_event.is_set():
        raise InterruptedError("Parada solicitada durante preenchimento")
    
    log("9. Preenchendo turma...")
    fill_step_turma(page, turma, cfg)
    log("10. Preenchendo principal recurso...")
    fill_step_principal_recurso(page, job, cfg)
    log("11. Preenchendo tipo de atividade...")
    fill_step_tipo_atividade(page, job, cfg)
    log("12. Clicando em enviar...")
    click_enviar(page, cfg, envio_real=job.confirmar_envio_real)
    log("13. Aguardando envio...")
    wait_after_submit(page, cfg)
    safe_add_result(f"Turma {turma}: Agendamento realizado com sucesso em {data_ref}", root)
    log(f"✓ Formulário preenchido com sucesso para turma {turma}")


def daterange_by_weekday(start_date: str, end_date: str, weekdays: list[str]) -> list[str]:
    weekday_map = {
        "Segunda": 0,
        "Terça": 1,
        "Quarta": 2,
        "Quinta": 3,
        "Sexta": 4,
        "Sábado": 5,
        "Domingo": 6,
    }
    wd = {weekday_map[d] for d in weekdays}
    start = parse_br_date(start_date)
    end = parse_br_date(end_date)

    results = []
    current = start
    while current <= end:
        if current.weekday() in wd:
            results.append(current.strftime("%d/%m/%Y"))
        current += dt.timedelta(days=1)
    return results


def run_jobs(headless: bool = False, wait_for_enter: bool = True, stop_event: Optional[threading.Event] = None, root: Optional[tk.Tk] = None, on_user_detected=None) -> None:
    bookings_url = require_bookings_url()
    cfg = load_json(SELECTORS_PATH)
    jobs = load_jobs(JOBS_PATH)

    with sync_playwright() as p:
        context = new_authenticated_context(p, headless=headless)
        page = context.new_page()
        open_bookings_page(page)
        detected_user = detect_logged_user(page=page)
        if detected_user not in {"Nenhum usuário logado", "Usuário desconhecido"}:
            log(f"Usuário logado em uso: {detected_user}")
            if on_user_detected:
                try:
                    on_user_detected(detected_user)
                except Exception:
                    pass

        try:
            for idx, job in enumerate(jobs, start=1):
                if stop_event and stop_event.is_set():
                    log("Parada solicitada: interrompendo execução")
                    break
                log("=" * 70)
                log(f"Iniciando job {idx}")
                
                for turma_dict in job.turmas:
                    if stop_event and stop_event.is_set():
                        log("Parada solicitada: interrompendo execução")
                        break
                    turma = turma_dict["turma"]
                    dias_semana = turma_dict["dias_semana"]
                    horario = turma_dict["horario"]
                    log(f"Processando turma {turma} com dias {dias_semana} e horário {horario}")
                    datas = daterange_by_weekday(job.data_inicio, job.data_fim, dias_semana)
                    log(f"Datas encontradas para turma {turma}: {datas}")
                    if not datas:
                        log(f"Nenhuma data encontrada para os dias da semana especificados para turma {turma}. Pulando.")
                        continue

                    for data_ref in datas:
                        if stop_event and stop_event.is_set():
                            log("Parada solicitada: interrompendo execução")
                            break
                        if msvcrt.kbhit() and msvcrt.getch() == b'\x1b':
                            log("ESC pressionado: interrompendo execução")
                            break
                        log(f"Agendamento -> Data {data_ref} | Turma {turma}")
                        try:
                            fill_single_booking(page, job, turma_dict, data_ref, cfg, root, stop_event=stop_event)
                        except InterruptedError:
                            log("Execução interrompida pelo usuário")
                            raise
                        except Exception as e:
                            fail(f"Falha ao preencher agendamento da turma {turma} em {data_ref}: {e}")
                            safe_add_result(f"Turma {turma}: Falha em {data_ref} - {e}", root)
                        finally:
                            try:
                                click_novo_agendamento(page, cfg)
                            except Exception:
                                page.goto(bookings_url, wait_until="domcontentloaded")
                                page.wait_for_timeout(2000)
        except KeyboardInterrupt:
            log("Execução interrompida pelo usuário (Ctrl+C)")
        except InterruptedError:
            log("Execução parada pelo botão de parada")

        log("Execução concluída.")
        if wait_for_enter:
            log("Pressione ENTER para fechar.")
            input()

        try:
            context.storage_state(path=str(STORAGE_STATE_PATH))
        except Exception:
            pass

        if hasattr(context, "browser") and context.browser:
            try:
                context.browser.close()
            except Exception:
                pass
        else:
            try:
                context.close()
            except Exception:
                pass


def inspect_mode() -> None:
    with sync_playwright() as p:
        context = new_authenticated_context(p, headless=False)
        page = context.new_page()
        open_bookings_page(page)
        print("Modo inspeção: use o navegador para inspecionar elementos.")
        input("Pressione ENTER quando terminar.")
        try:
            context.storage_state(path=str(STORAGE_STATE_PATH))
        except Exception:
            pass
        if hasattr(context, "browser") and context.browser:
            context.browser.close()
        else:
            context.close()


def _normalize_list_field(text: str) -> list[str]:
    return [item.strip() for item in text.split(",") if item.strip()]


def _build_gui_job_data(entries: Dict[str, StringVar], selected_turmas: list[str], turma_details: Dict[str, Dict[str, StringVar]]) -> dict:
    turmas_list = []
    for turma in selected_turmas:
        details = turma_details.get(turma, {})
        dias = _normalize_list_field(details.get("dias_semana", StringVar(value="")).get())
        horario = details.get("horario", StringVar(value="")).get().strip()
        turmas_list.append({
            "turma": turma,
            "dias_semana": dias,
            "horario": horario
        })
    return {
        "escolha_reserva": entries["escolha_reserva"].get().strip(),
        "equipe": _normalize_list_field(entries["equipe"].get()),
        "dias_semana": [],
        "horario": "",
        "notas": entries["notas"].get().strip(),
        "componente": entries["componente"].get().strip(),
        "publico": entries["publico"].get().strip(),
        "turmas": turmas_list,
        "principal_recurso": entries["principal_recurso"].get().strip(),
        "tipo_atividade": entries["tipo_atividade"].get().strip(),
        "data_inicio": entries["data_inicio"].get().strip(),
        "data_fim": entries["data_fim"].get().strip(),
        "confirmar_envio_real": True,  # Will be set in caller
    }


def _run_jobs_with_job_data(job_data: dict, confirm_real: bool, stop_event: threading.Event, headless: bool = False, root: Optional[tk.Tk] = None, on_user_detected=None) -> None:
    job_data["confirmar_envio_real"] = confirm_real
    temp_path = BASE_DIR / "__tmp_jobs_gui__.json"
    jobs_data = {"jobs": [job_data]}
    temp_path.write_text(json.dumps(jobs_data, ensure_ascii=False, indent=2), encoding="utf-8")
    global JOBS_PATH
    original_jobs_path = JOBS_PATH
    JOBS_PATH = temp_path
    try:
        run_jobs(headless=headless, wait_for_enter=False, stop_event=stop_event, root=root, on_user_detected=on_user_detected)
    finally:
        JOBS_PATH = original_jobs_path
        if temp_path.exists():
            temp_path.unlink()


def launch_gui() -> None:
    global log_text_widget, results_text_widget
    root = tk.Tk()
    root.title(APP_NAME)
    root.geometry("1240x920")
    root.minsize(920, 680)
    root.resizable(True, True)
    root.configure(bg="#eef3f1")
    
    # Configurar layout principal com grid para responsividade
    root.rowconfigure(1, weight=1)
    root.rowconfigure(3, weight=0)
    root.columnconfigure(0, weight=1)

    selectors_cfg = load_json(SELECTORS_PATH)
    field_options = selectors_cfg.get("field_values", {})
    app_settings = load_app_settings()

    colors = {
        "bg": "#f6f7f6",
        "panel": "#fcfdfc",
        "panel_alt": "#ffffff",
        "border": "#d9dfdc",
        "text": "#1c2b28",
        "muted": "#6a7974",
        "accent": "#1f5f57",
        "accent_soft": "#e7f2ef",
        "accent_red": "#cb4655",
        "accent_red_soft": "#fff5f6",
        "success_bg": "#f3f8f6",
        "success_border": "#d7e6e0",
        "log_bg": "#f8faf9",
    }

    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass

    style.configure(".", background=colors["bg"], foreground=colors["text"], font=("Segoe UI", 10))
    style.configure("App.TFrame", background=colors["bg"])
    style.configure("Card.TFrame", background=colors["panel"], relief="solid", borderwidth=1)
    style.configure("Toolbar.TFrame", background=colors["panel_alt"], relief="solid", borderwidth=1)
    style.configure("Inline.TFrame", background=colors["panel"])
    style.configure("TLabel", background=colors["bg"], foreground=colors["text"], font=("Segoe UI", 10))
    style.configure("Title.TLabel", background=colors["panel_alt"], foreground=colors["text"], font=("Segoe UI Semibold", 15))
    style.configure("Subtitle.TLabel", background=colors["panel_alt"], foreground=colors["muted"], font=("Segoe UI", 10))
    style.configure("MutedPanel.TLabel", background=colors["panel"], foreground=colors["muted"], font=("Segoe UI", 9))
    style.configure("SecondarySection.TLabel", background=colors["panel"], foreground="#7c8b86", font=("Segoe UI Semibold", 10))
    style.configure("User.TLabel", background=colors["panel_alt"], foreground=colors["accent"], font=("Segoe UI Semibold", 10))
    style.configure("Section.TLabel", background=colors["panel"], foreground=colors["text"], font=("Segoe UI Semibold", 12))
    style.configure("Footer.TLabel", background=colors["bg"], foreground=colors["muted"], font=("Segoe UI", 9))
    style.configure("Status.TLabel", background=colors["panel_alt"], foreground=colors["accent"], font=("Segoe UI Semibold", 10))
    style.configure("TButton", font=("Segoe UI", 10), padding=(11, 6), background="#f8fbfa", foreground=colors["text"], bordercolor=colors["border"], relief="solid")
    style.map("TButton", background=[("active", "#eef4f2")], bordercolor=[("active", "#c4d3cd")])
    style.configure("Soft.TButton", font=("Segoe UI", 10), padding=(11, 6), background="#f7faf9", foreground=colors["text"], bordercolor=colors["border"])
    style.map("Soft.TButton", background=[("active", "#edf3f1")], bordercolor=[("active", "#c5d3cd")])
    style.configure("Primary.TButton", font=("Segoe UI Semibold", 10), padding=(12, 7), background=colors["accent"], foreground="#ffffff", bordercolor=colors["accent"])
    style.map("Primary.TButton", background=[("active", "#184e47")], foreground=[("active", "#ffffff")])
    style.configure("Danger.TButton", font=("Segoe UI Semibold", 10), padding=(12, 7), background="#fff7f7", foreground="#b24152", bordercolor="#eed2d7")
    style.map("Danger.TButton", background=[("active", "#fff0f2")], bordercolor=[("active", "#e2bcc4")])
    style.configure("TCheckbutton", background=colors["panel"], foreground=colors["text"], font=("Segoe UI", 10))
    style.configure("TEntry", fieldbackground="#ffffff", bordercolor=colors["border"], lightcolor=colors["border"], darkcolor=colors["border"], padding=7)
    style.configure("TCombobox", fieldbackground="#ffffff", bordercolor=colors["border"], lightcolor=colors["border"], darkcolor=colors["border"], arrowsize=14, padding=6)

    # Menu buttons at top
    menu_frame = ttk.Frame(root, padding=18, style="Toolbar.TFrame")
    menu_frame.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 8))
    menu_frame.columnconfigure(0, weight=3)
    menu_frame.columnconfigure(1, weight=2)
    menu_frame.columnconfigure(2, weight=0)
    menu_frame.columnconfigure(3, weight=0)

    logged_user_label = None
    status_label = None
    run_test_button = None
    run_real_button = None
    bookings_url_var = StringVar(value=app_settings.get("bookings_url", DEFAULT_BOOKINGS_URL))

    def current_logged_user() -> str:
        return get_logged_user_display()

    def has_valid_logged_user() -> bool:
        return has_authenticated_session()

    def has_imported_field_options() -> bool:
        return any(field_options.get(key) for key in (
            "escolha_reserva",
            "equipe",
            "componente",
            "publico",
            "turma",
            "principal_recurso",
            "tipo_atividade",
            "horario",
        ))

    def refresh_execution_buttons() -> None:
        state = "normal" if has_valid_logged_user() else "disabled"
        if run_test_button:
            run_test_button.config(state=state)
        if run_real_button:
            run_real_button.config(state=state)

    def refresh_logged_user_label() -> None:
        nonlocal logged_user_label
        if logged_user_label:
            logged_user_label.config(text=f"{get_logged_user_display()}")
            root.update_idletasks()
        refresh_execution_buttons()

    def set_logged_user_label(user: str) -> None:
        nonlocal logged_user_label
        if logged_user_label and user and user not in {"Nenhum usuário logado", "Usuário desconhecido"}:
            logged_user_label.config(text=user)
            root.update_idletasks()
        refresh_execution_buttons()

    def refresh_logged_user_label_with_retry(attempt: int = 0) -> None:
        user = get_logged_user_display()
        refresh_logged_user_label()
        if user == "Nenhum usuário logado" and attempt < 10:
            root.after(500, lambda: refresh_logged_user_label_with_retry(attempt + 1))

    def refresh_status_label(value: str) -> None:
        if status_label:
            status_label.config(text=value)
            root.update_idletasks()

    def save_bookings_url(show_success: bool = True) -> Optional[str]:
        url = normalize_bookings_url(bookings_url_var.get())
        if not is_valid_bookings_url(url):
            messagebox.showerror(
                "Link inválido",
                "Informe o link completo do Bookings da unidade.\n\nExemplo: https://outlook.office.com/book/...",
            )
            return None
        app_settings["bookings_url"] = url
        save_app_settings(app_settings)
        bookings_url_var.set(url)
        if show_success:
            messagebox.showinfo("Link salvo", "O link da unidade foi salvo com sucesso.")
        return url

    def ensure_saved_bookings_url() -> Optional[str]:
        url = normalize_bookings_url(bookings_url_var.get())
        saved_url = normalize_bookings_url(app_settings.get("bookings_url", ""))
        if url != saved_url:
            return save_bookings_url(show_success=False)
        if not is_valid_bookings_url(saved_url):
            messagebox.showerror("Link obrigatório", "Salve o link do Bookings da unidade antes de continuar.")
            return None
        return saved_url

    def ask_login_confirmation(user: str) -> bool:
        result = {"value": False}
        done = threading.Event()

        def _ask() -> None:
            dialog = None

            def finish(value: bool) -> None:
                result["value"] = value
                try:
                    if dialog is not None and dialog.winfo_exists():
                        dialog.grab_release()
                        dialog.destroy()
                except Exception:
                    pass
                done.set()

            try:
                root.deiconify()
                root.lift()
                root.attributes("-topmost", True)
                root.focus_force()
                root.after(1200, lambda: root.attributes("-topmost", False))
            except Exception:
                pass

            shown_user = user if user and user not in {"Nenhum usuário logado", "Usuário desconhecido"} else "usuário autenticado"
            dialog = tk.Toplevel(root)
            dialog.title("Confirmar login")
            dialog.transient(root)
            dialog.resizable(False, False)
            dialog.configure(bg=colors["panel_alt"])
            dialog.protocol("WM_DELETE_WINDOW", lambda: finish(False))

            try:
                dialog.grab_set()
                dialog.lift()
                dialog.attributes("-topmost", True)
                dialog.focus_force()
            except Exception:
                pass

            container = ttk.Frame(dialog, padding=20, style="Toolbar.TFrame")
            container.pack(fill="both", expand=True)

            ttk.Label(
                container,
                text="Confirme o login somente depois de voltar para a página inicial do Bookings.",
                style="Title.TLabel",
                wraplength=430,
                justify="left",
            ).pack(anchor="w")
            ttk.Label(
                container,
                text=f"Usuário detectado: {shown_user}",
                style="User.TLabel",
                wraplength=430,
                justify="left",
            ).pack(anchor="w", pady=(12, 0))
            ttk.Label(
                container,
                text="Se a tela ainda estiver no login da Microsoft, não confirme agora.",
                style="Subtitle.TLabel",
                wraplength=430,
                justify="left",
            ).pack(anchor="w", pady=(10, 0))

            warning_box = ttk.Frame(container, style="Card.TFrame", padding=12)
            warning_box.pack(fill="x", pady=(16, 0))
            ttk.Label(
                warning_box,
                text="Confirme apenas quando o navegador já tiver voltado para o formulário de agendamento da unidade.",
                style="MutedPanel.TLabel",
                wraplength=410,
                justify="left",
            ).pack(anchor="w")

            buttons = ttk.Frame(container, style="Toolbar.TFrame")
            buttons.pack(fill="x", pady=(18, 0))
            buttons.columnconfigure(0, weight=1)

            cancel_button = ttk.Button(buttons, text="Ainda não voltei", command=lambda: finish(False))
            cancel_button.grid(row=0, column=0, sticky="w")

            confirm_button = ttk.Button(buttons, text="Já voltei ao Bookings", style="Primary.TButton", state="disabled")
            confirm_button.grid(row=0, column=1, sticky="e", padx=(12, 0))

            countdown_label = ttk.Label(container, text="Confirmação liberada em 4s...", style="Subtitle.TLabel")
            countdown_label.pack(anchor="e", pady=(10, 0))

            countdown = {"value": 4}

            def tick() -> None:
                if not dialog.winfo_exists():
                    return
                if countdown["value"] <= 0:
                    countdown_label.config(text="Confirmação liberada.")
                    confirm_button.config(state="normal", command=lambda: finish(True))
                    cancel_button.focus_set()
                    return
                countdown_label.config(text=f"Confirmação liberada em {countdown['value']}s...")
                countdown["value"] -= 1
                dialog.after(1000, tick)

            tick()

            dialog.update_idletasks()
            width = dialog.winfo_reqwidth()
            height = dialog.winfo_reqheight()
            x = root.winfo_rootx() + max((root.winfo_width() - width) // 2, 40)
            y = root.winfo_rooty() + max((root.winfo_height() - height) // 3, 40)
            dialog.geometry(f"{width}x{height}+{x}+{y}")

        root.after(0, _ask)
        done.wait()
        return result["value"]

    def save_login_session_and_refresh(selectors_cfg: Dict[str, Any]) -> None:
        try:
            save_login_session(
                selectors_cfg,
                confirm_login_callback=ask_login_confirmation,
                status_callback=lambda value: root.after(0, lambda: refresh_status_label(value)),
            )
        except Exception as exc:
            log(f"Erro ao salvar login: {exc}")
            root.after(0, lambda: refresh_status_label("Erro ao concluir login."))
            root.after(0, lambda: messagebox.showerror("Erro no login", f"Falha ao concluir login: {exc}"))
            return
        root.after(0, refresh_logged_user_label_with_retry)
        root.after(0, lambda: refresh_status_label("Login concluído."))

    def on_save_login():
        if not ensure_saved_bookings_url():
            return
        cfg = load_json(SELECTORS_PATH)
        refresh_status_label("Aguardando novo login...")
        thread = threading.Thread(target=save_login_session_and_refresh, args=(cfg,), daemon=True)
        thread.start()
        messagebox.showinfo("Info", "Abra o navegador, faça login e, quando voltar para o Bookings, confirme na janela do programa.")

    def on_switch_user():
        if STORAGE_STATE_PATH.exists() or PROFILE_DIR.exists():
            if not messagebox.askyesno("Trocar usuário", "Isso irá desconectar o usuário atual e abrir o navegador para novo login. Deseja continuar?"):
                return
            clear_login_session()
            root.after(0, refresh_logged_user_label)
        on_save_login()

    def on_exit_app():
        if messagebox.askyesno("Sair da sessão", "Deseja sair e limpar a sessão de login?\n\nVocê precisará fazer login novamente."):
            clear_login_session()
            root.after(0, refresh_logged_user_label)
            root.after(0, lambda: refresh_status_label("Sessão limpa. Faça login novamente."))
            messagebox.showinfo("Logout concluído", "Sessão encerrada com sucesso.\n\nClique em 'Fazer Login / Trocar usuário' para acessar novamente.")

    def on_inspect():
        if not ensure_saved_bookings_url():
            return
        thread = threading.Thread(target=inspect_mode, daemon=True)
        thread.start()

    def on_test():
        if not ensure_saved_bookings_url():
            return
        run_jobs(headless=False, wait_for_enter=False, root=root)

    def on_real():
        if not ensure_saved_bookings_url():
            return
        run_jobs(headless=False, root=root)

    def on_calibrate():
        if not ensure_saved_bookings_url():
            return
        thread = threading.Thread(target=calibrate_mode, daemon=True)
        thread.start()

    def on_debug():
        if not ensure_saved_bookings_url():
            return
        print("\nModo DEBUG: executando com apenas 1 agendamento para inspeção...")
        log("=" * 70)
        log("MODO DEBUG INICIADO")
        log("=" * 70)
        jobs_raw = load_json(JOBS_PATH)
        # Usar apenas primeiro job com primeira turma e 1 data
        for item in jobs_raw.get("jobs", []):
            item["confirmar_envio_real"] = False
            # Limitar a 1 turma
            item["turmas"] = item["turmas"][:1] if item.get("turmas") else []
            # Set data_fim to data_inicio for 1 data
            item["data_fim"] = item.get("data_inicio", "14/04/2026")
        temp_path = BASE_DIR / "__tmp_jobs_debug__.json"
        temp_path.write_text(json.dumps(jobs_raw, ensure_ascii=False, indent=2), encoding="utf-8")
        original_jobs_path = JOBS_PATH
        JOBS_PATH = temp_path
        try:
            run_jobs(headless=False, wait_for_enter=False, root=root)
        finally:
            JOBS_PATH = original_jobs_path
            if temp_path.exists():
                temp_path.unlink()

    def save_selectors_cfg() -> None:
        selectors_cfg["field_values"] = field_options
        SELECTORS_PATH.write_text(json.dumps(selectors_cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        log("Configurações de campo salvas no JSON.")

    def import_field_options_and_save() -> None:
        if not ensure_saved_bookings_url():
            return
        if not has_authenticated_session():
            messagebox.showerror(
                "Login obrigatório",
                "Faça login com o usuário da unidade antes de importar as opções direto do Bookings.",
            )
            return

        refresh_status_label("Importando opções do Bookings...")

        def worker() -> None:
            try:
                imported = import_field_options_from_bookings(selectors_cfg)
            except Exception as exc:
                log(f"Erro ao importar opções do Bookings: {exc}")
                root.after(0, lambda: refresh_status_label("Falha ao importar opções."))
                root.after(
                    0,
                    lambda: messagebox.showerror(
                        "Importação não concluída",
                        f"Não foi possível importar as opções direto do Bookings.\n\n{exc}",
                    ),
                )
                return

            summary_lines: list[str] = []
            for field_key, values in imported.items():
                existing_values = list(field_options.get(field_key, []))
                if field_key == "horario":
                    new_values = _merge_horario_options(existing_values, values)
                else:
                    new_values = _dedupe_option_values(values)

                if not new_values:
                    continue

                field_options[field_key] = new_values
                summary_lines.append(f"{field_key}: {len(new_values)} opções")

            root.after(0, save_selectors_cfg)
            root.after(0, lambda: refresh_status_label("Feche e abra o app para aplicar os campos importados."))
            root.after(
                0,
                lambda: messagebox.showinfo(
                    "Importação concluída",
                    "As opções foram lidas direto da tela do Bookings.\n\n"
                    + "\n".join(summary_lines or ["Nenhum campo novo foi identificado."])
                    + "\n\nReinicie o app para aplicar as novas listas nos campos do formulário.",
                ),
            )

        threading.Thread(target=worker, daemon=True).start()

    def open_field_options_editor():
        editor = tk.Toplevel(root)
        editor.title("Editor de opções de campos")
        editor.geometry("600x500")

        editable_fields = [
            "escolha_reserva",
            "componente",
            "publico",
            "turma",
            "principal_recurso",
            "tipo_atividade",
            "equipe",
            "horario",
        ]

        selected_field = StringVar(value=editable_fields[0])

        def load_field_values(*args) -> None:
            field = selected_field.get()
            values = field_options.get(field, [])
            text_widget.delete("1.0", tk.END)
            text_widget.insert(tk.END, ", ".join(values))
            field_label.config(text=f"Campo: {field}")

        def save_field_values() -> None:
            field = selected_field.get()
            raw = text_widget.get("1.0", tk.END).strip()
            values = [item.strip() for item in raw.split(",") if item.strip()]
            field_options[field] = values
            save_selectors_cfg()
            messagebox.showinfo("Salvo", f"Valores salvos para '{field}'.\nReinicie o app para aplicar novos valores.")

        top_frame = ttk.Frame(editor, padding=10)
        top_frame.pack(fill="x")
        ttk.Label(top_frame, text="Selecione o campo:").pack(side="left")
        field_combo = ttk.Combobox(top_frame, textvariable=selected_field, values=editable_fields, state="readonly", width=30)
        field_combo.pack(side="left", padx=8)
        field_combo.bind("<<ComboboxSelected>>", load_field_values)

        field_label = ttk.Label(editor, text=f"Campo: {selected_field.get()}", font=("Arial", 10, "bold"))
        field_label.pack(anchor="w", padx=10, pady=(10, 0))

        text_widget = tk.Text(editor, wrap="word", height=18)
        text_widget.pack(fill="both", expand=True, padx=10, pady=10)

        button_frame = ttk.Frame(editor)
        button_frame.pack(fill="x", padx=10, pady=10)
        ttk.Button(button_frame, text="Salvar valores", command=save_field_values).pack(side="left")
        ttk.Button(button_frame, text="Recarregar campo", command=load_field_values).pack(side="left", padx=6)
        ttk.Button(button_frame, text="Importar do Bookings", command=import_field_options_and_save).pack(side="left", padx=6)

        load_field_values()

    def open_dev_mode():
        dev_window = tk.Toplevel(root)
        dev_window.title("Modo Desenvolvedor")
        dev_window.geometry("450x365")
        ttk.Button(dev_window, text="Salvar Login", command=on_save_login).pack(pady=5, fill="x", padx=20)
        ttk.Button(dev_window, text="Modo Inspeção", command=on_inspect).pack(pady=5, fill="x", padx=20)
        ttk.Button(dev_window, text="Calibração", command=on_calibrate).pack(pady=5, fill="x", padx=20)
        ttk.Button(dev_window, text="Modo DEBUG", command=on_debug).pack(pady=5, fill="x", padx=20)
        ttk.Separator(dev_window, orient="horizontal").pack(fill="x", pady=10, padx=20)
        ttk.Button(dev_window, text="Importar opções do Bookings", command=import_field_options_and_save).pack(pady=5, fill="x", padx=20)
        ttk.Button(dev_window, text="Editar opções de campos", command=open_field_options_editor).pack(pady=5, fill="x", padx=20)

    root.bind("<Control-Shift-D>", lambda event: open_dev_mode())

    brand_card = tk.Frame(
        menu_frame,
        bg=colors["panel_alt"],
        highlightthickness=1,
        highlightbackground=colors["border"],
        bd=0,
        padx=18,
        pady=16,
    )
    brand_card.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 18))
    brand_accent = tk.Frame(brand_card, bg=colors["accent_red"], width=5)
    brand_accent.grid(row=0, column=0, rowspan=3, sticky="ns", padx=(0, 14))
    header_logo_image = load_logo_image(subsample=2)
    text_column = 1
    if header_logo_image is not None:
        logo_badge = tk.Frame(
            brand_card,
            bg=colors["accent_red_soft"],
            highlightthickness=1,
            highlightbackground="#efb6bc",
            bd=0,
            padx=10,
            pady=8,
        )
        logo_badge.grid(row=0, column=1, rowspan=3, sticky="nw", padx=(0, 16))
        header_logo_label = tk.Label(
            logo_badge,
            image=header_logo_image,
            bg=colors["accent_red_soft"],
            bd=0,
            highlightthickness=0,
        )
        header_logo_label.image = header_logo_image
        header_logo_label.pack()
        text_column = 2
    tk.Label(
        brand_card,
        text=APP_NAME,
        bg=colors["panel_alt"],
        fg=colors["text"],
        font=("Segoe UI Semibold", 15),
        wraplength=420,
        justify="left",
        anchor="w",
    ).grid(row=0, column=text_column, sticky="w")
    tk.Label(
        brand_card,
        text=APP_SUBTITLE,
        bg=colors["panel_alt"],
        fg=colors["muted"],
        font=("Segoe UI", 10),
        wraplength=420,
        justify="left",
        anchor="w",
    ).grid(row=1, column=text_column, sticky="w", pady=(6, 0))
    logged_user_label = tk.Label(
        brand_card,
        text=f"{get_logged_user_display()}",
        bg=colors["panel_alt"],
        fg=colors["accent"],
        font=("Segoe UI Semibold", 11),
    )
    logged_user_label.grid(row=2, column=text_column, sticky="w", pady=(14, 0))

    actions_frame = tk.Frame(menu_frame, bg=colors["panel_alt"], bd=0, highlightthickness=0)
    actions_frame.grid(row=0, column=1, sticky="n", pady=(4, 0))
    def _rebuild_top_buttons() -> None:
        for child in list(actions_frame.winfo_children()):
            try:
                child.grid_forget()
            except Exception:
                pass
            child.destroy()
        ttk.Button(actions_frame, text="Fazer Login / Trocar usuário", command=on_switch_user, width=24, style="Soft.TButton").pack(side="left", padx=(0, 10))
        ttk.Button(actions_frame, text="Fazer Logout", command=on_exit_app, width=15, style="Soft.TButton").pack(side="left", padx=(0, 10))
        ttk.Button(actions_frame, text="Importar campos do Bookings", command=import_field_options_and_save, width=27, style="Primary.TButton").pack(side="left")

    _rebuild_top_buttons()

    bookings_url_card = tk.Frame(
        menu_frame,
        bg=colors["panel_alt"],
        highlightthickness=1,
        highlightbackground=colors["border"],
        bd=0,
        padx=16,
        pady=14,
    )
    bookings_url_card.grid(row=1, column=1, sticky="ew", pady=(14, 0))
    bookings_url_card.columnconfigure(1, weight=1)
    tk.Label(
        bookings_url_card,
        text="Link do Bookings da unidade",
        bg=colors["panel_alt"],
        fg=colors["text"],
        font=("Segoe UI Semibold", 11),
    ).grid(row=0, column=0, sticky="w", padx=(0, 12))
    bookings_url_entry = ttk.Entry(bookings_url_card, textvariable=bookings_url_var, width=72)
    bookings_url_entry.grid(row=0, column=1, sticky="ew", padx=(0, 10))
    ttk.Button(bookings_url_card, text="Salvar link", command=save_bookings_url).grid(row=0, column=2, sticky="e")
    tk.Label(
        bookings_url_card,
        text="O OED deve informar aqui o link de agendamentos da unidade antes de fazer login ou reservar.",
        bg=colors["panel_alt"],
        fg=colors["muted"],
        font=("Segoe UI", 10),
    ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 0))
    bookings_url_entry.bind("<Return>", lambda event: save_bookings_url())

    # Área principal em duas colunas: formulário à esquerda e resultados à direita
    main_content_frame = ttk.Frame(root, style="App.TFrame")
    main_content_frame.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 8))
    main_content_frame.rowconfigure(0, weight=1)
    main_content_frame.columnconfigure(0, weight=1, uniform="main_split")
    main_content_frame.columnconfigure(1, weight=1, uniform="main_split")

    # Formulário
    form_canvas_frame = ttk.Frame(main_content_frame, style="Card.TFrame", padding=10)
    form_canvas_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
    form_canvas_frame.rowconfigure(0, weight=1)
    form_canvas_frame.columnconfigure(0, weight=1)
    
    form_canvas = tk.Canvas(form_canvas_frame, bg=colors["panel"], highlightthickness=0)
    form_scrollbar = ttk.Scrollbar(form_canvas_frame, orient="vertical", command=form_canvas.yview)
    form_frame_inner = ttk.Frame(form_canvas, style="App.TFrame", padding=(0, 0, 0, 6))
    form_frame_inner.bind("<Configure>", lambda e: form_canvas.configure(scrollregion=form_canvas.bbox("all")))

    form_window = form_canvas.create_window((0, 0), window=form_frame_inner, anchor="nw")
    form_canvas.configure(yscrollcommand=form_scrollbar.set)
    form_canvas.pack(side="left", fill="both", expand=True)
    form_scrollbar.pack(side="right", fill="y")

    def on_form_canvas_resize(event):
        form_canvas.itemconfigure(form_window, width=event.width)

    form_canvas.bind("<Configure>", on_form_canvas_resize)
    
    # Bind scroll wheel
    def on_mousewheel(event):
        form_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    form_canvas.bind_all("<MouseWheel>", on_mousewheel)
    
    form_frame_inner.columnconfigure(0, weight=1)
    form_shell = ttk.Frame(form_frame_inner, style="App.TFrame")
    form_shell.grid(row=0, column=0, sticky="ew")
    form_shell.columnconfigure(0, weight=1)
    form_shell.columnconfigure(1, weight=3)
    form_shell.columnconfigure(2, weight=1)

    form_frame = ttk.Frame(form_shell, style="App.TFrame")
    form_frame.grid(row=0, column=1, sticky="ew")
    form_frame.columnconfigure(1, weight=1)

    available_turmas = field_options.get("turma") or []
    reserva_options = field_options.get("escolha_reserva") or []
    weekday_options = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta"]
    horario_options = field_options.get("horario") or []
    reserva_vars: Dict[str, BooleanVar] = {option: BooleanVar(value=(i == 0)) for i, option in enumerate(reserva_options)}
    default_weekday = ""
    default_horario = ""

    # Reserva
    general_frame = ttk.Frame(form_frame, style="Card.TFrame", padding=14)
    general_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=6)
    general_frame.columnconfigure(1, weight=1)

    ttk.Label(general_frame, text="Reserva", style="Section.TLabel").grid(row=0, column=0, sticky="w")
    ttk.Label(general_frame, text="Escolha o tipo de reserva antes de executar.", style="MutedPanel.TLabel").grid(row=1, column=0, sticky="w", pady=(2, 10))
    reserva_check_frame = ttk.Frame(general_frame, style="Inline.TFrame")
    reserva_check_frame.grid(row=0, column=1, rowspan=2, sticky="w")
    def select_reserva(selected: str) -> None:
        for opt, var in reserva_vars.items():
            var.set(opt == selected)
    if reserva_options:
        for i, option in enumerate(reserva_options):
            ttk.Checkbutton(reserva_check_frame, text=option, variable=reserva_vars[option], command=lambda opt=option: select_reserva(opt)).grid(row=0, column=i, padx=6)
    else:
        ttk.Label(
            reserva_check_frame,
            text="Faça login e clique em 'Importar campos do Bookings' para carregar as opções da unidade.",
            style="MutedPanel.TLabel",
            wraplength=330,
            justify="left",
        ).grid(row=0, column=0, sticky="w")

    fields = [
        ("Equipe", "equipe", "", field_options.get("equipe") or []),
        ("Notas", "notas", "Aula recorrente criada via automação.", None),
        ("Componente", "componente", "", field_options.get("componente") or []),
        ("Público", "publico", "", field_options.get("publico") or []),
        ("Principal recurso", "principal_recurso", "", field_options.get("principal_recurso") or []),
        ("Tipo de atividade", "tipo_atividade", "", field_options.get("tipo_atividade") or []),
        ("Data início", "data_inicio", "", None),
        ("Data fim", "data_fim", "", None),
    ]

    entries: Dict[str, StringVar] = {}
    row = 1
    for label_text, key, default, options in fields:
        label = ttk.Label(form_frame, text=label_text)
        label.grid(row=row, column=0, sticky="e", pady=6, padx=(2, 14))
        value = StringVar(value=default)
        if options is not None:
            combo = ttk.Combobox(form_frame, textvariable=value, values=options, width=48, state="readonly")
            combo.grid(row=row, column=1, sticky="ew", pady=6)
            combo.set(default)
        else:
            entry = ttk.Entry(form_frame, textvariable=value, width=50)
            entry.grid(row=row, column=1, sticky="ew", pady=6)
        entries[key] = value
        row += 1
    entries["escolha_reserva"] = StringVar(value="")
    def refresh_reserva() -> None:
        selected = [opt for opt, var in reserva_vars.items() if var.get()]
        if selected:
            entries["escolha_reserva"].set(selected[0])
        else:
            entries["escolha_reserva"].set("")
    for var in reserva_vars.values():
        var.trace_add("write", lambda *args: refresh_reserva())

    # Turmas section
    ttk.Label(form_frame, text="Turmas", style="Section.TLabel").grid(row=row, column=0, sticky="w", pady=(12, 6))
    turmas_frame = ttk.Frame(form_frame, style="Card.TFrame", padding=12)
    turmas_frame.grid(row=row, column=1, sticky="ew", pady=(12, 6))
    selected_turmas = []
    turma_vars = {}
    if available_turmas:
        for i, turma in enumerate(available_turmas):
            var = BooleanVar()
            chk = ttk.Checkbutton(turmas_frame, text=turma, variable=var)
            chk.grid(row=0, column=i, padx=6)
            turma_vars[turma] = var
    else:
        ttk.Label(
            turmas_frame,
            text="As turmas serão exibidas aqui depois da leitura dos campos da página de agendamentos.",
            style="MutedPanel.TLabel",
            wraplength=420,
            justify="left",
        ).grid(row=0, column=0, sticky="w")

    row += 1

    # Turma details frame
    details_frame = ttk.Frame(form_frame, style="App.TFrame")
    details_frame.grid(row=row, column=0, columnspan=2, sticky="ew", pady=6)
    turma_details: Dict[str, Dict[str, StringVar]] = {}

    def update_turma_details():
        nonlocal selected_turmas
        selected_turmas = [t for t in available_turmas if turma_vars[t].get()]
        existing_values = {
            turma: {
                "dias_semana": details["dias_semana"].get(),
                "horario": details["horario"].get(),
            }
            for turma, details in turma_details.items()
        }
        turma_details.clear()
        # Clear details_frame
        for widget in details_frame.winfo_children():
            widget.destroy()
        if not selected_turmas:
            return
        ttk.Label(details_frame, text="Detalhes por turma", style="Section.TLabel").pack(anchor="w", pady=(4, 6))
        for turma in selected_turmas:
            sub_frame = ttk.Frame(details_frame, style="Card.TFrame", padding=12)
            sub_frame.pack(fill="x", pady=4, padx=6)
            sub_frame.columnconfigure(5, weight=1)
            ttk.Label(sub_frame, text=f"Turma {turma}:").grid(row=0, column=0, sticky="w", padx=(6, 10), pady=6)
            ttk.Label(sub_frame, text="Dia da semana:").grid(row=0, column=1, sticky="w", padx=(0, 4))
            dias_padrao = existing_values.get(turma, {}).get("dias_semana") or default_weekday
            dias_var = StringVar(value=dias_padrao)
            ttk.Combobox(sub_frame, textvariable=dias_var, values=weekday_options, state="readonly", width=18).grid(row=0, column=2, padx=(0, 12), pady=6)
            ttk.Label(sub_frame, text="Horário:").grid(row=0, column=3, sticky="w", padx=(0, 4))
            horario_padrao = existing_values.get(turma, {}).get("horario") or default_horario
            horario_var = StringVar(value=horario_padrao)
            ttk.Combobox(sub_frame, textvariable=horario_var, values=horario_options, state="readonly", width=10).grid(row=0, column=4, padx=(0, 6), pady=6)
            turma_details[turma] = {"dias_semana": dias_var, "horario": horario_var}

    # Bind to update
    for var in turma_vars.values():
        var.trace_add("write", lambda *args: update_turma_details())

    stop_event = threading.Event()

    def _run_in_thread(confirm_real: bool) -> None:
        # Clear results
        if results_text_widget:
            root.after(0, lambda: results_text_widget.config(state="normal"))
            root.after(0, lambda: results_text_widget.delete(1.0, tk.END))
            root.after(0, lambda: results_text_widget.config(state="disabled"))
        stop_event.clear()
        status_label.config(text="Executando automação... aguarde")
        root.update_idletasks()
        try:
            job_data = _build_gui_job_data(entries, selected_turmas, turma_details)
            validate_recurrence_period(job_data["data_inicio"], job_data["data_fim"])
            _run_jobs_with_job_data(
                job_data,
                confirm_real,
                stop_event,
                headless=False,
                root=root,
                on_user_detected=lambda user: root.after(0, lambda: set_logged_user_label(user)),
            )
            if stop_event.is_set():
                messagebox.showinfo("Parado", "Execução interrompida pelo usuário.")
                status_label.config(text="Execução parada.")
            else:
                messagebox.showinfo("Concluído", "Execução concluída com sucesso.")
                status_label.config(text="Execução concluída.")
        except Exception as exc:
            messagebox.showerror("Erro", f"Falha na execução: {exc}")
            status_label.config(text="Erro durante a execução.")

    def on_run_test() -> None:
        if not ensure_saved_bookings_url():
            return
        if not has_valid_logged_user():
            messagebox.showwarning("Login obrigatório", "Faça login no Bookings antes de executar o teste de reserva.")
            return
        if not has_imported_field_options():
            messagebox.showwarning("Campos não carregados", "Clique em 'Importar campos do Bookings' para ler os campos da unidade antes de continuar.")
            return
        required_fields = [
            ("Equipe", "equipe"),
            ("Componente", "componente"),
            ("Público", "publico"),
            ("Principal recurso", "principal_recurso"),
            ("Tipo de atividade", "tipo_atividade"),
            ("Data início", "data_inicio"),
            ("Data fim", "data_fim"),
        ]
        for label_text, key in required_fields:
            if not entries[key].get().strip():
                messagebox.showwarning("Campo obrigatório", f"Preencha ou selecione '{label_text}' antes de continuar.")
                return
        if not entries["escolha_reserva"].get().strip():
            messagebox.showwarning("Campo obrigatório", "Selecione o tipo de reserva antes de continuar.")
            return
        if not selected_turmas:
            messagebox.showerror("Erro", "Selecione pelo menos uma turma.")
            return
        stop_event.clear()
        thread = threading.Thread(target=_run_in_thread, args=(False,), daemon=True)
        thread.start()

    def on_run_real() -> None:
        if not ensure_saved_bookings_url():
            return
        if not has_valid_logged_user():
            messagebox.showwarning("Login obrigatório", "Faça login no Bookings antes de executar a reserva.")
            return
        if not has_imported_field_options():
            messagebox.showwarning("Campos não carregados", "Clique em 'Importar campos do Bookings' para ler os campos da unidade antes de continuar.")
            return
        required_fields = [
            ("Equipe", "equipe"),
            ("Componente", "componente"),
            ("Público", "publico"),
            ("Principal recurso", "principal_recurso"),
            ("Tipo de atividade", "tipo_atividade"),
            ("Data início", "data_inicio"),
            ("Data fim", "data_fim"),
        ]
        for label_text, key in required_fields:
            if not entries[key].get().strip():
                messagebox.showwarning("Campo obrigatório", f"Preencha ou selecione '{label_text}' antes de continuar.")
                return
        if not entries["escolha_reserva"].get().strip():
            messagebox.showwarning("Campo obrigatório", "Selecione o tipo de reserva antes de continuar.")
            return
        if not selected_turmas:
            messagebox.showerror("Erro", "Selecione pelo menos uma turma.")
            return
        if not messagebox.askyesno("Confirmar reserva", "Deseja realmente executar a reserva?"):
            return
        stop_event.clear()
        thread = threading.Thread(target=_run_in_thread, args=(True,), daemon=True)
        thread.start()

    def on_stop() -> None:
        stop_event.set()
        status_label.config(text="Solicitação de parada recebida...")

    # Coluna da direita: ações em um card e resultados em outro
    right_panel = ttk.Frame(main_content_frame, style="App.TFrame")
    right_panel.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
    right_panel.rowconfigure(1, weight=1)
    right_panel.columnconfigure(0, weight=1)

    actions_card = ttk.Frame(right_panel, padding=14, style="Card.TFrame")
    actions_card.grid(row=0, column=0, sticky="ew", pady=(0, 10))
    actions_card.columnconfigure(0, weight=1)
    actions_card.columnconfigure(1, weight=0)

    actions_row = ttk.Frame(actions_card, style="Card.TFrame")
    actions_row.grid(row=0, column=0, columnspan=2, sticky="ew")
    actions_row.columnconfigure(0, weight=1)
    actions_row.columnconfigure(1, weight=0)

    button_group = ttk.Frame(actions_row, style="Card.TFrame")
    button_group.grid(row=0, column=0, sticky="w")

    run_test_button = ttk.Button(button_group, text="Teste de reservar", command=on_run_test, style="Soft.TButton")
    run_test_button.grid(row=0, column=0, padx=(0, 6))
    run_real_button = ttk.Button(button_group, text="Reservar", command=on_run_real, style="Primary.TButton")
    run_real_button.grid(row=0, column=1, padx=6)
    stop_button = ttk.Button(button_group, text="Parar", command=on_stop, style="Danger.TButton")
    stop_button.grid(row=0, column=2, padx=(6, 0))
    initial_status_text = "Pronto para executar" if has_imported_field_options() else "Faça login e importe os campos"
    status_label = ttk.Label(actions_row, text=initial_status_text, style="Status.TLabel")
    status_label.grid(row=0, column=1, sticky="e")

    results_frame = ttk.Frame(right_panel, padding=14, style="Card.TFrame")
    results_frame.grid(row=1, column=0, sticky="nsew")
    results_frame.rowconfigure(1, weight=1)
    results_frame.columnconfigure(0, weight=1)
    ttk.Label(results_frame, text="Resultados dos Agendamentos", style="Section.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
    results_text = tk.Text(results_frame, height=10, wrap="word", bg=colors["success_bg"], fg=colors["text"], insertbackground=colors["text"], relief="flat", highlightthickness=1, highlightbackground=colors["success_border"], padx=12, pady=12)
    results_scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=results_text.yview)
    results_text.configure(yscrollcommand=results_scrollbar.set)
    results_text.grid(row=1, column=0, sticky="nsew", padx=(0, 5))
    results_scrollbar.grid(row=1, column=1, sticky="ns")
    results_text.config(state="disabled")
    results_text_widget = results_text

    # Log area
    log_frame = ttk.Frame(root, padding=12, style="Card.TFrame")
    log_frame.grid(row=3, column=0, sticky="ew", padx=12, pady=(0, 8))
    log_frame.rowconfigure(1, weight=1)
    log_frame.columnconfigure(0, weight=1)
    ttk.Label(log_frame, text="Log da Execução", style="SecondarySection.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 6))
    log_text = tk.Text(log_frame, height=4, wrap="word", bg="#fbfcfb", fg="#6d7d78", insertbackground=colors["text"], relief="flat", highlightthickness=1, highlightbackground="#e3e8e6", padx=10, pady=10)
    scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=log_text.yview)
    log_text.configure(yscrollcommand=scrollbar.set)
    log_text.grid(row=1, column=0, sticky="nsew", padx=(0, 5))
    scrollbar.grid(row=1, column=1, sticky="ns")
    log_text.config(state="disabled")
    log_text_widget = log_text

    dev_label = ttk.Label(root, text="Desenvolvido por Breno Malta Silva", style="Footer.TLabel")
    dev_label.grid(row=5, column=0, sticky="ew", pady=(2, 10))

    refresh_execution_buttons()

    root.mainloop()


def menu() -> None:
    global JOBS_PATH
    print(f"\n=== {APP_NAME} ===")
    print("1 - Salvar/atualizar login manual")
    print("2 - Abrir modo inspeção")
    print("3 - Rodar jobs em modo teste (sem enviar)")
    print("4 - Rodar jobs com comportamento definido no JSON")
    print("5 - Modo calibração de seletores")
    print("6 - Modo DEBUG (modo teste + inspecionável)")
    print("7 - Abrir interface gráfica")
    print("0 - Sair")
    choice = input("Escolha: ").strip()

    if choice == "1":
        cfg = load_json(SELECTORS_PATH)
        save_login_session(cfg)
    elif choice == "2":
        inspect_mode()
    elif choice == "3":
        jobs_raw = load_json(JOBS_PATH)
        for item in jobs_raw.get("jobs", []):
            item["confirmar_envio_real"] = False
        temp_path = BASE_DIR / "__tmp_jobs_test__.json"
        temp_path.write_text(json.dumps(jobs_raw, ensure_ascii=False, indent=2), encoding="utf-8")
        original_jobs_path = JOBS_PATH
        JOBS_PATH = temp_path
        try:
            run_jobs(headless=False)
        finally:
            JOBS_PATH = original_jobs_path
            if temp_path.exists():
                temp_path.unlink()
    elif choice == "4":
        run_jobs(headless=False)
    elif choice == "5":
        calibrate_mode()
    elif choice == "6":
        print("\nModo DEBUG: executando com apenas 1 agendamento para inspeção...")
        log("=" * 70)
        log("MODO DEBUG INICIADO")
        log("=" * 70)
        jobs_raw = load_json(JOBS_PATH)
        # Usar apenas primeiro job com primeira turma e 1 data
        for item in jobs_raw.get("jobs", []):
            item["confirmar_envio_real"] = False
            # Limitar a 1 turma
            item["turmas"] = item["turmas"][:1] if item.get("turmas") else []
            # Set data_fim to data_inicio for 1 data
            item["data_fim"] = item.get("data_inicio", "14/04/2026")
        temp_path = BASE_DIR / "__tmp_jobs_debug__.json"
        temp_path.write_text(json.dumps(jobs_raw, ensure_ascii=False, indent=2), encoding="utf-8")
        original_jobs_path = JOBS_PATH
        JOBS_PATH = temp_path
        try:
            run_jobs(headless=False)
        finally:
            JOBS_PATH = original_jobs_path
            if temp_path.exists():
                temp_path.unlink()
    elif choice == "7":
        launch_gui()
    elif choice == "0":
        print("Saindo.")
    else:
        print("Opção inválida.")


if __name__ == "__main__":
    try:
        launch_gui()
    except KeyboardInterrupt:
        print("\nEncerrado pelo usuário.")
        sys.exit(0)
    except Exception as exc:
        fail(str(exc))
        sys.exit(1)
