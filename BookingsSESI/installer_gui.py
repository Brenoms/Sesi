from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import threading
import zipfile
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


APP_NAME = "SESI Reservas Recorrentes"
APP_SUBTITLE = "Instalação assistida para agendamentos recorrentes"
AUTHOR_TEXT = "Desenvolvido por Breno Malta Silva"
DEFAULT_INSTALL_DIR = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "BookingsSESI"


def resource_dir() -> Path:
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))


def payload_zip() -> Path:
    return resource_dir() / "payload.zip"


def logo_path() -> Path:
    prepared_root = resource_dir() / "sesi_logo_app.png"
    if prepared_root.exists():
        return prepared_root
    prepared = resource_dir() / "imagens" / "sesi_logo_app.png"
    if prepared.exists():
        return prepared
    preferred = resource_dir() / "imagens" / "sesi_logo_vermelha.png"
    if preferred.exists():
        return preferred
    cropped = resource_dir() / "sesi_logo_recortada.png"
    if cropped.exists():
        return cropped
    return resource_dir() / "sesi_logo_vermelha.png"


class InstallerApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(f"{APP_NAME} - Instalador")
        self.root.geometry("680x430")
        self.root.minsize(680, 430)
        self.root.resizable(False, False)
        self.root.configure(bg="#eef3f1")

        self.install_dir_var = tk.StringVar(value=str(DEFAULT_INSTALL_DIR))
        self.launch_after_install = tk.BooleanVar(value=True)
        self.progress_text = tk.StringVar(value="Pronto para instalar.")
        self.current_step = 0
        self.install_completed = False
        self.logo_image = None

        self._configure_style()
        self._build_ui()
        self._show_step(0)

    def _configure_style(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(".", font=("Segoe UI", 10))
        style.configure("Card.TFrame", background="#ffffff")
        style.configure("Hero.TFrame", background="#1f5f57")
        style.configure("HeroTitle.TLabel", background="#1f5f57", foreground="#ffffff", font=("Segoe UI Semibold", 18))
        style.configure("HeroText.TLabel", background="#1f5f57", foreground="#d9ebe7", font=("Segoe UI", 10))
        style.configure("Title.TLabel", background="#ffffff", foreground="#1c2b28", font=("Segoe UI Semibold", 15))
        style.configure("Body.TLabel", background="#ffffff", foreground="#5f6f6a", font=("Segoe UI", 10))
        style.configure("Small.TLabel", background="#ffffff", foreground="#73817d", font=("Segoe UI", 9))
        style.configure("Footer.TFrame", background="#ffffff")
        style.configure("TButton", font=("Segoe UI", 10), padding=(11, 6))
        style.configure("Primary.TButton", font=("Segoe UI Semibold", 10), padding=(12, 7))

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, style="Card.TFrame", padding=0)
        container.pack(fill="both", expand=True, padx=16, pady=16)

        hero = ttk.Frame(container, style="Hero.TFrame", padding=24)
        hero.pack(fill="x")
        hero.columnconfigure(0, weight=1)
        hero.columnconfigure(1, weight=0)
        text_frame = ttk.Frame(hero, style="Hero.TFrame")
        text_frame.grid(row=0, column=0, sticky="nw")
        ttk.Label(text_frame, text=APP_NAME, style="HeroTitle.TLabel").pack(anchor="w")
        ttk.Label(text_frame, text=APP_SUBTITLE, style="HeroText.TLabel").pack(anchor="w", pady=(6, 0))
        ttk.Label(text_frame, text=AUTHOR_TEXT, style="HeroText.TLabel").pack(anchor="w", pady=(8, 0))

        logo_file = logo_path()
        if logo_file.exists():
            try:
                self.logo_image = tk.PhotoImage(file=str(logo_file)).subsample(36, 36)
                logo_label = tk.Label(hero, image=self.logo_image, bg="#1f5f57", bd=0, highlightthickness=0)
                logo_label.grid(row=0, column=1, sticky="ne", padx=(18, 0), pady=(4, 0))
            except Exception:
                self.logo_image = None

        body = ttk.Frame(container, style="Card.TFrame", padding=24)
        body.pack(fill="both", expand=True)

        self.steps = []
        self.steps.append(self._build_welcome(body))
        self.steps.append(self._build_directory(body))
        self.steps.append(self._build_progress(body))
        self.steps.append(self._build_finish(body))

        footer = ttk.Frame(container, style="Footer.TFrame", padding=(24, 10, 24, 18))
        footer.pack(fill="x")
        footer.columnconfigure(0, weight=1)
        ttk.Label(footer, text=AUTHOR_TEXT, style="Small.TLabel").grid(row=0, column=0, sticky="w")
        self.back_button = ttk.Button(footer, text="Voltar", command=self._back)
        self.back_button.grid(row=0, column=1, padx=(0, 8))
        self.next_button = ttk.Button(footer, text="Avançar", style="Primary.TButton", command=self._next)
        self.next_button.grid(row=0, column=2)

    def _build_welcome(self, parent: ttk.Frame) -> ttk.Frame:
        frame = ttk.Frame(parent, style="Card.TFrame")
        ttk.Label(frame, text="Bem-vindo", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            frame,
            text=f"Este assistente irá instalar o {APP_NAME} no seu computador.\n\nO pacote já inclui o ambiente Python portátil e o Chromium do Playwright.",
            style="Body.TLabel",
            justify="left",
        ).pack(anchor="w", pady=(14, 0))
        return frame

    def _build_directory(self, parent: ttk.Frame) -> ttk.Frame:
        frame = ttk.Frame(parent, style="Card.TFrame")
        ttk.Label(frame, text="Escolha a pasta de instalação", style="Title.TLabel").pack(anchor="w")
        ttk.Label(frame, text="Recomendado: manter o caminho padrão.", style="Body.TLabel").pack(anchor="w", pady=(10, 14))

        path_frame = ttk.Frame(frame, style="Card.TFrame")
        path_frame.pack(fill="x")
        entry = ttk.Entry(path_frame, textvariable=self.install_dir_var, width=68)
        entry.pack(side="left", fill="x", expand=True)
        ttk.Button(path_frame, text="Procurar...", command=self._browse).pack(side="left", padx=(8, 0))

        ttk.Checkbutton(
            frame,
            text=f"Abrir o {APP_NAME} após concluir a instalação",
            variable=self.launch_after_install,
        ).pack(anchor="w", pady=(16, 0))
        return frame

    def _build_progress(self, parent: ttk.Frame) -> ttk.Frame:
        frame = ttk.Frame(parent, style="Card.TFrame")
        ttk.Label(frame, text="Instalando", style="Title.TLabel").pack(anchor="w")
        ttk.Label(frame, textvariable=self.progress_text, style="Body.TLabel").pack(anchor="w", pady=(12, 14))
        self.progress = ttk.Progressbar(frame, mode="indeterminate", length=580)
        self.progress.pack(anchor="w", fill="x")
        return frame

    def _build_finish(self, parent: ttk.Frame) -> ttk.Frame:
        frame = ttk.Frame(parent, style="Card.TFrame")
        ttk.Label(frame, text="Instalação concluída", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            frame,
            text=f"O {APP_NAME} foi instalado com sucesso.\n\nVocê pode usar o atalho criado na Área de Trabalho ou no Menu Iniciar.\nPara desinstalar, use a opção 'Desinstalar {APP_NAME}' no Menu Iniciar.",
            style="Body.TLabel",
            justify="left",
        ).pack(anchor="w", pady=(14, 0))
        return frame

    def _show_step(self, index: int) -> None:
        for step in self.steps:
            step.pack_forget()
        self.steps[index].pack(fill="both", expand=True)
        self.current_step = index

        self.back_button.state(["!disabled"] if index > 0 and not self.install_completed and index != 2 else ["disabled"])

        if index == 0:
            self.next_button.config(text="Avançar", command=self._next)
            self.next_button.state(["!disabled"])
        elif index == 1:
            self.next_button.config(text="Instalar", command=self._start_install)
            self.next_button.state(["!disabled"])
        elif index == 2:
            self.next_button.config(text="Instalando...", command=lambda: None)
            self.next_button.state(["disabled"])
        else:
            self.next_button.config(text="Concluir", command=self.root.destroy)
            self.next_button.state(["!disabled"])

    def _browse(self) -> None:
        selected = filedialog.askdirectory(initialdir=self.install_dir_var.get() or str(DEFAULT_INSTALL_DIR))
        if selected:
            self.install_dir_var.set(selected)

    def _back(self) -> None:
        if self.current_step > 0:
            self._show_step(self.current_step - 1)

    def _next(self) -> None:
        self._show_step(self.current_step + 1)

    def _start_install(self) -> None:
        install_dir = self.install_dir_var.get().strip()
        if not install_dir:
            messagebox.showerror("Instalador", "Informe uma pasta de instalação.")
            return
        self._show_step(2)
        self.progress.start(12)
        worker = threading.Thread(target=self._run_install, args=(install_dir,), daemon=True)
        worker.start()

    def _run_install(self, install_dir: str) -> None:
        try:
            archive = payload_zip()
            if not archive.exists():
                raise FileNotFoundError(f"Payload não encontrado: {archive}")

            extract_dir = Path(tempfile.gettempdir()) / "BookingsSESI_Payload_Run"
            self._set_progress("Extraindo arquivos do instalador...")
            if extract_dir.exists():
                shutil.rmtree(extract_dir, ignore_errors=True)
            extract_dir.mkdir(parents=True, exist_ok=True)

            with zipfile.ZipFile(archive, "r") as zf:
                zf.extractall(extract_dir)

            install_script = extract_dir / "install_from_package.ps1"
            if not install_script.exists():
                raise FileNotFoundError(f"Script de instalação não encontrado: {install_script}")

            self._set_progress("Copiando arquivos e criando atalhos...")
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            command = [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(install_script),
                "-InstallDir",
                install_dir,
            ]
            if not self.launch_after_install.get():
                command.append("-NoLaunch")

            result = subprocess.run(command, check=False, creationflags=creationflags)
            if result.returncode != 0:
                raise RuntimeError(f"A instalação retornou código {result.returncode}.")

            self.install_completed = True
            self.root.after(0, self._finish_success)
        except Exception as exc:
            self.root.after(0, lambda: self._finish_error(str(exc)))

    def _set_progress(self, text: str) -> None:
        self.root.after(0, lambda: self.progress_text.set(text))

    def _finish_success(self) -> None:
        self.progress.stop()
        self.progress_text.set("Instalação concluída com sucesso.")
        self._show_step(3)

    def _finish_error(self, message: str) -> None:
        self.progress.stop()
        self.progress_text.set("Falha durante a instalação.")
        self._show_step(1)
        messagebox.showerror("Instalador", f"Não foi possível concluir a instalação.\n\n{message}")

    def run(self) -> None:
        self.root.mainloop()


if __name__ == "__main__":
    InstallerApp().run()
