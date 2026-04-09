from __future__ import annotations

from pathlib import Path
from typing import Iterable, List

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
OUTPUT_PDF = ROOT / "Tutorial_Configuracao_Unidades.pdf"
LOGO_PATH = ROOT / "imagens" / "sesi_logo_vermelha.png"
PRINTS_DIR = ROOT / "imagens dos prints"

PAGE_W, PAGE_H = 1654, 2339
MARGIN_X = 110
MARGIN_Y = 110
CONTENT_W = PAGE_W - (MARGIN_X * 2)


def load_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = []
    if bold:
        candidates.extend(
            [
                Path(r"C:\Windows\Fonts\segoeuib.ttf"),
                Path(r"C:\Windows\Fonts\arialbd.ttf"),
            ]
        )
    else:
        candidates.extend(
            [
                Path(r"C:\Windows\Fonts\segoeui.ttf"),
                Path(r"C:\Windows\Fonts\arial.ttf"),
            ]
        )
    for path in candidates:
        if path.exists():
            return ImageFont.truetype(str(path), size=size)
    return ImageFont.load_default()


FONT_TITLE = load_font(52, bold=True)
FONT_SUBTITLE = load_font(26)
FONT_H1 = load_font(34, bold=True)
FONT_H2 = load_font(28, bold=True)
FONT_BODY = load_font(24)
FONT_SMALL = load_font(20)
FONT_CAPTION = load_font(22, bold=True)


def text_height(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont) -> int:
    box = draw.textbbox((0, 0), text, font=font)
    return box[3] - box[1]


def wrap_text(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, max_width: int) -> List[str]:
    words = text.split()
    if not words:
        return [""]
    lines: List[str] = []
    current = words[0]
    for word in words[1:]:
        candidate = f"{current} {word}"
        if draw.textlength(candidate, font=font) <= max_width:
            current = candidate
        else:
            lines.append(current)
            current = word
    lines.append(current)
    return lines


class PageBuilder:
    def __init__(self) -> None:
        self.pages: List[Image.Image] = []
        self.page: Image.Image | None = None
        self.draw: ImageDraw.ImageDraw | None = None
        self.y = MARGIN_Y
        self.new_page()

    def new_page(self) -> None:
        self.page = Image.new("RGB", (PAGE_W, PAGE_H), "#f6f8f7")
        self.draw = ImageDraw.Draw(self.page)
        self.pages.append(self.page)
        self.y = MARGIN_Y

    def ensure_space(self, needed: int) -> None:
        if self.y + needed > PAGE_H - MARGIN_Y:
            self.new_page()

    def add_header_band(self) -> None:
        assert self.page is not None and self.draw is not None
        self.draw.rounded_rectangle(
            (MARGIN_X, self.y, PAGE_W - MARGIN_X, self.y + 180),
            radius=26,
            fill="#1f5f57",
        )
        if LOGO_PATH.exists():
            logo = Image.open(LOGO_PATH).convert("RGBA")
            logo.thumbnail((220, 110))
            self.page.paste(logo, (MARGIN_X + 50, self.y + 35), logo)
        self.draw.text((MARGIN_X + 310, self.y + 40), "SESI Reservas Recorrentes", font=FONT_TITLE, fill="white")
        self.draw.text(
            (MARGIN_X + 310, self.y + 110),
            "Tutorial de configuração por unidade",
            font=FONT_SUBTITLE,
            fill="#d8ece7",
        )
        self.y += 230

    def add_h1(self, text: str) -> None:
        assert self.draw is not None
        self.ensure_space(70)
        self.draw.text((MARGIN_X, self.y), text, font=FONT_H1, fill="#16312d")
        self.y += 62

    def add_h2(self, text: str) -> None:
        assert self.draw is not None
        self.ensure_space(56)
        self.draw.text((MARGIN_X, self.y), text, font=FONT_H2, fill="#1f5f57")
        self.y += 50

    def add_paragraph(self, text: str, font: ImageFont.ImageFont = FONT_BODY, color: str = "#243432", spacing: int = 14) -> None:
        assert self.draw is not None
        lines = wrap_text(self.draw, text, font, CONTENT_W)
        line_h = text_height(self.draw, "Ag", font) + 8
        self.ensure_space((line_h * len(lines)) + spacing)
        for line in lines:
            self.draw.text((MARGIN_X, self.y), line, font=font, fill=color)
            self.y += line_h
        self.y += spacing

    def add_bullets(self, items: Iterable[str], font: ImageFont.ImageFont = FONT_BODY) -> None:
        assert self.draw is not None
        line_h = text_height(self.draw, "Ag", font) + 8
        for item in items:
            wrapped = wrap_text(self.draw, item, font, CONTENT_W - 40)
            self.ensure_space((line_h * len(wrapped)) + 8)
            self.draw.text((MARGIN_X, self.y), u"\u2022", font=font, fill="#1f5f57")
            for idx, line in enumerate(wrapped):
                self.draw.text((MARGIN_X + 32, self.y + (idx * line_h)), line, font=font, fill="#243432")
            self.y += line_h * len(wrapped) + 8
        self.y += 8

    def add_numbered(self, items: Iterable[str], font: ImageFont.ImageFont = FONT_BODY) -> None:
        assert self.draw is not None
        line_h = text_height(self.draw, "Ag", font) + 8
        for i, item in enumerate(items, start=1):
            wrapped = wrap_text(self.draw, item, font, CONTENT_W - 60)
            self.ensure_space((line_h * len(wrapped)) + 8)
            self.draw.text((MARGIN_X, self.y), f"{i}.", font=font, fill="#1f5f57")
            for idx, line in enumerate(wrapped):
                self.draw.text((MARGIN_X + 44, self.y + (idx * line_h)), line, font=font, fill="#243432")
            self.y += line_h * len(wrapped) + 8
        self.y += 8

    def add_note_box(self, title: str, lines: Iterable[str]) -> None:
        assert self.draw is not None
        line_h = text_height(self.draw, "Ag", FONT_SMALL) + 8
        wrapped_lines = []
        for line in lines:
            wrapped_lines.extend(wrap_text(self.draw, line, FONT_SMALL, CONTENT_W - 60))
            wrapped_lines.append("")
        if wrapped_lines and wrapped_lines[-1] == "":
            wrapped_lines.pop()
        total_h = 24 + 34 + len(wrapped_lines) * line_h + 24
        self.ensure_space(total_h + 16)
        x1, y1, x2, y2 = MARGIN_X, self.y, PAGE_W - MARGIN_X, self.y + total_h
        self.draw.rounded_rectangle((x1, y1, x2, y2), radius=20, fill="#edf4f2", outline="#c8ddd8", width=2)
        self.draw.text((x1 + 24, y1 + 20), title, font=FONT_CAPTION, fill="#1f5f57")
        cy = y1 + 62
        for line in wrapped_lines:
            if line:
                self.draw.text((x1 + 24, cy), line, font=FONT_SMALL, fill="#435652")
            cy += line_h
        self.y = y2 + 22

    def add_image_section(self, title: str, subtitle: str, image_path: Path | None, max_height: int = 760) -> None:
        assert self.draw is not None and self.page is not None
        self.add_h2(title)
        self.add_paragraph(subtitle, font=FONT_SMALL, color="#5f6f6a", spacing=16)

        if image_path is None or not image_path.exists():
            self.ensure_space(380)
            x1, y1, x2, y2 = MARGIN_X, self.y, PAGE_W - MARGIN_X, self.y + 340
            self.draw.rounded_rectangle((x1, y1, x2, y2), radius=18, fill="white", outline="#cfded9", width=3)
            self.draw.text((x1 + 30, y1 + 40), "Imagem não encontrada no projeto.", font=FONT_CAPTION, fill="#8a3b42")
            self.draw.text((x1 + 30, y1 + 90), "Envie o print para incluir nesta versão do PDF.", font=FONT_BODY, fill="#5f6f6a")
            self.y = y2 + 24
            return

        image = Image.open(image_path).convert("RGB")
        border = 16
        target_w = CONTENT_W - (border * 2)
        ratio = min(target_w / image.width, max_height / image.height, 1.0)
        resized = image.resize((int(image.width * ratio), int(image.height * ratio)), Image.Resampling.LANCZOS)

        box_h = resized.height + (border * 2)
        self.ensure_space(box_h + 36)
        x1 = MARGIN_X
        y1 = self.y
        x2 = PAGE_W - MARGIN_X
        y2 = y1 + box_h
        self.draw.rounded_rectangle((x1, y1, x2, y2), radius=18, fill="white", outline="#cfded9", width=3)
        self.page.paste(resized, (x1 + border, y1 + border))
        self.y = y2 + 24


def print_path(name: str) -> Path:
    return PRINTS_DIR / name


def build_pdf() -> None:
    doc = PageBuilder()
    doc.add_header_band()

    doc.add_h1("Objetivo")
    doc.add_paragraph(
        "Este material orienta cada unidade escolar a configurar o sistema para que os dados do aplicativo fiquem iguais aos dados reais do Microsoft Bookings de cada escola."
    )
    doc.add_h2("Antes de começar")
    doc.add_bullets(
        [
            "Tenha em mãos o link oficial do Bookings da unidade.",
            "Use o usuário correto da unidade para fazer login no sistema.",
            "Separe os nomes exatos dos campos que aparecem no formulário da escola, como equipe, público, componente, principal recurso, tipo de atividade, turmas e horários.",
            "Os nomes precisam ser cadastrados exatamente como aparecem no Bookings.",
        ]
    )
    doc.add_note_box(
        "Importante",
        [
            "Se um nome estiver diferente do que aparece no Bookings da unidade, a automação pode não localizar a opção correta.",
            "Sempre use primeiro o botão Teste de reservar antes de usar Reservar.",
        ],
    )

    doc.add_image_section(
        "1. Tela principal do sistema",
        "Nesta tela a unidade deve informar o link do Bookings, salvar o link e depois fazer login com o usuário correto da escola.",
        print_path("modo desemvolvedor.png"),
        max_height=620,
    )
    doc.add_numbered(
        [
            "Cole no campo 'Link do Bookings da unidade' o endereço oficial do agendamento da escola.",
            "Clique em 'Salvar link'.",
            "Clique em 'Fazer Login / Trocar usuário'.",
            "Faça o login com a conta da unidade.",
            "Quando o programa pedir, confirme o login na própria tela do sistema.",
        ]
    )

    doc.add_image_section(
        "2. Acesso ao Modo Desenvolvedor",
        "Depois de salvar o link e concluir o login, a escola deve abrir o Modo Desenvolvedor para ajustar os dados do seu formulário.",
        print_path("menu modo desenvolvedor.png"),
        max_height=500,
    )
    doc.add_bullets(
        [
            "A opção principal para as escolas é 'Editar opções de campos'.",
            "Os demais botões são técnicos e normalmente não precisam ser usados no dia a dia das unidades.",
        ]
    )

    doc.add_image_section(
        "3. Editor de opções de campos",
        "Nesta tela a unidade escolhe qual campo deseja alterar e informa os valores separados por vírgula, exatamente como aparecem no Bookings da escola.",
        print_path("editor de opocoes de campos.png"),
        max_height=760,
    )
    doc.add_bullets(
        [
            "Campos disponíveis para edição: escolha_reserva, componente, publico, turma, principal_recurso, tipo_atividade, equipe e horario.",
            "Digite os valores separados por vírgula.",
            "Clique em 'Salvar valores' ao concluir a alteração.",
            "Se quiser voltar ao conteúdo salvo anteriormente, clique em 'Recarregar campo'.",
        ]
    )
    doc.add_note_box(
        "Exemplo prático",
        [
            "No campo equipe, a unidade deve cadastrar os nomes exatamente como aparecem no agendamento da escola.",
            "Exemplo: FABLAB CE109, Ateliê de Arte CE109, Auditório I CE109.",
        ],
    )

    doc.add_h1("Ordem recomendada de configuração")
    doc.add_numbered(
        [
            "Salvar o link do Bookings da unidade.",
            "Fazer login com o usuário correto.",
            "Abrir o Modo Desenvolvedor.",
            "Abrir Editar opções de campos.",
            "Ajustar os campos da unidade.",
            "Salvar os valores.",
            "Fechar e abrir o app novamente para aplicar as alterações.",
            "Fazer um teste antes de usar Reservar.",
        ]
    )

    doc.add_h2("Checklist final da escola")
    doc.add_bullets(
        [
            "Link do Bookings salvo corretamente.",
            "Usuário da unidade logado.",
            "Equipe atualizada.",
            "Público atualizado.",
            "Componente atualizado.",
            "Principal recurso atualizado.",
            "Tipo de atividade atualizado.",
            "Turmas e horários conferidos.",
            "Teste realizado com sucesso.",
        ]
    )

    rgb_pages = [page.convert("RGB") for page in doc.pages]
    rgb_pages[0].save(OUTPUT_PDF, save_all=True, append_images=rgb_pages[1:])
    print(f"PDF gerado em: {OUTPUT_PDF}")


if __name__ == "__main__":
    build_pdf()
