import os
import mimetypes
import smtplib
import tempfile
from email.message import EmailMessage
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

load_dotenv()

# Credenciais do VerdeApp
USER_EMAIL = os.getenv("EMAIL")
USER_PASSWORD = os.getenv("SENHA")

# Configuração de envio de e-mail
GMAIL_FROM = os.getenv("GMAIL_FROM")
GMAIL_TO = os.getenv("GMAIL_TO")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD")

LOGIN_URL = "https://verdeapp.verdelog.com.br/auth"
ESTOQUE_URL = "https://verdeapp.verdelog.com.br/estoque"


def validar_variaveis_ambiente():
    faltando = []
    if not USER_EMAIL:
        faltando.append("EMAIL")
    if not USER_PASSWORD:
        faltando.append("SENHA")
    if not GMAIL_FROM:
        faltando.append("GMAIL_FROM")
    if not GMAIL_TO:
        faltando.append("GMAIL_TO")
    if not GMAIL_APP_PASSWORD:
        faltando.append("GMAIL_APP_PASSWORD")
    if faltando:
        raise ValueError("Variáveis ausentes: " + ", ".join(faltando))


def enviar_email_com_anexo_bytes(nome_arquivo: str, conteudo: bytes):
    print("Enviando e-mail com anexo via Gmail...")

    mensagem = EmailMessage()
    mensagem["From"] = GMAIL_FROM
    mensagem["To"] = GMAIL_TO
    mensagem["Subject"] = "Relatório de Estoque - Analítico"

    mensagem.set_content(
        "Olá,\n\n"
        "Segue em anexo o relatório analítico de estoque gerado automaticamente.\n\n"
        "Atenciosamente,\n"
        "Sistema de Automação"
    )

    tipo_mime, _ = mimetypes.guess_type(nome_arquivo)
    if tipo_mime is None:
        tipo_mime = "application/octet-stream"
    maintype, subtype = tipo_mime.split("/")

    mensagem.add_attachment(
        conteudo,
        maintype=maintype,
        subtype=subtype,
        filename=nome_arquivo,
    )

    with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
        smtp.starttls()
        smtp.login(GMAIL_FROM, GMAIL_APP_PASSWORD)
        smtp.send_message(mensagem)

    print(f"E-mail enviado com sucesso para: {GMAIL_TO}")


def baixar_estoque_analitico_e_enviar_email(headless: bool = True):
    validar_variaveis_ambiente()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        # 1) Login
        print("Acessando página de login...")
        page.goto(LOGIN_URL)
        page.fill("#email", USER_EMAIL)
        page.fill("#password", USER_PASSWORD)
        page.click("button[type='submit']")
        page.wait_for_load_state("networkidle")

        # 2) /estoque
        print("Indo para /estoque...")
        try:
            page.wait_for_url("**/estoque", timeout=5000)
        except Exception:
            page.goto(ESTOQUE_URL)
            page.wait_for_load_state("networkidle")

        # 3) Analítico
        print('Clicando em "Analítico"...')
        try:
            page.get_by_role("button", name="Analítico").click()
        except Exception:
            page.click("button:has-text('Analítico')")
        page.wait_for_load_state("networkidle")

        # 4) Exportar Excel e capturar download
        print('Exportando "Excel"...')
        with page.expect_download() as download_info:
            try:
                page.click('button[title="Exportar para Excel (.xlsx)"]')
            except Exception:
                page.get_by_role("button", name="Excel").click()

        download = download_info.value
        nome_arquivo = "estoque_analitico_verdelog.xlsx"

        # Salva em arquivo temporário (runner) só para ler os bytes
        tmp_path = None
        try:
            fd, tmp_path = tempfile.mkstemp(suffix=".xlsx")
            os.close(fd)  # fecha o descritor; Playwright vai escrever no caminho

            download.save_as(tmp_path)

            with open(tmp_path, "rb") as f:
                conteudo = f.read()

        finally:
            if tmp_path and os.path.exists(tmp_path):
                os.remove(tmp_path)

        browser.close()

    # 5) Enviar e-mail com bytes (sem depender de “arquivo salvo”)
    enviar_email_com_anexo_bytes(nome_arquivo, conteudo)


if __name__ == "__main__":
    # No GitHub Actions: headless=True
    baixar_estoque_analitico_e_enviar_email(headless=True)