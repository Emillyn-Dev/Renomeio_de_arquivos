import os
import re
import shutil
import logging
from pdf2image import convert_from_path
import pytesseract

# ==========================
# CONFIGURAÇÕES
# ==========================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PASTA_PDFS = os.path.join(BASE_DIR, "Documentos")
PASTA_SEM_CPF = r"C:\Users\emill\Desktop\Programação\Rede_Dor\Renomear_Documento\Documentos\Sem CPF"
POPPLER_PATH = os.path.join(BASE_DIR, "poppler", "Library", "bin")
TESSERACT_PATH = os.path.join(BASE_DIR, "Tesseract-OCR", "tesseract.exe")
TESSDATA_PATH = os.path.join(BASE_DIR, "Tesseract-OCR", "tessdata")

pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
os.environ["TESSDATA_PREFIX"] = TESSDATA_PATH

logging.basicConfig(level=logging.INFO, format="%(message)s")
log = logging.getLogger(__name__)

# ==========================
# FUNÇÕES
# ==========================
def extrair_texto_pdf(caminho_pdf: str) -> str:
    try:
        paginas = convert_from_path(caminho_pdf, poppler_path=POPPLER_PATH, dpi=300)
    except Exception as e:
        log.error(f"❌ Erro ao abrir PDF: {e}")
        return ""

    texto_completo = ""
    for i, pagina in enumerate(paginas):
        try:
            texto = pytesseract.image_to_string(pagina, lang="por")
        except Exception:
            try:
                log.warning("⚠️ 'por' falhou, usando inglês...")
                texto = pytesseract.image_to_string(pagina, lang="eng")
            except Exception as e:
                log.error(f"❌ Erro OCR página {i}: {e}")
                continue
        texto_completo += texto + "\n"
    return texto_completo


def extrair_cpf(texto: str) -> str | None:
    padrao = r'\d{3}\.?\d{3}\.?\d{3}-?\d{2}'
    resultado = re.search(padrao, texto)
    if resultado:
        cpf = re.sub(r"\D", "", resultado.group())
        return cpf
    return None


def extrair_data(texto: str) -> str | None:
    padrao = r'\d{2}/\d{2}/\d{4}'
    resultado = re.search(padrao, texto)
    if resultado:
        return resultado.group().replace("/", "_")
    return None


def ja_esta_no_formato(nome_arquivo: str) -> bool:
    padrao = r'^\d{11} - \d{2}_\d{2}_\d{4}\.pdf$'
    return re.match(padrao, nome_arquivo) is not None


def mover_para_sem_cpf(caminho: str) -> None:
    destino = os.path.join(PASTA_SEM_CPF, os.path.basename(caminho))
    if os.path.exists(destino):
        log.warning(f"⚠️ Já existe em 'Sem CPF': {os.path.basename(caminho)}")
        return
    shutil.move(caminho, destino)
    log.info(f"📁 Movido para 'Sem CPF': {os.path.basename(caminho)}")


def renomear_arquivo(caminho: str, cpf: str, data: str) -> bool:
    pasta = os.path.dirname(caminho)
    extensao = os.path.splitext(caminho)[1]
    novo_nome = f"{cpf} - {data}{extensao}"
    novo_caminho = os.path.join(pasta, novo_nome)

    if os.path.exists(novo_caminho):
        log.warning(f"⚠️ Já existe: {novo_nome} (original: {os.path.basename(caminho)})")
        return False

    os.rename(caminho, novo_caminho)
    log.info(f"✅ Renomeado: {novo_nome}")
    return True


# ==========================
# EXECUÇÃO
# ==========================
def main():
    log.info("🚀 Iniciando processamento...\n")

    if not os.path.exists(PASTA_PDFS):
        log.error(f"❌ Pasta não encontrada: {PASTA_PDFS}")
        return

    arquivos = [
        f for f in os.listdir(PASTA_PDFS)
        if f.lower().endswith(".pdf") and os.path.isfile(os.path.join(PASTA_PDFS, f))
    ]
    if not arquivos:
        log.warning("⚠️ Nenhum PDF encontrado.")
        return

    total = len(arquivos)
    renomeados = 0
    sem_cpf = 0
    falhas = 0

    for arquivo in arquivos:
        if ja_esta_no_formato(arquivo):
            log.info(f"✔️ Já formatado: {arquivo}")
            continue

        caminho_pdf = os.path.join(PASTA_PDFS, arquivo)
        log.info(f"\n📄 Lendo: {arquivo}")

        texto = extrair_texto_pdf(caminho_pdf)
        if not texto.strip():
            log.error("❌ Não conseguiu extrair texto.")
            falhas += 1
            continue

        cpf = extrair_cpf(texto)
        data = extrair_data(texto)

        if not cpf:
            log.warning(f"⚠️ CPF não encontrado em: {arquivo}")
            mover_para_sem_cpf(caminho_pdf)
            sem_cpf += 1
            continue

        if data:
            if renomear_arquivo(caminho_pdf, cpf, data):
                renomeados += 1
        else:
            log.error(f"❌ Data não encontrada em: {arquivo}")
            falhas += 1

    log.info(f"\n🎉 Finalizado! {renomeados}/{total} renomeados | {sem_cpf} sem CPF | {falhas} falhas")


if __name__ == "__main__":
    main()
