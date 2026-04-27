import os
import re
from pdf2image import convert_from_path
import pytesseract

# ==========================
# CONFIGURAÇÕES
# ==========================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 👉 Pasta onde estão os PDFs
PASTA_PDFS = os.path.join(BASE_DIR, "Documentos")

# 👉 Poppler dentro do projeto
POPPLER_PATH = os.path.join(BASE_DIR, "poppler", "Library", "bin")

# 👉 Tesseract dentro do projeto
TESSERACT_PATH = os.path.join(BASE_DIR, "Tesseract-OCR", "tesseract.exe")
TESSDATA_PATH = os.path.join(BASE_DIR, "Tesseract-OCR", "tessdata")

# 👉 Configuração do pytesseract
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
os.environ["TESSDATA_PREFIX"] = TESSDATA_PATH

# 👉 Verificação (evita erro silencioso)
if not os.path.exists(TESSERACT_PATH):
    print(f"❌ Tesseract não encontrado em:\n{TESSERACT_PATH}")
    input("Pressione ENTER para sair...")
    exit()

# ==========================
# FUNÇÕES
# ==========================

def extrair_texto_pdf(caminho_pdf):
    try:
        paginas = convert_from_path(
            caminho_pdf,
            poppler_path=POPPLER_PATH
        )
    except Exception as e:
        print(f"❌ Erro ao abrir PDF: {e}")
        return ""

    texto_completo = ""

    for i, pagina in enumerate(paginas):
        try:
            texto = pytesseract.image_to_string(pagina, lang="por")
        except:
            try:
                print("⚠️ 'por' falhou, usando inglês...")
                texto = pytesseract.image_to_string(pagina, lang="eng")
            except Exception as e:
                print(f"❌ Erro OCR página {i}: {e}")
                continue

        texto_completo += texto + "\n"

    return texto_completo


def extrair_cpf(texto):
    padrao = r'\d{3}\.?\d{3}\.?\d{3}-?\d{2}'
    resultado = re.search(padrao, texto)
    if resultado:
        return re.sub(r"\D", "", resultado.group())
    return None


def extrair_data(texto):
    padrao = r'\d{2}/\d{2}/\d{4}'
    resultado = re.search(padrao, texto)
    if resultado:
        return resultado.group().replace("/", "_")
    return None


def ja_esta_no_formato(nome_arquivo):
    padrao = r'^\d{11} - \d{2}_\d{2}_\d{4}\.pdf$'
    return re.match(padrao, nome_arquivo) is not None


def renomear_arquivo(caminho, cpf, data):
    pasta = os.path.dirname(caminho)
    extensao = os.path.splitext(caminho)[1]  # mantém .pdf

    novo_nome = f"{cpf} - {data}"
    novo_caminho = os.path.join(pasta, novo_nome + extensao)

    if os.path.exists(novo_caminho):
        print(f"⚠️ Já existe: {novo_nome}{extensao}")
    else:
        os.rename(caminho, novo_caminho)
        print(f"✅ Renomeado: {novo_nome}{extensao}")


# ==========================
# EXECUÇÃO
# ==========================

print("🚀 Iniciando processamento...\n")

if not os.path.exists(PASTA_PDFS):
    print(f"❌ Pasta não encontrada: {PASTA_PDFS}")
    exit()

arquivos = os.listdir(PASTA_PDFS)

if not arquivos:
    print("⚠️ Nenhum arquivo encontrado.")
    exit()

for arquivo in arquivos:
    if not arquivo.lower().endswith(".pdf"):
        continue

    if ja_esta_no_formato(arquivo):
        print(f"✔️ Já formatado: {arquivo}")
        continue

    caminho_pdf = os.path.join(PASTA_PDFS, arquivo)

    print(f"\n📄 Lendo: {arquivo}")

    texto = extrair_texto_pdf(caminho_pdf)

    if not texto.strip():
        print("❌ Não conseguiu extrair texto.")
        continue

    cpf = extrair_cpf(texto)
    data = extrair_data(texto)

    if cpf and data:
        renomear_arquivo(caminho_pdf, cpf, data)
    else:
        print(f"❌ CPF ou data não encontrados em: {arquivo}")

print("\n🎉 Finalizado!")