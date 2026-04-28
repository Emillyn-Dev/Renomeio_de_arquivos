# Renomeio_de_arquivos
Lê o arquivo e renomeia de forma personalizada. 
<br>

Este projeto é um script em Python que automatiza a leitura e organização de arquivos PDF utilizando OCR (Reconhecimento Óptico de Caracteres). Ele extrai informações específicas — como CPF e data — diretamente do conteúdo dos documentos e renomeia os arquivos de forma padronizada. <br>

🚀 Funcionalidades
🔍 Leitura de PDFs (inclusive arquivos escaneados) <br>
🧠 Extração de texto via OCR com suporte a português <br>
🪪 Identificação automática de CPF <br>
📅 Extração de datas no formato dd/mm/aaaa <br>
<br>

🏷️ Renomeação automática dos arquivos no padrão: CPF - DD_MM_AAAA.pdf <br>
⚠️ Tratamento de erros (fallback para OCR em inglês, arquivos já renomeados, etc.)

🛠️ Tecnologias Utilizadas
<ul>
<li> Python </li>
<li> Tesseract OCR </li>
<li> Poppler (para conversão de PDF em imagem) </li>
<li> Bibliotecas: </li>
    <ul>
    <li>pytesseract </li>
    <li>pdf2image </li>
    </ul>
</ul>
<br>

Como Funciona
<ul>
<li> O script percorre todos os arquivos PDF dentro da pasta Documentos. </li>
<li> Cada PDF é convertido em imagens usando o Poppler. </li>
<li> O texto das imagens é extraído com o Tesseract OCR. </li>
<li> O sistema busca padrões de CPF e data usando regex. </li>
<li> Se encontrados, o arquivo é renomeado automaticamente. </li>
</ul>
