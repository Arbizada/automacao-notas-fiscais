
import os
import json
import shutil
import requests
import pymupdf

from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook


# ==========================================
# 1. CONFIGURAÇÕES
# ==========================================

load_dotenv(override=True)

chave_api = os.getenv("GEMINI_API_KEY")

if not chave_api:
    print("ERRO: Chave GEMINI_API_KEY não encontrada no arquivo .env")
    exit()

print("Chave carregada com sucesso!")

pasta_documentos = "Documentos"
pasta_processados = "Processados"

os.makedirs(pasta_documentos, exist_ok=True)
os.makedirs(pasta_processados, exist_ok=True)

arquivo_excel = "notas_fiscais.xlsx"

url = "https://generativelanguage.googleapis.com/v1beta/interactions"

modelo = "gemini-3.5-flash-lite"


# ==========================================
# 2. CABEÇALHOS DO EXCEL
# ==========================================

colunas = [
    "Fornecedor",
    "CNPJ",
    "Número da nota",
    "Data de emissão",
    "Data de vencimento",
    "Cliente",
    "Produtos ou serviços",
    "Subtotal",
    "Impostos",
    "Valor total"
]


# ==========================================
# 3. FUNÇÃO PARA EXTRAIR TEXTO DO GEMINI
# ==========================================

def extrair_texto_gemini(resposta):

    textos = []

    for etapa in resposta.get("steps", []):

        if etapa.get("type") == "model_output":

            conteudo = etapa.get("content", [])

            if isinstance(conteudo, list):

                for item in conteudo:

                    if item.get("type") == "text":
                        textos.append(item.get("text", ""))

            elif isinstance(conteudo, str):
                textos.append(conteudo)

    return "\n".join(textos)


# ==========================================
# 4. FUNÇÃO PARA ORGANIZAR O JSON
# ==========================================

def converter_json(texto):

    texto = texto.strip()

    if texto.startswith("```"):

        linhas = texto.splitlines()

        linhas = [
            linha for linha in linhas
            if not linha.strip().startswith("```")
        ]

        texto = "\n".join(linhas)

    inicio = texto.find("{")
    fim = texto.rfind("}")

    if inicio == -1 or fim == -1:
        raise ValueError("O Gemini não retornou um JSON válido.")

    texto = texto[inicio:fim + 1]

    return json.loads(texto)


# ==========================================
# 5. PROCESSAMENTO DOS PDFs
# ==========================================

arquivos = os.listdir(pasta_documentos)

for arquivo in arquivos:

    if not arquivo.lower().endswith(".pdf"):
        continue

    caminho_pdf = os.path.join(pasta_documentos, arquivo)

    print("\n================================")
    print("Processando:", arquivo)
    print("================================")

    try:

        documento = pymupdf.open(caminho_pdf)

        texto_completo = ""

        for pagina in documento:
            texto_completo += pagina.get_text()

        documento.close()

        if not texto_completo.strip():

            print("O PDF não contém texto extraível.")
            continue

        print("PDF lido com sucesso!")

    except Exception as erro:

        print("Erro ao ler PDF:", erro)
        continue


    # ======================================
    # 6. ENVIO PARA O GEMINI
    # ======================================

    prompt = f"""
    Analise o texto desta nota fiscal e extraia os dados.

    Retorne somente um JSON válido, sem explicações,
    sem comentários e sem blocos Markdown.

    Utilize exatamente estas chaves:

    {{
        "fornecedor": "",
        "cnpj": "",
        "numero_nota": "",
        "data_emissao": "",
        "data_vencimento": "",
        "cliente": "",
        "produtos_servicos": "",
        "subtotal": "",
        "impostos": "",
        "valor_total": ""
    }}

    Não invente informações.
    Se algum dado não estiver disponível, deixe vazio.

    TEXTO DA NOTA FISCAL:

    {texto_completo}
    """

    try:

        resposta_api = requests.post(
            url,
            headers={
                "x-goog-api-key": chave_api,
                "Content-Type": "application/json"
            },
            json={
                "model": modelo,
                "input": prompt
            },
            timeout=120
        )

        resposta_api.raise_for_status()

        dados_resposta = resposta_api.json()

        texto_gemini = extrair_texto_gemini(dados_resposta)

        dados_nota = converter_json(texto_gemini)

        print("Análise do Gemini concluída!")

    except Exception as erro:

        print("Erro na análise do Gemini:", erro)
        continue


    # ======================================
    # 7. ABRIR OU CRIAR PLANILHA
    # ======================================

    try:

        if os.path.exists(arquivo_excel):

            planilha = load_workbook(arquivo_excel)

            if "Notas fiscais" in planilha.sheetnames:
                aba = planilha["Notas fiscais"]
            else:
                aba = planilha.active
                aba.title = "Notas fiscais"

        else:

            planilha = Workbook()
            aba = planilha.active
            aba.title = "Notas fiscais"

            aba.append(colunas)


        # ==================================
        # 8. VERIFICAR DUPLICIDADE
        # ==================================

        cnpj_novo = str(dados_nota.get("cnpj", "")).strip()
        numero_novo = str(dados_nota.get("numero_nota", "")).strip()

        duplicada = False

        for linha in aba.iter_rows(min_row=2, values_only=True):

            cnpj_existente = str(linha[1] or "").strip()
            numero_existente = str(linha[2] or "").strip()

            if (
                cnpj_existente == cnpj_novo
                and numero_existente == numero_novo
            ):

                duplicada = True
                break


        if duplicada:

            print("\nNOTA FISCAL JÁ CADASTRADA!")
            print("CNPJ:", cnpj_novo)
            print("Número da nota:", numero_novo)

        else:

            # ==============================
            # 9. SALVAR DADOS NO EXCEL
            # ==============================

            produtos = dados_nota.get("produtos_servicos", "")

            if isinstance(produtos, (list, dict)):
                produtos = json.dumps(
                    produtos,
                    ensure_ascii=False
                )

            nova_linha = [
                dados_nota.get("fornecedor", ""),
                dados_nota.get("cnpj", ""),
                dados_nota.get("numero_nota", ""),
                dados_nota.get("data_emissao", ""),
                dados_nota.get("data_vencimento", ""),
                dados_nota.get("cliente", ""),
                produtos,
                dados_nota.get("subtotal", ""),
                dados_nota.get("impostos", ""),
                dados_nota.get("valor_total", "")
            ]

            aba.append(nova_linha)

            print("\n================================")
            print("NOTA FISCAL SALVA COM SUCESSO!")
            print("================================")

            print("\nDados identificados:")

            for chave, valor in dados_nota.items():
                print(f"{chave}: {valor}")


        # Salvar e fechar a planilha

        planilha.save(arquivo_excel)
        planilha.close()

        print("\nPlanilha:", os.path.abspath(arquivo_excel))


        # ==================================
        # 10. MOVER PDF PARA PROCESSADOS
        # ==================================

        nome_destino = arquivo

        caminho_destino = os.path.join(
            pasta_processados,
            nome_destino
        )

        contador = 1

        while os.path.exists(caminho_destino):

            nome_base, extensao = os.path.splitext(arquivo)

            nome_destino = f"{nome_base}_{contador}{extensao}"

            caminho_destino = os.path.join(
                pasta_processados,
                nome_destino
            )

            contador += 1

        shutil.move(caminho_pdf, caminho_destino)

        print("PDF movido para:", caminho_destino)


    except Exception as erro:

        print("Erro ao salvar ou verificar Excel:", erro)
        continue


# ==========================================
# 11. FINALIZAÇÃO
# ==========================================

print("\n================================")
print("PROCESSAMENTO FINALIZADO!")
print("================================")
