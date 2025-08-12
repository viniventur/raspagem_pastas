import os
import pandas as pd
import datetime

# Mapeamento das extensões para categorias simplificadas
MAPA_TIPOS = {
    ".pdf": "PDF",
    ".xlsm": "Excel",
    ".xlsx": "Excel",
    ".xls": "Excel",
    ".xltm": "Excel",
    ".ods": "Excel",
    ".docx": "Doc",
    ".doc": "Doc",
    ".pptx": "Apresentação",
    ".txt": "Texto",
    ".csv": "Texto",
    ".ini": "Texto",
    ".jpeg": "Imagem",
    ".jpg": "Imagem",
    ".png": "Imagem",
    ".ai": "Imagem",
    ".mp4": "Vídeo",
    ".pbix": "Power BI",
    ".zip": "Compactado",
    ".crdownload": "Arquivo temporário",
    ".lnk": "Atalho",
    ".kml": "Outro"
}

def classificar_tipo_padronizado(extensao, tipo):
    if tipo == "Pasta":
        return "Pasta"
    return MAPA_TIPOS.get(extensao.lower(), "Outro")

def obter_info_arquivo(caminho):
    try:
        data_modificacao = datetime.datetime.fromtimestamp(os.path.getmtime(caminho))
    except Exception:
        data_modificacao = None

    try:
        tamanho_mb = round(os.path.getsize(caminho) / (1024 * 1024), 2)
    except Exception:
        tamanho_mb = None

    return tamanho_mb, data_modificacao

def raspar_diretorio(diretorio_raiz):
    dados = []

    for raiz, pastas, arquivos in os.walk(diretorio_raiz):
        for pasta in pastas:
            caminho_completo = os.path.join(raiz, pasta)
            pasta_mae = os.path.basename(raiz)
            caminho_pasta = raiz
            tamanho_mb, data_modificacao = obter_info_arquivo(caminho_completo)

            dados.append({
                "Nome do arquivo": pasta,
                "Tipo": "Pasta",
                "Extensão": "Pasta",
                "Tipo Simplificado": "Pasta",
                "Tamanho (MB)": tamanho_mb,
                "Caminho": caminho_completo,
                "Pasta Mãe": pasta_mae,
                "Caminho Pasta": caminho_pasta,
                "Data/Hora da última modificação": data_modificacao
            })

        for arquivo in arquivos:
            if arquivo.endswith(".crdownload"):
                continue

            nome, extensao = os.path.splitext(arquivo)
            caminho_completo = os.path.join(raiz, arquivo)
            pasta_mae = os.path.basename(raiz)
            caminho_pasta = raiz
            tipo_simplificado = classificar_tipo_padronizado(extensao, "Arquivo")
            tamanho_mb, data_modificacao = obter_info_arquivo(caminho_completo)

            dados.append({
                "Nome do arquivo": nome,
                "Tipo": "Arquivo",
                "Extensão": extensao if extensao else "Sem extensão",
                "Tipo Simplificado": tipo_simplificado,
                "Tamanho (MB)": tamanho_mb,
                "Caminho": caminho_completo,
                "Pasta Mãe": pasta_mae,
                "Caminho Pasta": caminho_pasta,
                "Data/Hora da última modificação": data_modificacao
            })

    df = pd.DataFrame(dados)
    return df


if __name__ == "__main__":

    diretorio = r"C:\Users\vcventura\Dropbox\Pasta da equipe Governança Alagoas"
    resultado = raspar_diretorio(diretorio)

    # Salvar como Excel
    nome_arquivo = 'arquivos_listados'
    resultado.to_excel(f"resultados/{nome_arquivo}.xlsx", index=False)


    print(f"Planilha criada com sucesso: {nome_arquivo}.xlsx")
