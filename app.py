import pandas as pd
from fpdf import FPDF
import logging

# Configuração do logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

# Parâmetros de nota por área
NOTA_MIN_LINGUAGENS = 287.0
NOTA_MAX_LINGUAGENS = 820.8
NOTA_MIN_HUMANAS = 289.9
NOTA_MAX_HUMANAS = 823.0
NOTA_MIN_MATEMATICA = 319.8
NOTA_MAX_MATEMATICA = 958.6
NOTA_MIN_NATUREZA = 314.4
NOTA_MAX_NATUREZA = 868.4

# Parâmetros da prova
NUMERO_QUESTOES_POR_PROVA = 45

# Parâmetros do PDF
TITULO_DOCUMENTO = "Resultados do Simulado ENEM"
ORDEM = "RD"  # "AC", "AD", "RC", "RD" para ordenar a lista


def calcular_multiplicador(nota_min, nota_max):
    """Calcular o multiplicador baseado nos valores mínimo e máximo das notas."""
    logging.debug(f"Calculando multiplicador para notas entre {nota_min} e {nota_max}.")
    return (nota_max - nota_min) / NUMERO_QUESTOES_POR_PROVA


def calcular_acertos(df, questoes_inicio, questoes_fim):
    """Calcular a soma dos acertos das questões e retornar a pontuação ajustada."""
    logging.debug(
        f"Calculando acertos para questões de {questoes_inicio} a {questoes_fim}."
    )
    colunas_pontuacao = [
        f"Q {i} Marks" for i in range(questoes_inicio, questoes_fim + 1)
    ]
    return df[colunas_pontuacao].sum(axis=1)


def carregar_planilhas(arquivo):
    """Carregar planilhas do arquivo Excel e retornar os DataFrames."""
    logging.info(f"Carregando planilhas do arquivo {arquivo}.")
    try:
        return pd.read_excel(arquivo, sheet_name=None)
    except Exception as e:
        logging.error(f"Erro ao ler o arquivo Excel: {e}")
        raise RuntimeError(f"Erro ao ler o arquivo Excel: {e}")


def verificar_nomes_consistentes(
    df_prova_linguagens_humanas, df_prova_matematica_natureza, df_prova_redacao
):
    """Verificar se os nomes dos alunos são consistentes entre as planilhas."""
    logging.info("Verificando a consistência dos nomes dos alunos entre as planilhas.")
    nomes_1 = df_prova_linguagens_humanas["Name"]
    nomes_2 = df_prova_matematica_natureza["Name"]
    nomes_3 = df_prova_redacao["Name"]

    if not nomes_1.equals(nomes_2) or not nomes_1.equals(nomes_3):
        logging.error("Os nomes dos alunos nas planilhas não coincidem.")
        raise ValueError("Os nomes dos alunos nas planilhas não coincidem.")

    return nomes_1


def calcular_notas(df_linguagens_humanas, df_matematica_natureza, df_redacao):
    """Calcular as notas para todas as áreas e a média geral."""
    logging.info("Calculando as notas para todas as áreas e a média geral.")
    linguagens_acertos = calcular_acertos(df_linguagens_humanas, 1, 45)
    humanas_acertos = calcular_acertos(df_linguagens_humanas, 46, 90)
    matematica_acertos = calcular_acertos(df_matematica_natureza, 1, 45)
    natureza_acertos = calcular_acertos(df_matematica_natureza, 46, 90)

    linguagens = (
        linguagens_acertos
        * calcular_multiplicador(NOTA_MIN_LINGUAGENS, NOTA_MAX_LINGUAGENS)
    ) + NOTA_MIN_LINGUAGENS
    humanas = (
        humanas_acertos * calcular_multiplicador(NOTA_MIN_HUMANAS, NOTA_MAX_HUMANAS)
    ) + NOTA_MIN_HUMANAS
    matematica = (
        matematica_acertos
        * calcular_multiplicador(NOTA_MIN_MATEMATICA, NOTA_MAX_MATEMATICA)
    ) + NOTA_MIN_MATEMATICA
    natureza = (
        natureza_acertos * calcular_multiplicador(NOTA_MIN_NATUREZA, NOTA_MAX_NATUREZA)
    ) + NOTA_MIN_NATUREZA
    redacao = df_redacao["Nota Redacao"]

    media_geral = (linguagens + humanas + matematica + natureza + redacao) / 5

    return linguagens, humanas, matematica, natureza, redacao, media_geral


def ordenar_resultados(df, ordem):
    """Ordenar os resultados conforme a constante ORDEM."""
    logging.info(f"Ordenando os resultados conforme a constante ORDEM: {ordem}.")
    if ordem == "AC":
        df = df.sort_values(by="Name", ascending=True)
    elif ordem == "AD":
        df = df.sort_values(by="Name", ascending=False)
    elif ordem == "RC":
        df = df.sort_values(by="Média Geral", ascending=True)
    elif ordem == "RD":
        df = df.sort_values(by="Média Geral", ascending=False)
    return df


def criar_pdf(df, output_path, logo_path=None):
    """Criar o PDF com os resultados dos simulados, incluindo um logotipo se fornecido."""
    logging.info("Criando o PDF com os resultados dos simulados.")
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=13)
    pdf.add_page()

    # Inserir logo, se fornecido
    if logo_path:
        logging.info(f"Inserindo a logomarca do arquivo {logo_path}.")
        logo_width = 45  # Largura desejada da logomarca
        page_width = pdf.w  # Largura da página
        logo_x = (
            page_width - logo_width
        ) / 2  # Calcula a posição x para centralizar a logomarca
        pdf.image(
            logo_path, x=logo_x, y=8, w=logo_width
        )  # Insere a logomarca centralizada
        pdf.ln(25)  # Espaço adicional após a logo

    # Título
    pdf.set_font("Arial", size=14, style="B")
    pdf.cell(0, 10, TITULO_DOCUMENTO, ln=True, align="C")

    # Definir largura das colunas e altura das linhas
    colunas = [
        "Nome",
        "Linguagens",
        "Humanas",
        "Matemática",
        "Natureza",
        "Redação",
        "Média Geral",
    ]
    largura_colunas = [58, 22, 22, 22, 22, 22, 22]
    altura_linha = 7

    # Adicionar cabeçalhos da tabela
    pdf.set_font("Arial", style="B", size=10)
    for coluna, largura in zip(colunas, largura_colunas):
        pdf.cell(largura, altura_linha, coluna, border=1, align="C")
    pdf.ln()

    # Adicionar dados na tabela
    pdf.set_font("Arial", size=10)
    for index, row in df.iterrows():
        dados = [
            row["Name"],
            f"{row['Linguagens']:.1f}",
            f"{row['Humanas']:.1f}",
            f"{row['Matemática']:.1f}",
            f"{row['Natureza']:.1f}",
            f"{row['Redação']:.1f}",
            f"{row['Média Geral']:.1f}",
        ]
        for dado, largura in zip(dados, largura_colunas):
            pdf.cell(largura, altura_linha, dado, border=1, align="C")
        pdf.ln()

    # Salvar o PDF
    pdf.output(output_path)
    logging.info(f"PDF gerado com sucesso: '{output_path}'")


def main():
    """Função principal para executar o processo."""
    arquivo_simulado = "simulado.xlsx"
    output_path = "Resultados_Simulado_ENEM.pdf"
    logo_path = "logo.png"  # Caminho para o arquivo da logo

    logging.info("Iniciando o processo de geração do PDF.")

    # Carregar dados
    planilhas = carregar_planilhas(arquivo_simulado)
    df_linguagens_humanas = planilhas["Planilha1"]
    df_matematica_natureza = planilhas["Planilha2"]
    df_redacao = planilhas["Planilha3"]

    # Verificar nomes
    nomes = verificar_nomes_consistentes(
        df_linguagens_humanas, df_matematica_natureza, df_redacao
    )

    # Calcular notas
    linguagens, humanas, matematica, natureza, redacao, media_geral = calcular_notas(
        df_linguagens_humanas, df_matematica_natureza, df_redacao
    )

    # Criar DataFrame consolidado
    df_resultados = pd.DataFrame(
        {
            "Name": nomes,
            "Linguagens": linguagens,
            "Humanas": humanas,
            "Matemática": matematica,
            "Natureza": natureza,
            "Redação": redacao,
            "Média Geral": media_geral,
        }
    )

    # Ordenar resultados
    df_resultados = ordenar_resultados(df_resultados, ORDEM)

    # Criar o PDF
    criar_pdf(df_resultados, output_path, logo_path)

    logging.info("Processo concluído com sucesso.")


if __name__ == "__main__":
    main()
