import pandas as pd
from fpdf import FPDF
import logging

# Constantes de nota
NOTA_MIN_LINGUAGENS = 287.0
NOTA_MAX_LINGUAGENS = 820.8
NOTA_MIN_HUMANAS = 289.9
NOTA_MAX_HUMANAS = 823.0
NOTA_MIN_MATEMATICA = 319.8
NOTA_MAX_MATEMATICA = 958.6
NOTA_MIN_NATUREZA = 314.4
NOTA_MAX_NATUREZA = 868.4
NUMERO_QUESTOES_POR_PROVA = 45
TITULO_DOCUMENTO = "Resultados do Simulado ENEM"

# Configuração de logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)


def calcular_multiplicador(area):
    """Calcular o multiplicador baseado nos valores mínimo e máximo das notas."""
    if area == "linguagens":
        return (NOTA_MAX_LINGUAGENS - NOTA_MIN_LINGUAGENS) / NUMERO_QUESTOES_POR_PROVA
    elif area == "humanas":
        return (NOTA_MAX_HUMANAS - NOTA_MIN_HUMANAS) / NUMERO_QUESTOES_POR_PROVA
    elif area == "matematica":
        return (NOTA_MAX_MATEMATICA - NOTA_MIN_MATEMATICA) / NUMERO_QUESTOES_POR_PROVA
    elif area == "natureza":
        return (NOTA_MAX_NATUREZA - NOTA_MIN_NATUREZA) / NUMERO_QUESTOES_POR_PROVA
    else:
        raise ValueError(f"Área desconhecida: {area}")


def calcular_acertos(df, questoes_inicio, questoes_fim):
    """Calcular a soma dos acertos das questões e retornar a pontuação ajustada."""
    colunas_pontuacao = [
        f"Q {i} Marks" for i in range(questoes_inicio, questoes_fim + 1)
    ]
    return df[colunas_pontuacao].sum(axis=1)


def carregar_planilhas(arquivo):
    """Carregar planilhas do arquivo Excel e retornar os DataFrames."""
    try:
        planilhas = pd.read_excel(arquivo, sheet_name=None)
        logging.info(f"Planilhas carregadas com sucesso do arquivo {arquivo}")
        return planilhas
    except Exception as e:
        logging.error(f"Erro ao ler o arquivo Excel: {e}")
        raise RuntimeError(f"Erro ao ler o arquivo Excel: {e}")


def verificar_nomes_consistentes(df_linguagens, df_matematica, df_redacao):
    """Verificar se os nomes dos alunos são consistentes entre as planilhas."""
    try:
        nomes_1 = df_linguagens["Name"]
        nomes_2 = df_matematica["Name"]
        nomes_3 = df_redacao["Name"]

        if not nomes_1.equals(nomes_2) or not nomes_1.equals(nomes_3):
            raise ValueError("Os nomes dos alunos nas planilhas não coincidem.")

        logging.info("Verificação de consistência dos nomes concluída com sucesso.")
        return nomes_1

    except KeyError as e:
        logging.error(f"Coluna ausente na planilha: {e}")
        raise KeyError(f"Coluna ausente na planilha: {e}")


def calcular_notas(df_linguagens_humanas, df_matematica_natureza, df_redacao):
    """Calcular as notas para todas as áreas e a média geral."""
    linguagens_acertos = calcular_acertos(df_linguagens_humanas, 1, 45)
    humanas_acertos = calcular_acertos(df_linguagens_humanas, 46, 90)
    matematica_acertos = calcular_acertos(df_matematica_natureza, 1, 45)
    natureza_acertos = calcular_acertos(df_matematica_natureza, 46, 90)

    linguagens = (
        linguagens_acertos * calcular_multiplicador("linguagens") + NOTA_MIN_LINGUAGENS
    )
    humanas = humanas_acertos * calcular_multiplicador("humanas") + NOTA_MIN_HUMANAS
    matematica = (
        matematica_acertos * calcular_multiplicador("matematica") + NOTA_MIN_MATEMATICA
    )
    natureza = natureza_acertos * calcular_multiplicador("natureza") + NOTA_MIN_NATUREZA
    redacao = df_redacao["Nota Redacao"]

    media_geral = (linguagens + humanas + matematica + natureza + redacao) / 5

    return linguagens, humanas, matematica, natureza, redacao, media_geral


def criar_pdf(
    nomes,
    linguagens,
    humanas,
    matematica,
    natureza,
    redacao,
    media_geral,
    output_path,
    logo_path=None,
):
    """Criar o PDF com os resultados dos simulados, incluindo um logotipo se fornecido."""
    try:
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=13)
        pdf.add_page()

        # Inserir logo, se fornecido
        if logo_path:
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
        for i in range(len(nomes)):
            dados = [
                nomes.iloc[i],
                f"{linguagens.iloc[i]:.1f}",
                f"{humanas.iloc[i]:.1f}",
                f"{matematica.iloc[i]:.1f}",
                f"{natureza.iloc[i]:.1f}",
                f"{redacao.iloc[i]:.1f}",
                f"{media_geral.iloc[i]:.1f}",
            ]
            for dado, largura in zip(dados, largura_colunas):
                pdf.cell(largura, altura_linha, dado, border=1, align="C")
            pdf.ln()

        # Salvar o PDF
        pdf.output(output_path)
        logging.info(f"PDF gerado com sucesso: '{output_path}'")

    except Exception as e:
        logging.error(f"Erro ao criar o PDF: {e}")
        raise RuntimeError(f"Erro ao criar o PDF: {e}")


def main():
    """Função principal para executar o processo."""
    arquivo_simulado = "simulado.xlsx"
    output_path = "Resultados_Simulado_ENEM.pdf"
    logo_path = "logo.png"  # Caminho para o arquivo da logo

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

    # Criar PDF com a logo
    criar_pdf(
        nomes,
        linguagens,
        humanas,
        matematica,
        natureza,
        redacao,
        media_geral,
        output_path,
        logo_path,
    )


if __name__ == "__main__":
    main()
