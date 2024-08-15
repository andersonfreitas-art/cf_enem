import pandas as pd
from fpdf import FPDF

NOTA_MIN_LINGUAGENS = 287.0
NOTA_MAX_LINGUAGENS = 820.8
NOTA_MIN_HUMANAS = 289.9
NOTA_MAX_HUMANAS = 823.0
NOTA_MIN_MATEMATICA = 319.8
NOTA_MAX_MATEMATICA = 958.6
NOTA_MIN_NATUREZA = 314.4
NOTA_MAX_NATUREZA = 868.4

NUMERO_QUESTOES_POR_PROVA = 45


def calcular_multiplicador(nota_min, nota_max):
    multiplicador = (nota_max - nota_min) / NUMERO_QUESTOES_POR_PROVA
    return multiplicador


def calcular_acertos(df, questoes_inicio, questoes_fim):
    # Somar os scores das questões corretas
    colunas_pontuacao = [
        f"Q {i} Marks" for i in range(questoes_inicio, questoes_fim + 1)
    ]
    return df[colunas_pontuacao].sum(axis=1)


# Ler todas as planilhas de um arquivo Excel
arquivo_simulado = "simulado.xlsx"

try:
    planilhas = pd.read_excel(arquivo_simulado, sheet_name=None)
except Exception as e:
    raise RuntimeError(f"Erro ao ler o arquivo Excel: {e}")

# Processar cada planilha
df_linguagens_humanas = planilhas["Planilha1"]
df_matematica_natureza = planilhas["Planilha2"]
df_redacao = planilhas["Planilha3"]

# Verificar consistência dos nomes
try:
    nomes_1 = df_linguagens_humanas["Name"]
    nomes_2 = df_matematica_natureza["Name"]
    nomes_3 = df_redacao["Name"]

    if not nomes_1.equals(nomes_2) or not nomes_1.equals(nomes_3):
        raise ValueError("Os nomes dos alunos nas planilhas não coincidem.")
except KeyError as e:
    raise KeyError(f"Coluna ausente na planilha: {e}")

# Calcular notas por área
linguagens = (
    calcular_acertos(df_linguagens_humanas, 1, 45)
    * calcular_multiplicador(NOTA_MIN_LINGUAGENS, NOTA_MAX_LINGUAGENS)
) + NOTA_MIN_LINGUAGENS
humanas = (
    calcular_acertos(df_linguagens_humanas, 46, 90)
    * calcular_multiplicador(NOTA_MIN_HUMANAS, NOTA_MAX_HUMANAS)
) + NOTA_MIN_HUMANAS
matematica = (
    calcular_acertos(df_matematica_natureza, 1, 45)
    * calcular_multiplicador(NOTA_MIN_MATEMATICA, NOTA_MAX_MATEMATICA)
) + NOTA_MIN_MATEMATICA
natureza = (
    calcular_acertos(df_matematica_natureza, 46, 90)
    * calcular_multiplicador(NOTA_MIN_NATUREZA, NOTA_MAX_NATUREZA)
) + NOTA_MIN_NATUREZA
redacao = df_redacao["Nota Redacao"]

# Calcular a média geral
media_geral = (linguagens + humanas + matematica + natureza + redacao) / 5

# Criar o PDF
pdf = FPDF()
pdf.set_auto_page_break(auto=True, margin=13)
pdf.add_page()

# Título
pdf.set_font("Arial", size=14)
pdf.cell(0, 10, "Resultados do Simulado ENEM", ln=True, align="C")

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
largura_colunas = [58, 22, 22, 22, 22, 22, 22]  # Largura das colunas ajustada
altura_linha = 7

# Adicionar cabeçalhos da tabela
pdf.set_font("Arial", style="B", size=10)
for i, coluna in enumerate(colunas):
    pdf.cell(largura_colunas[i], altura_linha, coluna, border=1, align="C")
pdf.ln()

# Adicionar dados na tabela com fonte menor
pdf.set_font("Arial", size=10)
for i in range(len(nomes_1)):
    dados = [
        nomes_1.iloc[i],
        f"{linguagens.iloc[i]:.1f}",
        f"{humanas.iloc[i]:.1f}",
        f"{matematica.iloc[i]:.1f}",
        f"{natureza.iloc[i]:.1f}",
        f"{redacao.iloc[i]:.1f}",
        f"{media_geral.iloc[i]:.1f}",
    ]
    for j, dado in enumerate(dados):
        pdf.cell(largura_colunas[j], altura_linha, dado, border=1, align="C")
    pdf.ln()

# Salvar o PDF
pdf_output_path = "Resultados_Simulado_ENEM.pdf"
pdf.output(pdf_output_path)

print(f"PDF gerado com sucesso: '{pdf_output_path}'")
