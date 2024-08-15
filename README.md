# CF_ENEM

Script escrito em python para gerar uma folha de resultados estilo ENEM à partir de tabelas geradas pelo app EvalBee.

## Como será realizado o cálculo da nota geral?

Será realizada uma média aritmética das 5 provas com score máximo de 1000 pontos (cada).

- Redação - até 1000 pontos
- Linguagens - 45 questões
- Ciências da Natureza - 45 questões
- Matemática - 45 questões
- Ciências Humanas - 45 questões

nota_geral = (nota_redacao + nota_linguagens + nota_natureza + nota_matematica + nota_humanas) / 5

OBS. 1: A Teoria de Resposta ao Ítem (TRI) não será considerada nesse cálculo.

## Quanto vale cada questão?

Os multiplicadores de score para cada questão foram calculados por área à partir dos resultados de nota mínima e máxima do ENEM 2023. Então, ninguém atinge nota 1000 nem 0.

## Como os resultados serão apresentados?

Os resultados serão apresentados em formato de lista contendo o nome do aluno, nota geral e nota por prova.

## De onde os dados serão obtidos?

A través de um arquivo XLSX, gerado pelo aplicativo Evalbee (plataforma de gestão de gabaritos), os dados serão lidos pelo script, que gerará um PDF contendo os resultados já com todas as médias calculadas. O arquivo de teste "simulado.xlsx" ilustra um exemplo.
