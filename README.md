# CF_ENEM

## Como será realizado o cálculo da nota geral?

Será realizada uma média aritmética das 5 provas com score máximo de 1000 pontos (cada).

- Redação - até 1000 pontos
- Linguagens - 45 questões valendo 22,22 pontos (cada)
- Ciências da Natureza - 45 questões valendo 22,22 pontos (cada)
- Matemática - 45 questões valendo 22,22 pontos (cada)
- Ciências Humanas - 45 questões valendo 22,22 pontos (cada)

nota_geral = (nota_redacao + nota_linguagens + nota_natureza + nota_matematica + nota_humanas) / 5

OBS. 1: A Teoria de Resposta ao Ítem (TRI) não será considerada nesse cálculo.
OBS. 2: O multiplicador 22.22 é resultado da divisão da nota máxima da prova (1000) pela quantidade de questões (45).

## Como os resultados serão apreentados?

Os resultados serão apresentados em formato de lista contendo o nome do aluno, nota geral e nota por prova.

## De onde os dados serão obtidos?

A través de um arquivo XLSX, gerado pelo aplicativo Evalbee (plataforma de gestão de gabaritos), os dados serão lidos pelo script, que gerará um PDF contendo os resultados já com todas as médias calculadas. O arquivo de teste "simulado.xlsx" ilustra um exemplo.
