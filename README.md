<div align="center">

# EMS PLUVI VBA

**Automação em Excel para organizar e resumir séries históricas de precipitação da ANA.**

![VBA](https://img.shields.io/badge/VBA-Excel-217346?logo=microsoft-excel&logoColor=white)
![Hidrologia](https://img.shields.io/badge/aplicação-hidrologia-002060)
![Status](https://img.shields.io/badge/status-em_desenvolvimento-EF7726)

</div>

## Problema

Arquivos pluviométricos exportados de sistemas da Agência Nacional de Águas e Saneamento Básico (ANA) exigem várias etapas manuais antes de uma análise: importação do TXT, limpeza de campos, tratamento de datas, identificação de falhas e consolidação dos totais mensais e anuais.

O **EMS PLUVI VBA** reúne macros experimentais para automatizar parte desse fluxo no Microsoft Excel.

## Fluxo automatizado

```text
Arquivo TXT da ANA
        ↓
Importação e cópia dos dados
        ↓
Remoção de colunas auxiliares
        ↓
Criação dos campos mês e ano
        ↓
Identificação de valores ausentes
        ↓
Totais mensais, anuais e médias
        ↓
Tabela de resultados para análise
```

## Módulos disponíveis

| Arquivo | Função principal |
|---|---|
| `Processar_TXT` | Coordena a seleção, abertura e sequência geral de processamento |
| `CopiarDados_do_TXT` | Copia a base importada para a planilha `dados` |
| `ExcluirColunas_TXT` | Remove campos auxiliares não utilizados no resumo |
| `Calcular_Mes_ano_TXT` | Cria variáveis de mês e ano a partir das datas |
| `Processa_Resultados_TXT` | Estrutura a tabela de resultados, identifica falhas e calcula resumos |

## Requisitos

- Microsoft Excel para Windows com suporte a VBA.
- Arquivo pluviométrico TXT em estrutura compatível com a exportação utilizada no desenvolvimento.
- Macros habilitadas em ambiente confiável.
- Conhecimento básico de Excel/VBA para integrar os módulos a uma pasta de trabalho.

## Uso atual

1. Crie uma cópia de segurança do arquivo de dados.
2. Importe os módulos para uma pasta de trabalho habilitada para macros (`.xlsm`).
3. Confirme os nomes esperados das planilhas: `dados` e `resultados`.
4. Revise os nomes e posições das colunas do TXT antes da execução.
5. Execute a rotina principal `AjustarArquivo` em um arquivo de teste.
6. Compare os resultados com uma conferência manual antes de utilizá-los.

> [!CAUTION]
> Este repositório é um protótipo em desenvolvimento. Algumas rotinas chamadas por `AjustarArquivo` ainda não estão publicadas e o código depende da estrutura específica do arquivo de origem. Não utilize os resultados sem validação independente.

## Limitações conhecidas

- Dependência de posições e nomes específicos de colunas.
- Uso de seleções de células em algumas rotinas, reduzindo portabilidade e desempenho.
- Ausência de testes automatizados e arquivo demonstrativo anonimizado.
- Valores ausentes são representados internamente por `-99999` em parte do fluxo.
- A rotina principal referencia módulos complementares ainda não incluídos no repositório.

## Roadmap

- [ ] Publicar uma pasta de trabalho demonstrativa com dados fictícios.
- [ ] Reunir as rotinas em módulos `.bas` documentados.
- [ ] Remover dependências de `Select`, `Selection` e `ActiveSheet`.
- [ ] Validar automaticamente o formato do arquivo de entrada.
- [ ] Adicionar testes com séries contendo falhas.
- [ ] Migrar o processamento para Python/Pandas, mantendo exportação para Excel.

## Autor

**Emanuel Gomes Soares** - Engenheiro Ambiental e pesquisador em Recursos Hídricos<br>
[LinkedIn](https://www.linkedin.com/in/emanuelgomessoares/) · [GitHub](https://github.com/emanuel2806pb)
