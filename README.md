# Conversor Prévia → Planilha Domínio

Programa para Windows que lê a **Prévia da folha** (relatório *Relação de Cálculo* em PDF)
e gera a planilha **"Relação de Valores para Folha de Pagamento"** (dominio planilha)
no mesmo modelo usado pelo escritório. Nada é tirado nem acrescentado: o arquivo gerado
tem a mesma formatação, os mesmos comentários de ajuda, as mesmas colunas e códigos e as
mesmas fórmulas de total. Só os valores mudam.

## Como baixar

1. Abra a aba **Actions** do repositório no GitHub e clique na execução mais recente de
   **"Testes e executável Windows"** que estiver com ✅.
2. No fim da página, em **Artifacts**, baixe **ConversorPrevia-windows** (vem em .zip).
3. Extraia o `ConversorPrevia.exe` para uma pasta sua (por exemplo, `Documentos\Conversor`).
   Não precisa instalar nada.

> Na primeira execução o Windows pode mostrar "O Windows protegeu o computador", porque o
> programa não tem assinatura digital. Clique em **Mais informações → Executar assim mesmo**.

## Como usar

1. Abra o `ConversorPrevia.exe`. Também dá para arrastar um ou mais PDFs sobre o ícone do programa.
2. Clique em **Adicionar PDFs...** e escolha as Prévias (pode escolher várias de uma vez).
3. Na primeira vez que uma empresa aparece, o programa pergunta qual **razão social** vai na
   planilha. Já vem preenchida com a do PDF; altere se precisar (ex.: trocar `EMPRESA EXEMPLO EIRELI - ME` por `EMPRESA EXEMPLO LTDA`).
   A resposta fica guardada para as próximas vezes. Para mudar depois, dê dois cliques na linha.
4. Escolha onde salvar: na mesma pasta de cada PDF, ou numa pasta à sua escolha.
5. Clique em **Gerar planilha(s)**. O arquivo se chama
   `dominio planilha - <empresa> - <mês>-<ano>.xlsx` (ex.: `dominio planilha - 0123 - 08-2026.xlsx`).
   Se já existir um arquivo com esse nome, o programa pergunta antes de substituir.

A área **Resultado** mostra os totais de cada coluna, para conferir com a Prévia, e qualquer aviso.

## Como os dados são preenchidos

| Planilha | Vem do PDF |
|---|---|
| Código Empresa | `Empresa: 0123 - ...` |
| Razão Social | a razão social cadastrada no programa (padrão: a do PDF) |
| Inscrição CNPJ | `Inscrição Federal` |
| Competência | mês/ano do `Período` |
| Tipo de Cálculo (col. A) | `Mensal` = 11 (Adiantamento = 41, Complementar = 42, 13º Adto = 51, 13º Integral = 52) |
| Código / Nome (col. B e C) | linha `Func:` de cada funcionário, em ordem de código |

| Coluna | Código | Evento da Prévia |
|---|---|---|
| Comissão | 37 | 90 Comissões |
| Horas Extras 50% | 150 | evento com descrição "Horas Extras 50%" |
| Prêmio Destaque | 299 | evento com descrição "Prêmio Destaque" |
| Prêmio Mensal | 301 | 932 Premiações Mensais |
| Acréscimos Domingo | 303 | 879 Acréscimo de comissão domingo |
| Prêmio proatividade | 306 | 883 Prêmios Produtividade |
| Prêmio posição / Ranking / desempenho / mês anterior | 309 / 312 / 316 / 318 | evento com a descrição correspondente |
| Complemento VR e VA | 305 | 1306 COMPLEMENTO VR E VA |
| Auxílio Transporte | 259 | 687 Auxílio Vale Transporte |
| Desconto Refeição | 260 | *(nenhum evento; ver observações)* |
| Sindical | 52 | 1860 Contribuição Sindical |
| Refeição | 9999 | 692 e 1211 Desconto de Vale Refeição |

Um evento entra na coluna pelo **código** ou pela **descrição** (sem diferença de acentos ou
maiúsculas). Assim, outras empresas que usam códigos diferentes para o mesmo evento também funcionam.

Não entram na planilha: DSR, salário, diferença de salário, complemento de salário normativo,
faltas, férias e médias de férias, 13º, verbas de rescisão, INSS, IRRF, desconto de vale
transporte, adiantamento e empréstimo. Marque **"Listar eventos do PDF que não vão para a
planilha"** para ver a lista de cada Prévia.

### Observações

- **Férias:** as médias de comissão de férias (357, 364, 661, 672…) não entram na coluna
  Comissão, como na planilha de 07/2026.
- **Demitidos:** aparecem na planilha (código e nome), mas com os valores em branco.
  Para preencher os valores da rescisão, mude `demitidos_sem_valores` para `false` na configuração.
- **Desconto Refeição (260):** na planilha de 07/2026 esta coluna tem valores que não aparecem
  na Prévia do mês. O desconto de vale refeição da Prévia vai para a coluna **Refeição (9999)**,
  como na planilha de 07/2026. Para mudar, veja *Configuração* abaixo.
- Funcionários além das linhas do modelo: o programa acrescenta linhas e ajusta as fórmulas de total.

## Conferências automáticas

Para cada funcionário, o programa confere se a soma dos eventos lidos bate com os totais
**Proventos** e **Descontos** do PDF. Também confere se a quantidade de funcionários lidos é igual
ao **Total Funcionários** do resumo. Se algo não bater, aparece um aviso em laranja.

## Configuração

O botão **Configuração...** abre o arquivo `conversor_config.json`, que fica na mesma pasta do
programa (ou em `%APPDATA%\ConversorPrevia` se a pasta não permitir gravação). Nele ficam:

- `colunas`: para cada código de coluna da planilha, os códigos (`eventos`) e as descrições
  (`descricoes`) dos eventos da Prévia que vão para ela. Exemplo: para levar o desconto de
  refeição para a coluna 260 em vez da 9999, mova `692, 1211` e a descrição de `"9999"` para `"260"`.
- `razao_social`: razão social por código de empresa.
- `nome_arquivo`: padrão do nome do arquivo gerado.

Se apagar o arquivo, o programa cria outro com o padrão.

## Para quem mantém o programa

Requisitos: Python 3.10+.

```bash
pip install -r requirements-dev.txt
python -m pytest                                   # testes (usam uma Prévia fictícia)
python -m conversor                                # abre a janela
python -m conversor previa.pdf --saida pasta/      # converte pela linha de comando
```

- `conversor/leitor_pdf.py`: leitura da Prévia pela posição das palavras na página.
- `conversor/mapeamento.py`: regras evento → coluna e configuração.
- `conversor/planilha.py`: preenche o modelo editando o XML do .xlsx (preserva tudo do modelo).
- `modelo/modelo_dominio.xlsx`: modelo vazio, criado a partir da planilha de 07/2026 com
  `python ferramentas/criar_modelo_limpo.py planilha.xlsx`.

O executável é gerado automaticamente pelo GitHub Actions (`.github/workflows/build.yml`) a cada
envio. Para publicar uma versão na página **Releases**, rode a automação manualmente (Actions → "Testes e executável Windows" → Run workflow) informando a versão (ex.: `v1.0.1`), ou crie uma tag `v*`.

**Privacidade:** este repositório é público. Não envie Prévias, planilhas de clientes nem o
`conversor_config.json`; o `.gitignore` já bloqueia esses arquivos.
