# Gerador de Relatório DRX

Gera relatórios em PDF de análise de difração de raios X (refinamento de
Rietveld) a partir das exportações do **XRDanalysis PRO** (Anton Paar). Usa
preferencialmente os metadados da medição no arquivo HDF5, mas também gera o
relatório somente com as planilhas `.xlsx`. O relatório é produzido em
português, em páginas A4, com figuras, tabelas e lista de reflexões.

O projeto inclui uma interface gráfica (Tkinter) para uso no dia a dia e uma
interface de linha de comando para automação.

## Recursos

- Relatório completo: resumo, condições de medição, quantificação de fases,
  estruturas cristalinas, índices de ajuste, diagnóstico e apêndice de reflexões.
- Funciona com HDF5 (recomendado) ou somente com as planilhas `.xlsx`.
- Fonte de raios X selecionável: Cu, Co, Cr, Mo ou Ag.
- Gráficos do padrão observado/calculado, fundo, posições de Bragg e diferença.
- Ampliações automáticas do padrão, Rwp por região angular e FWHM das reflexões.
- Consulta ao COD (Crystallography Open Database) para grupo espacial, Z e
  referências ausentes nas exportações, com cache local e modo offline.
- Formatação pt-BR (vírgula decimal, separador de milhar e sinal de menos).
- Fontes Roboto incluídas em `.fonts/`; na ausência delas, usa DejaVu Sans.

## Requisitos

- Python 3.8 ou superior.
- Dependências:

```bash
pip install numpy pandas openpyxl matplotlib reportlab pillow
```

- Para ler arquivos HDF5: `pip install h5py`. Sem HDF5, o relatório pode ser
  gerado apenas com as planilhas e o `h5py` não é necessário.
- Para a interface gráfica, o Tkinter (normalmente já incluído no Python).
  No Linux, se necessário: `sudo apt install python3-tk`.

## Estrutura do projeto

| Arquivo | Descrição |
| --- | --- |
| `gerar_relatorio_drx.py` | Motor de geração do PDF e interface de linha de comando. |
| `interface_drx.py` | Interface gráfica Tkinter. |
| `Info.txt` | Exemplo de informações do laboratório. |
| `.fonts/` | Fontes Roboto (Light, Regular, Medium) e licença SIL OFL. |
| `LICENSE` | Licença MIT do projeto. |

## Como usar

### Interface gráfica

```bash
python interface_drx.py
```

No Windows, o atalho recomendado é `pythonw.exe interface_drx.py` para abrir
sem janela de console.

Na interface:

1. Selecione a pasta da medição.
2. Confirme o arquivo HDF5 ou escolha **— sem HDF5 —** para usar apenas as
   planilhas.
3. Confira a fonte de raios X. Quando o HDF5 informa a fonte, ela é exibida e
   fica bloqueada; sem HDF5, selecione Cu, Co, Cr, Mo ou Ag. Se o campo ficar
   vazio, a fonte será inferida das reflexões do `PeakList`.
4. Edite o nome do analista e os comentários, se necessário. As alterações
   valem apenas para o relatório atual e não modificam o `Info.txt`.
5. Marque **Não consultar o COD (offline)** se estiver sem internet.
6. Clique em **Gerar relatório** e acompanhe o andamento no log.

O PDF é salvo na pasta da medição. A interface oferece botões para abrir o PDF
gerado e a pasta de destino.

### Linha de comando

```bash
python gerar_relatorio_drx.py PASTA
```

Opções:

| Opção | Descrição |
| --- | --- |
| `-o`, `--saida` | Caminho do PDF de saída. |
| `--hdf5` | Arquivo HDF5 da medição (padrão: descoberto na pasta). |
| `--sem-hdf5` | Ignora HDF5 e usa somente as planilhas e o `Info.txt`. |
| `--info` | Caminho do `Info.txt` (padrão: `Info.txt` na pasta dos dados). |
| `--fonte` | Fonte de raios X: `Cu`, `Co`, `Cr`, `Mo` ou `Ag`. |
| `--offline` | Não consulta o COD pela internet. |

Exemplo:

```bash
python gerar_relatorio_drx.py Export -o Relatorio.pdf --offline
python gerar_relatorio_drx.py Export --sem-hdf5 --fonte Cu
```

Sem `--saida`, o arquivo é salvo como
`Relatorio_DRX_<nome_da_varredura>.pdf` na pasta dos dados.

## Arquivos de entrada

Coloque na mesma pasta:

- Uma ou mais planilhas `.xlsx` exportadas pelo XRDanalysis PRO. O nome do
  arquivo não importa: cada planilha é identificada pelo nome das abas.
- O arquivo `.hdf5` (ou `.h5`) da medição. É recomendado e fornece os
  metadados experimentais; sem ele, o relatório usa apenas as planilhas e o
  `Info.txt`.
- Um `Info.txt` (opcional) com dados do laboratório e a fonte de raios X.

Outros arquivos na pasta são ignorados, inclusive PDFs e relatórios já gerados.

### Abas necessárias nas planilhas

As abas podem estar distribuídas em um ou mais arquivos `.xlsx`.

| Aba | Conteúdo usado |
| --- | --- |
| `Quantitative - Phase List` | Fases, átomos e orientação preferencial. |
| `PeakList` | Posições, intensidades, distâncias interplanares e FWHM. |
| `GlobalInstrument` | Condições do instrumento e parâmetros globais. |
| `RValue` | Índices R do refinamento. |
| `Relative Mass` | Frações em massa das fases. |
| `PrimaryChart` | Curvas observada, calculada, fundo e componentes das fases. |
| `Differential Chart` | Curva de diferença. |
| `Chebyshev` | Coeficientes do fundo (opcional). |

Se alguma aba obrigatória não for encontrada, o programa lista as abas
disponíveis e interrompe a geração.

### HDF5 (recomendado)

O HDF5 fornece nome da varredura, amostra, datas, geometria, fendas, detector,
fonte de raios X, parâmetros da varredura e demais metadados experimentais.

Quando há apenas um HDF5 na pasta, ele é usado automaticamente. Com vários
arquivos, o programa tenta associar o nome do HDF5 ao nome da coluna de dados;
se não conseguir, informe o arquivo com `--hdf5` ou selecione-o na interface.
Para ignorar os HDF5 existentes, use `--sem-hdf5` ou a opção **— sem HDF5 —**.

Sem HDF5, o relatório é gerado com as planilhas e o `Info.txt`; as seções que
dependem de metadados ausentes são omitidas e uma observação registra a
limitação. A fonte de raios X continua disponível pela interface, pelo
`Info.txt` (`Fonte:`) ou pela inferência do `PeakList`.

### Info.txt

Arquivo de texto simples, opcional:

```text
XRDynamic 500 - S/N 12345678
XRDanalysis PRO 1.2.5.3658
Fonte: Cu
Analista: Nome do analista

Texto livre após a primeira linha em branco vira comentários no relatório.
Separe os parágrafos com uma linha em branco.
```

- Linha 1: equipamento e número de série.
- Linha 2: software de análise e versão.
- `Fonte: Cu`: fonte de raios X (`Cu`, `Co`, `Cr`, `Mo` ou `Ag`). Também
  aceita `Radiação: Cu Kα` ou `Ânodo: Cu`. É opcional quando há HDF5.
- `Analista: nome`: aparece no rodapé.
- Após a primeira linha em branco: comentários do relatório.

O número de série do instrumento e a versão do software de análise não constam
no HDF5; esses dados só aparecem se o `Info.txt` existir. O motor procura o
`Info.txt` na pasta dos dados; a interface gráfica usa o `Info.txt` da pasta do
programa.

### Fonte de raios X

A fonte é resolvida nesta ordem:

1. `Source target material` do HDF5.
2. Fonte selecionada na interface ou com `--fonte`.
3. `Fonte:` do `Info.txt`.
4. Inferência pelo comprimento de onda calculado das reflexões do `PeakList`.
5. Se nada resolver, a geração é interrompida com uma mensagem pedindo a fonte.

A fonte afeta a identificação da radiação e avisos de consistência; os
resultados quantitativos continuam vindo do refinamento exportado.

## Conteúdo do relatório

1. Página inicial com identificação da amostra, radiação, faixa angular,
   número de fases e índices R.
2. Padrão de difração observado, calculado, fundo, posições de Bragg e
   diferença.
3. Condições de medição: instrumento, fonte de raios X, óptica, detector,
   varredura, parâmetros globais e modelo do refinamento. Campos sem dados são
   omitidos; sem HDF5, permanecem apenas os disponíveis nas planilhas e no
   `Info.txt`.
4. Resultados: quantificação das fases, célula, grupo espacial, átomos,
   qualidade do ajuste (com e sem fundo) e observações automáticas.
5. Comentários informados no `Info.txt` ou na interface.
6. Diagnóstico: ampliações do padrão, Rwp por região e FWHM comparada à
   função instrumental (Caglioti).
7. Apêndice com número, índices hkl, 2θ, d, intensidade relativa e FWHM das
   reflexões.

## COD e modo offline

As referências no formato `99xxxxxxx` das exportações são convertidas para
identificadores do COD. Por padrão, o programa consulta
`https://www.crystallography.net` para obter grupo espacial, Z, nome e
referência bibliográfica ausentes nas planilhas.

As respostas ficam em `cod_cache.json` na pasta do programa e podem ser
reutilizadas sem internet. Use `--offline` (CLI) ou a opção correspondente na
interface para não fazer consultas.

Se o COD não estiver disponível e os dados não estiverem no cache, o relatório
é gerado sem essas informações, com aviso no log e uma observação discreta no
próprio PDF. Para preencher manualmente, edite `INFO_FASES` no início de
`gerar_relatorio_drx.py`:

```python
INFO_FASES = {"LaBO3": {"grupo_espacial": "Pnma (n.º 62)", "Z": 4}}
```

## Configurações

No início de `gerar_relatorio_drx.py`:

- `INFO_FASES`: dados de fases ausentes nas exportações e no COD.
- `LIMITE_APENDICE`: intensidade relativa mínima (%) para listar reflexões no
  apêndice. O padrão é `0.5`; use `0` para listar todas.
- `ANALISTA`: nome padrão do rodapé quando o `Info.txt` não define um.
- `TOL_FONTE`: tolerância em Å para validar/inferir a fonte de raios X.
- `PALETA_FASES` e `COR`: cores das fases e dos elementos gráficos.

## Fontes

As fontes Roboto em `.fonts/` são usadas pelo ReportLab e pelo Matplotlib para
garantir a mesma aparência do PDF em qualquer computador. O programa também
procura fontes em `fonts/`, `.fonts/`, na pasta dos dados, na pasta do usuário
e nas fontes instaladas no sistema. Se não encontrar a Roboto, usa DejaVu Sans.

A licença das fontes está em `.fonts/OFL.txt`.

## Limitações

- As exportações do XRDanalysis PRO não trazem incertezas (desvios-padrão) nem
  indicam quais parâmetros estruturais foram refinados.
- As frações de fase são relativas apenas às fases incluídas no modelo;
  conteúdo amorfo e fases não modeladas não são contabilizados.
- Os coeficientes de Chebyshev exportados são verificados. Se não
  reproduzirem a curva de fundo, o relatório usa a curva exportada e omite a
  lista de coeficientes.
- Sem COD, cache local ou `INFO_FASES`, grupo espacial e Z podem ficar
  ausentes.
- Sem HDF5, metadados como datas da varredura, óptica, detector, tempos e
  condições completas do instrumento podem não aparecer no relatório.
- A fonte de raios X precisa ser informada ou inferível do `PeakList`; caso
  contrário, a geração é interrompida.

## Licença

Código distribuído sob a licença MIT. Consulte `LICENSE`.
