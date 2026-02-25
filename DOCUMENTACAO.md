# Sistema de Detalhamento de Suportes DXF

## Sumário

1. [Visão Geral](#visão-geral)
2. [Funcionalidades](#funcionalidades)
3. [Arquitetura do Sistema](#arquitetura-do-sistema)
4. [Requisitos](#requisitos)
5. [Instalação](#instalação)
6. [Configuração](#configuração)
7. [Guia de Uso](#guia-de-uso)
8. [Exportação PDF](#exportação-pdf)
9. [Estrutura do Excel](#estrutura-do-excel)
10. [Templates DXF](#templates-dxf)
11. [Solução de Problemas](#solução-de-problemas)

---

## Visão Geral

O **Sistema de Detalhamento de Suportes DXF** é uma aplicação desktop multi-plataforma desenvolvida em Python que automatiza a criação de desenhos de suportes a partir de dados planilhados no Excel. O sistema utiliza a biblioteca `ezdxf` para manipular arquivos DXF diretamente, sem dependência do AutoCAD.

### Fluxo de Trabalho

```
Excel (.xlsx)  →  Sistema  →  ezdxf  →  Arquivos .dxf (+ PDF opcional)
                      ↓
                 PySide6 GUI
                      ↓
                 Relatório de Processamento
```

### Multi-plataforma

O sistema funciona em:
- **Linux** (Ubuntu, Fedora, etc.)
- **Windows** (10 ou superior)
- **macOS** (10.15+)

---

## Funcionalidades

| Funcionalidade | Descrição |
|----------------|-----------|
| **Processamento em Lote** | Gera múltiplos desenhos a partir de uma planilha Excel |
| **Interface Gráfica** | Interface intuitiva com PySide6/Qt |
| **Multi-plataforma** | Funciona em Linux, Windows e macOS |
| **Exportação PDF** | Gera PDFs automaticamente junto com DXF (opcional) |
| **Tratamento de Duplicatas** | Detecta posições duplicadas e adiciona sufixos (_2, _3, etc.) |
| **Barra de Progresso** | Acompanhamento visual do progresso do processamento |
| **Log Detalhado** | Registro completo de todas as operações realizadas |
| **Cancelamento** | Possibilidade de cancelar o processamento a qualquer momento |
| **Relatório Final** | Resumo estatístico com detalhes de erros e alertas |
| **Multi-threading** | Processamento em thread separada para não travar a interface |
| **Independente de AutoCAD** | Não requer AutoCAD instalado |

---

## Arquitetura do Sistema

### Diagrama de Classes

```
┌─────────────────────────────────────────────────────────────┐
│                     MainWindow (QMainWindow)                │
│  - Interface gráfica principal                              │
│  - Seleção de Excel e Templates                             │
│  - Checkbox para geração de PDF                             │
│  - Controle de progresso                                    │
│  - Exibição de logs                                         │
└──────────────────────────┬──────────────────────────────────┘
                           │
                           │ cria
                           ▼
┌─────────────────────────────────────────────────────────────┐
│                      DXFWorker (QThread)                    │
│  - Thread de processamento                                  │
│  - Leitura Excel via pandas                                 │
│  - Manipulação DXF via ezdxf                                │
│  - Conversão PDF via matplotlib                             │
│  - Gerenciamento de estatísticas                            │
└──────────────────────────┬──────────────────────────────────┘
                           │
                           │ usa
                           ▼
┌─────────────────────────────────────────────────────────────┐
│                       Classes de Apoio                       │
│                                                              │
│  ┌─────────────────┐  ┌────────────────┐                    │
│  │ProcessingConfig │  │ProcessingStats │                    │
│  │ - Constantes    │  │ - Contadores   │                    │
│  │ - Colunas       │  │ - Detalhes     │                    │
│  │ - Tags DXF      │  │ - Estatísticas │                    │
│  └─────────────────┘  └────────────────┘                    │
└─────────────────────────────────────────────────────────────┘
```

### Classes Principais

#### `ProcessingConfig`
Centraliza todas as configurações do sistema:
- Colunas obrigatórias do Excel
- Tags de atributos do DXF
- Extensões de arquivo

#### `ProcessingStats`
Gerencia estatísticas do processamento:
- Total de registros
- Sucessos e erros
- Templates não encontrados
- Templates sem atributos
- Posições duplicatas tratadas
- **PDFs gerados e falhados** (novo)
- Detalhes de cada ocorrência

#### `DXFWorker`
Thread de processamento que:
- Lê o arquivo Excel
- Processa cada linha do Excel
- Manipula arquivos DXF via ezdxf
- Converte DXF para PDF (quando habilitado)
- Emite sinais de progresso

#### `MainWindow`
Interface gráfica principal com:
- Seleção de arquivo Excel
- Seleção de pasta de templates
- **Checkbox "Gerar Pdf's"** (novo)
- Botão de processamento
- Barra de progresso
- Área de log

---

## Requisitos

### Sistema Operacional
- **Linux** (Ubuntu 20.04+, Fedora, Debian, etc.)
- **Windows 10** ou superior
- **macOS 10.15+** (Catalina ou superior)

### Python e Dependências

| Dependência | Versão Mínima | Descrição |
|-------------|---------------|-----------|
| Python | 3.8+ | Interpretador Python |
| PySide6 | 6.0+ | Interface gráfica Qt |
| pandas | 1.3+ | Leitura de arquivos Excel |
| openpyxl | 3.0+ | Suporte a .xlsx |
| ezdxf | 1.0+ | Manipulação de arquivos DXF |
| matplotlib | 3.5+ | Exportação de PDF (opcional) |

### Arquivos Necessários

1. **Arquivo Excel** (.xlsx ou .xls) com os dados dos suportes
2. **Pasta de Templates** contendo os arquivos .dxf template
3. **Script Python** `dxf-detalhamento-suportes.py`

---

## Instalação

### 1. Instalar Python

#### Linux
```bash
sudo apt update
sudo apt install python3 python3-pip python3-venv
```

#### Windows
Baixe e instale o Python em [python.org](https://www.python.org/).
Durante a instalação, marque a opção **"Add Python to PATH"**.

#### macOS
```bash
brew install python@3.11
```

### 2. Criar Ambiente Virtual (Recomendado)

#### Linux/macOS
```bash
cd /home/saul/Dev/Autocad_Suportes
python3 -m venv venv
source venv/bin/activate
```

#### Windows
```cmd
cd C:\Users\SeuUsuario\Dev\Autocad_Suportes
python -m venv venv
venv\Scripts\activate
```

### 3. Instalar Dependências

```bash
pip install -r requirements.txt
```

### requirements.txt

```
# Integração Excel-DXF (Multi-plataforma)
# Dependências Python para processamento de arquivos DXF com ezdxf

# GUI Framework
PySide6>=6.0.0

# Processamento de dados
pandas>=1.3.0
openpyxl>=3.0.0

# Biblioteca DXF (substitui pywin32/AutoCAD COM)
ezdxf>=1.0.0

# PDF Export
matplotlib>=3.5.0
```

### 4. Verificar Instalação

```bash
python -c "import ezdxf; import PySide6; print('Instalação OK!')"
```

---

## Configuração

### Configurações de Processamento

As configurações estão definidas na classe `ProcessingConfig` (linhas 16-36):

```python
class ProcessingConfig:
    # Colunas obrigatórias do Excel
    REQUIRED_COLUMNS = [
        'POSICAO', 'TipoSuporte', 'Elevacao',
        'MEDIDA_H', 'MEDIDA_L', 'MEDIDA_M',
        'MEDIDA_H1', 'MEDIDA_H2', 'MEDIDA_L1', 'MEDIDA_L2', 'MEDIDA_B'
    ]

    # Mapeamento de atributos do DXF
    ATTRIBUTE_TAGS = [
        "POSICAO", "TIPOSUPORTE", "ELEVACAO",
        "H", "L", "M", "H1", "H2", "L1", "L2", "B", "DATA_ATUAL"
    ]

    # Extensões de arquivo
    TEMPLATE_EXTENSION = ".dxf"
    OUTPUT_EXTENSION = ".dxf"
```

---

## Guia de Uso

### Passo a Passo

#### 1. Iniciar a Aplicação

```bash
source venv/bin/activate  # Linux/macOS
# ou
venv\Scripts\activate     # Windows

python dxf-detalhamento-suportes.py
```

#### 2. Selecionar o Arquivo Excel

1. Clique no botão **"Selecionar Arquivo Excel"**
2. Navegue até o local do seu arquivo .xlsx
3. Selecione o arquivo e clique em "Abrir"

#### 3. Selecionar a Pasta de Templates

1. Clique no botão **"Selecionar Pasta de Templates"**
2. Navegue até a pasta contendo os arquivos .dxf template
3. Clique em "Selecionar Pasta"

#### 4. (Opcional) Habilitar Geração de PDF

1. Marque o checkbox **"Gerar Pdf's"** se desejar criar arquivos PDF
2. Os PDFs serão salvos na pasta `Pdf/` criada automaticamente

#### 5. Processar os Dados

1. O botão **"Processar Dados"** será habilitado automaticamente
2. Clique nele para iniciar o processamento
3. Acompanhe o progresso na barra de progresso e no log

#### 6. Resultados

Ao final, você receberá:
- **Arquivos .dxf** na mesma pasta do Excel
- **Arquivos .pdf** (se habilitado) na pasta `Pdf/`
- **Relatório detalhado** no log
- **Mensagem de resumo** com estatísticas

---

## Exportação PDF

### Ativando a Exportação PDF

Para gerar arquivos PDF junto com os DXFs:

1. Marque o checkbox **"Gerar Pdf's"** na interface principal
2. Execute o processamento normalmente

### Como Funciona

```
DXF gerado → ezdxf lê o arquivo → matplotlib renderiza → PDF salvo
```

O processo:
1. Após cada DXF ser criado com sucesso
2. O sistema lê o DXF usando `ezdxf`
3. Renderiza o PaperSpace usando `matplotlib`
4. Salva como PDF na pasta `Pdf/`

### Configurações PDF

| Configuração | Valor |
|--------------|-------|
| Tamanho da página | A4 (8.27 x 11.69 pol) |
| Resolução (DPI) | 300 |
| Layout processado | Primeiro PaperSpace encontrado |
| Pasta de saída | `Pdf/` (ao lado do Excel) |

### Estatísticas de PDF

O relatório final inclui:
- **PDFs gerados com sucesso**: Quantidade de PDFs criados
- **PDFs falhados**: Quantidade de PDFs com erro
- **Detalhes de PDFs falhados**: Lista de arquivos que falharam

### Log de PDF

```
[1/5] ✅ POS-001.dxf (9 atribs)
      📄 POS-001.pdf criado
[2/5] ✅ POS-002.dxf (9 atribs)
      📄 POS-002.pdf criado
[3/5] ✅ POS-003.dxf (9 atribs)
      ⚠️ PDF falhou: Nenhum layout PaperSpace encontrado
```

### Limitações

- Apenas o **primeiro layout PaperSpace** é renderizado
- O ModelSpace é **ignorado**
- Desenhos muito complexos podem ter tempo de renderização maior
- A qualidade é adequada para visualização e impressão básica
- Para qualidade profissional extrema, considere usar o AutoCAD ou LibreCAD

---

## Estrutura do Excel

### Colunas Obrigatórias

O arquivo Excel deve conter as seguintes colunas:

| Coluna | Descrição | Exemplo | Observação |
|--------|-----------|---------|------------|
| POSICAO | Identificação única da posição | POS-001, SUP-A-01 | Usado como nome do arquivo |
| TipoSuporte | Nome do template a usar | SUP-TIPO-01 | Deve corresponder a um arquivo .dxf |
| Elevacao | Altura de instalação | +5,50 | Vírgula é convertida para ponto |
| MEDIDA_H | Medida horizontal principal | 500 | Vazio = "-" |
| MEDIDA_L | Medida longitudinal | 300 | Vazio = "-" |
| MEDIDA_M | Medida média/secundária | 200 | Vazio = "-" |
| MEDIDA_H1 | Medida H1 (alternativa) | 400 | Vazio = "-" |
| MEDIDA_H2 | Medida H2 (alternativa) | 300 | Vazio = "-" |
| MEDIDA_L1 | Medida L1 (alternativa) | 200 | Vazio = "-" |
| MEDIDA_L2 | Medida L2 (alternativa) | 150 | Vazio = "-" |
| MEDIDA_B | Medida de base/largura | 100 | Vazio = "-" |

### Formato da Elevação

- **Correto:** `+5,50` ou `5.50` ou `5,50`
- O sistema converte vírgula para ponto automaticamente

### Células Vazias

- Células vazias são tratadas como `"-"` (hífen)
- O atributo no DXF receberá o valor `"-"`

### Exemplo de Planilha

```
┌────────────────────────────────────────────────────────────────────────┐
│  POSICAO    │ TipoSuporte │ Elevacao │ H   │ L   │ M   │ H1  │ B   │
├────────────────────────────────────────────────────────────────────────┤
│  POS-001    │ SUP-A       │ +5,50    │ 500 │ 300 │     │     │ 100 │
│  POS-002    │ SUP-B       │ +6,00    │     │     │     │ 400 │ 120 │
│  POS-003    │ SUP-C       │ 4.50     │ 600 │ 400 │ 250 │     │     │
└────────────────────────────────────────────────────────────────────────┘
```

---

## Templates DXF

### Estrutura do Template

Os arquivos .dxf template devem conter:

1. **Blocos com Atributos**: Blocos do tipo `INSERT` com atributos
2. **Tags de Atributo**: Os atributos devem ter as tags corretas
3. **PaperSpace**: O layout deve estar no PaperSpace (não ModelSpace)

### Tags de Atributo

| Tag | Descrição | Origem |
|-----|-----------|--------|
| POSICAO | Posição/Identificação | Coluna POSICAO |
| TIPOSUPORTE | Tipo de suporte | Coluna TipoSuporte |
| ELEVACAO | Elevação de instalação | Coluna Elevacao |
| H | Medida H | Coluna MEDIDA_H |
| L | Medida L | Coluna MEDIDA_L |
| M | Medida M | Coluna MEDIDA_M |
| H1 | Medida H1 | Coluna MEDIDA_H1 |
| H2 | Medida H2 | Coluna MEDIDA_H2 |
| L1 | Medida L1 | Coluna MEDIDA_L1 |
| L2 | Medida L2 | Coluna MEDIDA_L2 |
| B | Medida B | Coluna MEDIDA_B |
| DATA_ATUAL | Data atual (automático) | Data do sistema |

### Criando um Template

#### Via AutoCAD
1. Crie o desenho no AutoCAD
2. Crie um bloco com atributos
3. Coloque o bloco no PaperSpace
4. Salve como DXF (`File > Save As > DXF`)

#### Via LibreCAD (gratuito)
1. Crie o desenho no LibreCAD
2. Use o comando para criar blocos com atributos
3. Salve como DXF

#### Via ezdxf (programático)
```python
import ezdxf

doc = ezdxf.new('R2010')
msp = doc.modelspace()
# ... criar desenho ...
doc.saveas('template.dxf')
```

### Exemplo de Bloco com Atributos

```
┌─────────────────────────────┐
│  POSICAO: [POS-001]         │
│  TIPO: [SUP-TIPO-01]        │
│  ELEV: [+5.50]              │
│  H: [500]  L: [300]  B:[100]│
│  DATA: [25/02/2026]         │
└─────────────────────────────┘
```

---

## Solução de Problemas

### Erro: "Colunas faltando no Excel"

**Causa:** O Excel não contém todas as colunas obrigatórias.

**Solução:** Verifique se o Excel contém todas as colunas listadas em [Colunas Obrigatórias](#colunas-obrigatórias).

### Erro: "Template não encontrado"

**Causa:** O arquivo .dxf template não existe na pasta selecionada.

**Solução:**
1. Verifique o nome do template no Excel
2. Verifique se o arquivo .dxf existe na pasta de templates
3. Os nomes devem ser idênticos (incluindo maiúsculas/minúsculas no Linux)

### Aviso: "Sem atributos encontrados"

**Causa:** O template não possui blocos com atributos no PaperSpace.

**Solução:**
1. Abra o template em um editor CAD
2. Verifique se existe um bloco com atributos
3. Certifique-se de que o bloco está no PaperSpace, não no ModelSpace

### Erro: "PDF falhou: Nenhum layout PaperSpace encontrado"

**Causa:** O DXF não possui layouts PaperSpace (apenas ModelSpace).

**Solução:**
1. Abra o template e crie um layout PaperSpace
2. Coloque o conteúdo a ser renderizado no PaperSpace
3. Salve novamente

### PDF gerado está em branco

**Causa:** O conteúdo do DXF está apenas no ModelSpace.

**Solução:**
Mova o conteúdo para o PaperSpace ou crie um viewport no PaperSpace.

### O Processamento Está Muito Lento

**Soluções:**
1. Desabilite a geração de PDF se não for necessária
2. Feche outros aplicativos pesados
3. Divida o Excel em arquivos menores

### Erro ao importar matplotlib

**Causa:** matplotlib não está instalado.

**Solução:**
```bash
pip install matplotlib
```

---

## Exemplos de Funcionamento

### EXEMPLO 1: Processamento Simples

**Cenário:** Processar 3 suportes sem PDF.

#### Arquivo Excel

| POSICAO | TipoSuporte | Elevacao | MEDIDA_H | MEDIDA_L | MEDIDA_B |
|---------|-------------|----------|----------|----------|----------|
| POS-001 | SUP-A | +5,50 | 500 | 300 | 100 |
| POS-002 | SUP-B | +6,00 | 600 | 400 | 120 |
| POS-003 | SUP-A | +5,50 | 500 | 300 | 100 |

#### Log de Processamento

```
=== INÍCIO DO PROCESSAMENTO ===
Data/Hora: 25/02/2026 10:30:15
Arquivo Excel: /home/saul/Projetos/dados.xlsx
Pasta de Templates: /home/saul/Templates
--------------------------------------------------
Lendo arquivo Excel...
Processando 3 registros em 2 grupo(s).

==================================================
TEMPLATE: SUP-A.dxf (2 docs)
==================================================
[1/3] ✅ POS-001.dxf (9 atribs)
[2/3] ✅ POS-003.dxf (9 atribs)

==================================================
TEMPLATE: SUP-B.dxf (1 docs)
==================================================
[3/3] ✅ POS-002.dxf (9 atribs)

===== PROCESSAMENTO CONCLUÍDO =====

==================================================
RELATÓRIO FINAL DE PROCESSAMENTO
--------------------------------------------------
Total de registros processados: 3
Arquivos criados com sucesso: 3
Templates não encontrados: 0
Templates sem atributos: 0
Posicoes duplicatas tratadas: 0
Erros durante o processamento: 0

==================================================
Processamento finalizado em: 25/02/2026 10:30:20
```

---

### EXEMPLO 2: Processamento com PDF

**Cenário:** Mesmo dados anteriores, mas com PDF habilitado.

#### Log de Processamento

```
=== INÍCIO DO PROCESSAMENTO ===
Checkbox "Gerar Pdf's": marcado
--------------------------------------------------

[1/3] ✅ POS-001.dxf (9 atribs)
      📄 POS-001.pdf criado
[2/3] ✅ POS-002.dxf (9 atribs)
      📄 POS-002.pdf criado
[3/3] ✅ POS-003.dxf (9 atribs)
      📄 POS-003.pdf criado

===== PROCESSAMENTO CONCLUÍDO =====

==================================================
RELATÓRIO FINAL DE PROCESSAMENTO
--------------------------------------------------
Total de registros processados: 3
Arquivos criados com sucesso: 3
PDFs gerados com sucesso: 3
PDFs falhados: 0
==================================================
```

#### Arquivos Gerados

```
/home/saul/Projetos/
├── dados.xlsx         (original)
├── POS-001.dxf
├── POS-002.dxf
├── POS-003.dxf
└── Pdf/               (nova pasta)
    ├── POS-001.pdf
    ├── POS-002.pdf
    └── POS-003.pdf
```

---

### EXEMPLO 3: Tratamento de Duplicatas

**Cenário:** Planilha com posições duplicadas.

#### Arquivo Excel

| POSICAO | TipoSuporte | Elevacao | MEDIDA_H | MEDIDA_L |
|---------|-------------|----------|----------|----------|
| POS-101 | SUP-A | +4,00 | 400 | 300 |
| POS-102 | SUP-B | +4,00 | 500 | 400 |
| POS-101 | SUP-A | +4,50 | 400 | 300 |
| POS-101 | SUP-B | +5,50 | 500 | 400 |

#### Log de Processamento

```
[1/4] ✅ POS-101.dxf (9 atribs)
[2/4] ✅ POS-102.dxf (9 atribs)
[3/4] ✅ POS-101_2.dxf (9 atribs)
[4/4] ✅ POS-101_3.dxf (9 atribs)

==================================================
RELATÓRIO FINAL DE PROCESSAMENTO
--------------------------------------------------
Posicoes duplicatas tratadas: 2

Detalhes de Posicoes Duplicatas:
  - POS-101 -> POS-101_2
  - POS-101 -> POS-101_3
==================================================
```

---

## Tabela de Referência Rápida

### Atalhos da Interface

| Ação | Descrição |
|-------------|------|
| Botão "Selecionar Arquivo Excel" | Abre diálogo para selecionar .xlsx |
| Botão "Selecionar Pasta de Templates" | Abre diálogo para selecionar pasta |
| Checkbox "Gerar Pdf's" | Habilita/desabilita exportação PDF |
| Botão "Processar Dados" | Inicia o processamento |
| Botão "Cancelar Processamento" | Interrompe o processamento |

### Extensões de Arquivo

| Extensão | Uso |
|----------|-----|
| .xlsx / .xls | Arquivo de entrada (dados) |
| .dxf | Template e arquivo de saída |
| .pdf | Arquivo de saída (opcional) |

---

## Changelog

### Versão 3.0 (25/02/2026)
- **NOVO:** Exportação opcional para PDF
- **NOVO:** Checkbox "Gerar Pdf's" na interface
- **NOVO:** Estatísticas de PDF no relatório final
- **MELHORIA:** Sistema multi-plataforma (Linux/Windows/macOS)
- **MELHORIA:** Substituição de AutoCAD COM por ezdxf

### Versão 2.0 (24/02/2026)
- **NOVO:** Tratamento de posições duplicadas com sufixos automáticos
- **NOVO:** Estatísticas de duplicatas no relatório final
- **MELHORIA:** Log mais detalhado com nome do arquivo de saída

### Versão 1.0
- Versão inicial do sistema

---

## Contato e Suporte

Para dúvidas ou sugestões sobre o sistema, consulte a documentação técnica ou entre em contato com a equipe de desenvolvimento.

---

*Documento atualizado em 25/02/2026*
