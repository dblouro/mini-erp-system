# 📦 Sistema de Gestão de Artigos e Stock

Um programa em **Python** para gestão de artigos, stock, preços e análise de dados. Permite registar, consultar, alterar e eliminar artigos, exportar dados para vários formatos (Word, PDF, Excel), gerar relatórios visuais com **Sweetviz** e analisar dados com **Pandas** e **Matplotlib**.

## 📌 Funcionalidades

### 📂 Gestão de Dados
- Registar, consultar, alterar e eliminar artigos.
- Campos: nome, quantidade, preço, fornecedor, categoria, localização, data de validade, stock mínimo.

### 📤 Exportação de Dados
- Exportar para **Word** (`.docx`).
- Exportar para **PDF** (`.pdf`).
- Exportar para **Excel** (`.xlsx`).
- Gerar ficheiros **Pickle** (`.pk1`).
- Gerar **relatórios Sweetviz** (`.html`).

### 📊 Gestão de Stock
- Calcular **valor total do stock**.
- Verificar **artigos com stock mínimo**.
- **Estatísticas de preços** (média, mediana, desvio padrão).
- **Gráficos** (circular, barras, dispersão).

### 🔍 Análise de Dados
- Filtrar por **categoria**, **fornecedor** ou **localização**.
- Ordenar por **preço**.
- Análise de **stock por localização**.
- Análise de **validade** (dias até expirar).
- **Pesquisa avançada** com múltiplos critérios.

## 🛠 Tecnologias Utilizadas
- **Python** 3.6+
- **Pandas** (análise de dados)
- **Matplotlib** e **Seaborn** (gráficos)
- **Sweetviz** (relatórios visuais)
- **OpenPyXL** (exportação para Excel)
- **Python-DocX** (exportação para Word)
- **FPDF** (exportação para PDF)
- **Turtle** (desenhos interativos)

## 🚀 Como Executar
1. **Clonar o repositório:**
   ```bash
   git clone https://github.com/dblouro/mini-erp-system.git

2. **Navegar até à pagina do projeto:**
   ```bash
   cd mini-erp-system

3. **Instalar dependências:**
   pip install pandas matplotlib seaborn openpyxl python-docx fpdf sweetviz numpy

4. **Executar o programa:**
   python mini-erp-system.py

## 📜 Licença
Este projeto está sob a licença MIT. Consulte o ficheiro LICENSE para mais detalhes.
