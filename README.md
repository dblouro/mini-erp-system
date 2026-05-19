# 📦 Mini-ERP Management System

[![Python](https://img.shields.io/badge/Python-3.6%2B-blue)](https://www.python.org/)
[![Pandas](https://img.shields.io/badge/Pandas-2.0%2B-red)](https://pandas.pydata.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

---

## 🌍 **Idiomas / Languages**
- [Português](#-português)
- [English](#-english)

---
## 🇬🇧 English
# 📦 Mini-ERP Management System

## 📌 Description

The Mini-ERP Management System is a Python program designed to manage articles, quantities, prices, suppliers, and warehouse locations. This system allows you to:

## 📂 Features

- Register, consult, modify, and delete articles.
- Export data to Word, PDF, Excel, and Pickle.
- Generate visual reports with Sweetviz.
- Analyze data with statistics and graphs (using Pandas, Matplotlib, and Turtle).
- Check minimum stock and expiry dates.
Ideal for small and medium-sized businesses that need a simple and efficient solution for inventory management.


## 🚀 How to run the program

Requirements:
- Python 3.6 or higher.
- Required libraries (see requirements.txt)

1. **Clone repo:**
   ```bash
   git clone https://github.com/dblouro/mini-erp-system.git

2. **Navegate to the folder:**
   ```bash
   cd mini-erp-system

3. **Install depedencies:**
   ```bash
   pip install pandas matplotlib seaborn openpyxl python-docx fpdf sweetviz numpy

4. **Run the program:**
   ```bash
   python mini-erp-system.py

## 📜 License

This project is licensed under the MIT License.

**Author:**
[Diogo Louro](https://www.linkedin.com/in/diogo-louro/)


## 🇵🇹 Português
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
   ```bash
   pip install pandas matplotlib seaborn openpyxl python-docx fpdf sweetviz numpy

4. **Executar o programa:**
   ```bash
   python mini-erp-system.py

### 📌 Exemplo de Output
$ python mini-erp-system.py
==================================================
  BEM-VINDO(A) AO PROGRAMA DE GESTÃO DE ARTIGOS
==================================================
Este programa permite o registo de artigos, quantidade e preço.

 ===== MENU PRINCIPAL =====
 1. Gestão de Dados
 2. Exportação de Dados
 3. Gestão de Stock
 4. Análise de Dados
 0. Terminar programa
==================================================
Indique a opção a executar:  

## 📜 Licença

Este projeto está sob a licença MIT. Consulte o ficheiro LICENSE para mais detalhes.

**Autor:**
[Diogo Louro](https://www.linkedin.com/in/diogo-louro/)

