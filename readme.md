# 🧮 Calculadora de Lucro em Python  
### Aplicação integrada com Excel (openpyxl) e interface gráfica Tkinter

Este projeto implementa uma calculadora de lucro com registro de produtos, cálculos automáticos, salvamento em planilha Excel e visualização por meio de uma interface gráfica utilizando Tkinter.

---

## 📌 Funcionalidades

- Cálculo automático de:
  - Lucro líquido  
  - Margem de lucro (%)

- Salvamento dos dados em uma planilha Excel  
- Interface gráfica Tkinter para:
  - Visualizar produtos registrados  
  - Excluir produtos da planilha  

- Manipulação direta da planilha via script

---

## 🗂 Estrutura do Projeto

calculadora-lucro/


│

├── app.py # Interface gráfica (Tkinter)

├── calculadora_lucro.py # Calculadora + salvamento no Excel

├── view.py # Leitura e exclusão direta de dados

├── planilha.xlsx # Gerada automaticamente

└── README.md # Documentação

---

## 🛠 Tecnologias Utilizadas

| Tecnologia  | Descrição |
|-------------|-----------|
| Python 3    | Linguagem base do projeto |
| openpyxl    | Manipulação de planilhas Excel |
| Tkinter / ttk | Interface gráfica desktop |
| Treeview    | Exibição tabular dos produtos |

---

## ▶️ Como Executar

### 1. Instale as dependências

```bash
pip install openpyxl
