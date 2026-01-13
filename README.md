# 📊 Relatório Acadêmico Automatizado

> Um projeto Python que automatiza a geração de relatórios acadêmicos, separando alunos em grupos de aprovados e reprovados e gerando planilhas Excel formatadas com base em seus desempenhos.

## 📋 Visão Geral

Este projeto demonstra competências em:
- **Manipulação de arquivos Excel** com a biblioteca `openpyxl`
- **Processamento de dados** e lógica de negócio
- **Formatação avançada** de planilhas (cores, fontes, bordas, alinhamento)
- **Reutilização de código** através de funções bem estruturadas
- **Tratamento de exceções** em Python

## 🎯 Funcionalidades

- ✅ Leitura de dados de uma planilha Excel (`alunos.xlsx`)
- ✅ Separação automática de alunos em dois grupos:
  - **Aprovados**: nota final ≥ 7.0
  - **Reprovados**: nota final < 7.0
- ✅ Geração de duas novas planilhas Excel (`aprovados.xlsx` e `reprovados.xlsx`)
- ✅ Formatação profissional com:
  - Títulos com fundo cinza escuro
  - Cabeçalhos coloridos e destacados
  - Bordas e alinhamento centralizado
  - Fontes personalizadas
- ✅ Exibição de estatísticas no terminal:
  - Quantidade de aprovados e reprovados
  - Nota média da turma
  - Nome do aluno com maior nota

## 🚀 Como Usar

### Pré-requisitos

- Python 3.7+
- Biblioteca `openpyxl`
- IDE `Visual Studio Code` (de preferência)

### Instalação

```bash
pip install openpyxl
```

### Executando o Projeto

1. Certifique-se de que o arquivo `alunos.xlsx` está no mesmo diretório que `main.py`
2. Execute o script:

```bash
python main.py
```

3. O programa gerará:
   - `aprovados.xlsx` - planilha com alunos aprovados
   - `reprovados.xlsx` - planilha com alunos reprovados

### Formato do Arquivo de Entrada

O arquivo `alunos.xlsx` deve conter as seguintes colunas:

| Coluna | Tipo | Descrição |
|--------|------|-----------|
| Nome | String | Nome completo do aluno |
| Curso | String | Curso do aluno |
| Idade | Integer | Idade do aluno |
| Nota Final | Float | Nota final (0.0 - 10.0) |
| Data de Matrícula | String/Date | Data de matrícula |

## 💻 Estrutura do Código

### Função Principal: `create_template()`

Cria um template padrão para as planilhas de saída, evitando repetição de código:

```python
def create_template(wb: Workbook, sheet_title: str, title: str, 
                   merge_cells: str, students: Workbook) -> Workbook:
```

**Parâmetros:**
- `wb`: Objeto Workbook para ser formatado
- `sheet_title`: Título da planilha
- `title`: Título exibido na primeira linha mesclada
- `merge_cells`: Intervalo de células a mesclar (ex: "A1:E1")
- `students`: Planilha de origem com os dados dos alunos

**Retorna:**
- Objeto `Workbook` formatado ou `False` em caso de erro

### Principais Bibliotecas Utilizadas

- **openpyxl**: Manipulação de arquivos Excel
  - `Font`: Customização de fontes
  - `PatternFill`: Preenchimento de células
  - `Border` e `Side`: Bordas de células
  - `Alignment`: Alinhamento de texto

## 📊 Exemplo de Saída

**Terminal:**
```
Quantidade de Aprovados: 22
Quantidade de Reprovados: 8
Nota Média da Turma: 7.85
Aluno com Maior Nota: João Silva (9.5)
```

**Planilhas Geradas:**
- Cabeçalhos formatados em cinza
- Dados organizados em colunas
- Bordas em todas as células
- Alinhamento centralizado

## 🛠️ Possíveis Extensões

Este projeto pode ser expandido com:

- 📈 Gráficos de desempenho nas planilhas
- 🔍 Filtros por curso ou faixa de notas
- 📧 Envio automático de relatórios por e-mail
- 📱 Interface gráfica (GUI) com tkinter ou PyQt
- 📁 Suporte a diferentes formatos de entrada (CSV, JSON)
- 🎨 Templates de formatação customizáveis

## 📚 Conceitos Demonstrados

- Programação orientada ao processamento de dados
- Manipulação de objetos complexos (Workbook, Cell)
- Type hints para melhor legibilidade
- Boas práticas de tratamento de exceções
- Documentação de funções com docstrings
- Princípio DRY (Don't Repeat Yourself)

## 📄 Licença

Este projeto é fornecido como está para fins educacionais e de portfólio.

## 👤 Sobre o Autor

Projeto desenvolvido como parte do aprendizado em automação de dados e processamento de planilhas com Python.

---

**⭐ Se este projeto foi útil, considere deixar uma estrela!**
