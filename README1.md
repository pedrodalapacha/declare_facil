<p align="center">
  <img src="imagens/logo.png" alt="Logo Declare Fácil" width="120">
</p>

<h1 align="center">Declare Fácil</h1>

<p align="center">
  <strong>Agregador de dados para a declaração do Imposto de Renda (IRPF), feito 100% no Excel.</strong><br>
  Projeto desenvolvido no bootcamp <em>Santander – Excel com IA (Claude)</em> da <a href="https://www.dio.me/">DIO</a>.
</p>

<p align="center">
  <img alt="Excel" src="https://img.shields.io/badge/Microsoft%20Excel-005741?style=for-the-badge&logo=microsoftexcel&logoColor=white">
  <img alt="IRPF" src="https://img.shields.io/badge/IRPF-organizado-478C27?style=for-the-badge">
  <img alt="DIO" src="https://img.shields.io/badge/DIO-Santander-31731F?style=for-the-badge">
  <img alt="Claude" src="https://img.shields.io/badge/IA-Claude-025930?style=for-the-badge">
</p>

---

## 📌 Sumário

- [Sobre o projeto](#-sobre-o-projeto)
- [O desafio](#-o-desafio)
- [Telas](#-telas)
- [Estrutura da planilha](#-estrutura-da-planilha)
- [Validações automáticas](#-validações-automáticas)
- [Navegação e links rápidos](#-navegação-e-links-rápidos)
- [Identidade visual](#-identidade-visual)
- [Como usar](#-como-usar)
- [Como personalizar](#-como-personalizar)
- [Recursos do Excel utilizados](#-recursos-do-excel-utilizados)
- [Aprendizados](#-aprendizados)
- [Próximos passos](#-próximos-passos)
- [Autor](#-autor)

---

## 💡 Sobre o projeto

Todo ano é a mesma história: na hora de declarar o Imposto de Renda, os dados estão espalhados por e-mails, PDFs de bancos, holerites e anotações soltas.

O **Declare Fácil** junta tudo em um único arquivo Excel com cara de aplicativo: um **menu lateral de navegação**, **campos validados automaticamente** e **links rápidos**, para que, quando chegar a hora de abrir o programa da Receita Federal, todas as informações já estejam organizadas, conferidas e em um só lugar.

> ⚠️ **Aviso:** o Declare Fácil é uma ferramenta de **organização** pessoal. Ele não substitui o programa oficial da Receita Federal nem o acompanhamento de um contador. Todos os dados de exemplo da planilha são fictícios.

---

## 🎯 O desafio

> Criar uma ferramenta no Excel que ajude a organizar e reunir informações essenciais para a declaração de imposto de renda. A proposta é construir um agregador de dados no qual o usuário possa controlar suas entradas de maneira eficiente e validada, com menus de navegação, validações automáticas e funcionalidades extras, como links rápidos. A solução será completamente construída no Excel, com recursos que tornam a ferramenta robusta, mas com uma interface amigável e prática.

**O que foi entregue:**

| Requisito do desafio | Como foi atendido no Declare Fácil |
|---|---|
| Agregador de dados para o IR | 3 seções: dados do titular, informes bancários e entradas mensais |
| Entradas controladas e validadas | Validação de dados em todos os campos de preenchimento, com mensagens de ajuda e de erro em português |
| Menus de navegação | Menu lateral fixo com botões clicáveis em todas as abas + botões **Anterior / Próximo** |
| Validações automáticas | Verificação dos dígitos do CPF, e-mail, datas, CEP, telefones, valores e listas suspensas |
| Funcionalidades extras / links rápidos | Link dinâmico de e-mail, link para abrir os PDFs dos informes, link para o LinkedIn do autor |
| 100% Excel | Sem macros (VBA): o arquivo é um `.xlsx` comum e abre sem alertas de segurança |

---

## 🖼️ Telas

### 1. Titular
![Tela Titular](imagens/tela-titular.png)

### 2. Informes de rendimentos bancários
![Tela Informes](imagens/tela-informes.png)

### 3. Notas bancárias / extrato de holerites
![Tela Notas](imagens/tela-notas.png)

---

## 🗂️ Estrutura da planilha

O arquivo `DeclareFacil.xlsx` tem **4 abas**:

```
DeclareFacil.xlsx
├── TITULAR    → dados cadastrais da pessoa física
├── INFORMES   → saldos dos bancos + anexos dos informes
├── NOTAS      → entradas mês a mês (holerite, CNPJ, freelance)
└── TABELAS    → listas auxiliares (aba oculta)
```

Todas as abas visíveis seguem o mesmo layout:

- **Coluna A:** menu lateral (logo, nome do app, botões de navegação, assinatura e LinkedIn);
- **Linha 3:** título da seção;
- **Linha 4:** faixa com a instrução de preenchimento;
- **Coluna C:** rótulos dos campos;
- **Coluna D:** células de preenchimento (fundo verde-claro `#E6F3F2`);
- **Coluna E:** dicas de preenchimento, status de validação e links rápidos.

### 1️⃣ TITULAR — Dados do titular

| Campo | Formato / regra |
|---|---|
| NOME | Texto de 3 a 150 caracteres |
| CPF | Apenas números, exibido como `000.000.000-00`, com **checagem dos dígitos verificadores** |
| NASCIMENTO | Data `dd/mm/aaaa`, entre 01/01/1900 e hoje |
| TÍTULO DE ELEITOR | Apenas números, exibido como `0000 0000 0000` |
| CÔNJUGE | Texto (opcional) |
| RUA | Endereço completo (até 150 caracteres) |
| RUA ABREVIADA | Versão curta do endereço (**máximo 40 caracteres**) |
| CEP | 8 dígitos, exibido como `00000-000` |
| TELEFONE | DDD + 8 dígitos, exibido como `(00) 0000-0000` |
| CELULAR | DDD + 9 dígitos, exibido como `(00) 00000-0000` |
| E-MAIL | Precisa ter `@`, domínio com ponto e não pode ter espaços |
| HOUVE ALTERAÇÕES DA ENTREGA ANTERIOR | Lista: `SIM` / `NÃO` |
| DEPENDENTE CÔNJUGE | Lista: `SIM` / `NÃO` |
| RESIDENTE DO EXTERIOR | Lista: `SIM` / `NÃO` |

### 2️⃣ INFORMES — Informes de rendimentos bancários

- Bloco **TOTAL** em destaque, somando automaticamente o valor atual dos 3 bancos: `=SUM(D11,D16,D21)`;
- Para cada banco (1º, 2º e 3º):
  - **BANCO:** lista suspensa com 50 instituições no formato `código - nome` (ex.: `336 - C6 Bank`);
  - **VALOR ATUAL:** saldo em 31/12, em formato `R$`;
  - **ANEXO 🖇️:** nome do PDF do informe, com link **📂 abrir anexo** ao lado.

### 3️⃣ NOTAS — Notas bancárias ou extrato de holerites

- Tabela do Excel chamada **`Entradas`** (intervalo `C8:E35`) com as colunas **DATA**, **CATEGORIA** e **VALOR**;
- A categoria é escolhida em uma lista: `HOLERITE`, `CNPJ` ou `FREELANCE`;
- A data é exibida como mês por extenso (`janeiro/2025`);
- **TOTAL DE ENTRADAS** calculado com referência estruturada: `=SUM(Entradas[VALOR])`.

### 🔒 TABELAS (oculta)

Guarda as listas que alimentam as listas suspensas, expostas como **nomes definidos**:

| Nome definido | Intervalo | Usado em |
|---|---|---|
| `ListaBancos` | `TABELAS!$A$2:$A$51` | INFORMES → campo BANCO |
| `ListaCategorias` | `TABELAS!$C$2:$C$4` | NOTAS → coluna CATEGORIA |

---

## ✅ Validações automáticas

Todas as células de entrada usam **Validação de Dados** com:

- 💬 **Mensagem de entrada:** aparece ao selecionar a célula e explica o que digitar;
- ⛔ **Alerta de erro:** impede valores fora da regra e explica o motivo.

### Verificação de CPF (dígitos verificadores)

A célula **E7** confere matematicamente se o CPF digitado é válido, usando o algoritmo oficial dos dígitos verificadores (módulo 11):

```excel
=IF(D7="","",
  IF(AND(
      MOD(MOD(SUMPRODUCT(--MID(TEXT(D7,"00000000000"),{1,2,3,4,5,6,7,8,9},1),{10,9,8,7,6,5,4,3,2})*10,11),10)
        = --MID(TEXT(D7,"00000000000"),10,1),
      MOD(MOD(SUMPRODUCT(--MID(TEXT(D7,"00000000000"),{1,2,3,4,5,6,7,8,9,10},1),{11,10,9,8,7,6,5,4,3,2})*10,11),10)
        = --MID(TEXT(D7,"00000000000"),11,1),
      TEXT(D7,"00000000000")<>REPT(LEFT(TEXT(D7,"00000000000"),1),11)
  ),"✔ CPF válido","✖ CPF inválido"))
```

- `TEXT(D7,"00000000000")` garante os 11 dígitos, inclusive zeros à esquerda;
- `SUMPRODUCT` multiplica cada dígito pelo seu peso (10→2 e depois 11→2);
- `MOD(MOD(soma*10,11),10)` calcula o dígito verificador (resto 10 vira 0);
- CPFs com todos os dígitos iguais (ex.: `111.111.111-11`) são recusados.

A **formatação condicional** pinta o resultado de verde (válido) ou vermelho (inválido).

### Outras validações

| Onde | Regra |
|---|---|
| E-mail | `AND(ISNUMBER(FIND("@",D16)), ISNUMBER(FIND(".",D16,FIND("@",D16))), LEN(D16)>=6, ISERROR(FIND(" ",D16)))` |
| Nascimento | Data entre `DATE(1900,1,1)` e `TODAY()` |
| CEP / telefone / celular / título | Número inteiro dentro da faixa de dígitos esperada |
| Rua abreviada | Comprimento do texto ≤ 40 |
| Valores dos bancos | Decimal ≥ 0 |
| Valores das entradas | Decimal > 0 |
| Banco / categoria / SIM-NÃO | Lista suspensa (não aceita valor fora da lista) |
| Banco sem nome | Formatação condicional destaca o campo BANCO em vermelho se houver valor sem banco |

---

## 🧭 Navegação e links rápidos

### Menu lateral

Presente em todas as abas, como num aplicativo:

- **Logo + nome "DECLARE FÁCIL"**;
- Botões **TITULAR**, **INFORMES** e **NOTAS**: a aba atual aparece destacada (●, preenchimento em degradê verde), as demais com contorno;
- Assinatura **SYSTEM BY PEDRO CHRISTO 💚** e ícone do **LinkedIn**, clicável.

Os botões são formas do Excel com **hiperlink para um local no documento** (`#'TITULAR'!A1`, `#'INFORMES'!A1`, `#'NOTAS'!A1`).

### Botões Anterior / Próximo

- **TITULAR:** `PRÓXIMO →`
- **INFORMES:** `← ANTERIOR` e `PRÓXIMO →`
- **NOTAS:** `← ANTERIOR` (no topo, ao lado do total, porque a tabela é longa)

### Links rápidos (dinâmicos)

| Link | Fórmula | O que faz |
|---|---|---|
| ✉ enviar e-mail | `=IF(D16="","",HYPERLINK("mailto:"&D16&"?subject=Declara%C3%A7%C3%A3o%20de%20IR","✉ enviar e-mail"))` | Abre o programa de e-mail com o endereço e o assunto preenchidos. Se o e-mail mudar, o link acompanha |
| 📂 abrir anexo | `=IF(D12="","",HYPERLINK(D12,"📂 abrir anexo"))` | Abre o PDF do informe. Basta salvar o PDF na **mesma pasta** da planilha |
| LinkedIn | Hiperlink no ícone do menu | Abre o perfil do autor |

---

## 🎨 Identidade visual

Paleta **"Aromaterapia e Bem-estar Moderno Verde"**, aplicada nas células, nas formas e também no **tema do arquivo** (cores de destaque), para que a tabela `Entradas` e os hiperlinks sigam o mesmo padrão.

| Cor | Hex | Uso |
|---|---|---|
| 🟩 Verde-petróleo | `#005741` | Fundo do menu lateral, rótulos, cabeçalho da tabela (accent1) |
| 🟩 Verde-escuro | `#025930` | Títulos, textos de entrada, blocos de total |
| 🟩 Verde-musgo | `#31731F` | Degradê dos botões, subtítulos, links |
| ⬜ Verde-menta claro | `#E6F3F2` | Células de preenchimento, faixa de instrução, textos do menu |
| 🟩 Verde-folha | `#478C27` | Botão ativo, faixa "ENTRADAS", bordas dos campos |

- **Tipografia:** Segoe UI (células e botões) e Arial Black (marca);
- **Linhas de grade ocultas** para visual limpo;
- **Cor das guias** de cada aba com tons da paleta.

---

## 🚀 Como usar

1. Baixe o arquivo **`DeclareFacil.xlsx`** e abra no Microsoft Excel (2016, 2019, 2021 ou Microsoft 365);
2. Na aba **TITULAR**, substitua os dados de exemplo pelos seus. Digite só números em CPF, CEP e telefones: a formatação é automática;
3. Confira o status do CPF na coluna E (**✔ CPF válido**);
4. Clique em **PRÓXIMO →** (ou em **INFORMES** no menu) e escolha os bancos na lista, informe os saldos em 31/12 e o nome dos PDFs dos informes;
5. Salve os PDFs na mesma pasta da planilha para que o link **📂 abrir anexo** funcione;
6. Na aba **NOTAS**, registre cada entrada do ano (data, categoria e valor). O total é atualizado sozinho;
7. Com tudo preenchido, abra o programa da Receita Federal e transcreva os dados.

> 💡 **Dica:** para limpar os dados de exemplo, selecione as células de preenchimento (fundo verde-claro) e pressione `Delete`. As fórmulas e validações continuam funcionando.

---

## 🛠️ Como personalizar

| Quero... | Como fazer |
|---|---|
| Adicionar um banco à lista | Reexiba a aba **TABELAS** (botão direito numa guia → *Reexibir*), inclua o banco na coluna A e ajuste o intervalo de `ListaBancos` em *Fórmulas → Gerenciador de Nomes* |
| Criar uma categoria de entrada | Inclua o nome na coluna C da aba **TABELAS** e ajuste `ListaCategorias` |
| Mais linhas de entradas | Digite logo abaixo da última linha da tabela `Entradas`: ela se expande sozinha |
| Um 4º banco | Copie um bloco de banco (4 linhas) da aba INFORMES e inclua a nova célula de valor na fórmula do TOTAL |
| Trocar o link do LinkedIn | Clique com o botão direito no ícone → *Editar Link* |

---

## 🧰 Recursos do Excel utilizados

- Validação de Dados (lista, número inteiro, decimal, data, comprimento de texto e fórmula personalizada)
- Mensagens de entrada e alertas de erro personalizados
- Formatação condicional baseada em fórmula
- Formatos numéricos personalizados (CPF, CEP, telefone, título, moeda e data `[$-416]mmmm/yyyy`)
- Tabela do Excel com referência estruturada (`Entradas[VALOR]`)
- Nomes definidos (`ListaBancos`, `ListaCategorias`)
- Funções: `SUM`, `SUMPRODUCT`, `MOD`, `MID`, `TEXT`, `REPT`, `LEFT`, `FIND`, `ISNUMBER`, `ISERROR`, `LEN`, `AND`, `IF`, `HYPERLINK`, `TODAY`, `DATE`
- Formas com hiperlink interno (menu e botões) e externo (LinkedIn)
- Imagens (logo e ícone)
- Aba oculta para tabelas auxiliares
- Tema de cores personalizado, linhas de grade ocultas e configuração de impressão (paisagem, ajustada à largura)

---

## 📚 Aprendizados

- Pensar a planilha como **produto**, com interface, navegação e experiência de uso, e não só como tabela;
- Usar **validação de dados** para evitar erros na origem, em vez de corrigir depois;
- Transformar uma regra de negócio real (dígito verificador do CPF) em fórmula;
- Usar **IA (Claude)** como copiloto para estruturar fórmulas, revisar validações e documentar o projeto.

---

## 🔭 Próximos passos

- [ ] Aba de **dependentes** (nome, CPF, parentesco, nascimento)
- [ ] Aba de **despesas dedutíveis** (saúde, educação, previdência)
- [ ] Aba de **bens e direitos** (imóveis, veículos, investimentos)
- [ ] Painel com gráfico das entradas por categoria e por mês
- [ ] Checklist de documentos pendentes

---

## 👤 Autor

<table>
  <tr>
    <td>
      <strong>Pedro Christo</strong><br>
      Projeto do bootcamp Santander – Excel com IA (Claude) · DIO<br><br>
      <a href="https://www.linkedin.com/in/pedro-christo-9001b62b4/">
        <img alt="LinkedIn" src="https://img.shields.io/badge/LinkedIn-Pedro%20Christo-005741?style=for-the-badge&logo=linkedin&logoColor=white">
      </a>
    </td>
  </tr>
</table>

<p align="center">Feito com 💚 e Excel</p>
