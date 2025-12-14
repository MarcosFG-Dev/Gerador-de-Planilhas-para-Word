# Relatório Pro - Gerador de Orçamentos e Relatórios

Aplicação web simples e eficiente para gerar relatórios e orçamentos com visual profissional (estilo Dashboard/Excel) e exportá-los diretamente para o Microsoft Word (`.doc`).

Totalmente desenvolvida com **HTML, CSS e JavaScript puro**, sem necessidade de instalação, banco de dados ou backend. Funciona diretamente no navegador e pode ser usada offline.

## 📦 Guia de Instalação Completa

Siga este passo a passo para colocar o projeto para rodar localmente de forma segura e previsível, mesmo sem conexão com a internet.

### 1) Requisitos mínimos
- Navegador moderno: **Chrome, Edge, Firefox ou Safari** (versões recentes).
- Espaço em disco: menos de **5 MB** (apenas arquivos estáticos).
- Nenhuma dependência adicional: não precisa de Node, bancos de dados ou servidores.

### 2) Obter os arquivos
Escolha uma das opções abaixo:
- **Download ZIP** (sem Git):
  1. Clique em **Code > Download ZIP** no GitHub e salve o arquivo.
  2. Extraia o conteúdo do `.zip` para uma pasta acessível (ex.: `Documentos/RelatorioPro`).
- **Clonar com Git** (mantém fácil atualização):
  ```bash
  git clone https://github.com/<seu-usuario>/Gerador-de-Planilhas-para-Word.git
  cd Gerador-de-Planilhas-para-Word
  ```

### 3) Abrir o projeto
Há duas maneiras igualmente válidas:
- **Abrir diretamente o arquivo**: clique duas vezes em `index.html` (ou abra no navegador via `Ctrl/Cmd + O`).
- **Usar um servidor estático simples (opcional, recomendado para evitar bloqueios de navegador)**:
  - Python 3 instalado? Execute na pasta do projeto:
    ```bash
    python -m http.server 8000
    ```
    Depois acesse `http://localhost:8000` no navegador.
  - VS Code com Live Server? Abra o projeto e clique em **Go Live** (porta padrão 5500).

### 4) Testar se está tudo ok
- Verifique se a tabela é carregada e se o somatório é calculado ao inserir QNT e VALOR.
- Clique em **Exportar Word** e confirme o download do `.doc`.
- Feche e reabra a página para garantir que o **salvamento automático** recupera os dados digitados.

### 5) Uso offline
Após abrir ao menos uma vez, os arquivos permanecem na máquina. Você pode rodar o `index.html` mesmo sem internet. Se usar um servidor local, não esqueça de iniciá-lo antes de abrir o navegador.

### 6) Atualizar para a versão mais recente
- Via Git: `git pull origin main` (ou o nome do branch que você usa).
- Via ZIP: baixe o novo pacote, extraia em uma pasta vazia e substitua os arquivos antigos.

### 7) Limpeza e desinstalação
Basta excluir a pasta onde o projeto foi salvo. Os dados armazenados no navegador podem ser apagados limpando o **Local Storage** (Configurações do navegador > Privacidade > Limpar dados de site para o arquivo/porta local).

## 🚀 Funcionalidades

- **Edição Estilo Planilha**: Interface ágil para inserção de itens (Tipo, Quantidade, Valor).
- **Cálculos Automáticos**: Multiplicação de Qtd x Valor e somatório total em tempo real.
- **Resumo por Categoria**: Agrupa automaticamente os valores por "Tipo" (Material, Serviço, etc.).
- **Exportação Premium para Word**:
  - Design "Compact Professional": Visual gráfico rico com KPIs e barras de progresso.
  - Otimizado para impressão em **1 única página**.
  - Identidade visual corporativa (Verde/Escuro).
- **Salvamento Automático (Autosave)**: Seus dados são salvos no navegador (Local Storage) para você não perder nada se fechar a aba.
- **Validação de Dados**: Impede a geração de documentos com campos vazios ou inconsistentes.
- **Colagem em Bloco**: Permite copiar dados do Excel e colar diretamente na tabela.

## 🛠️ Como Usar

1. Abra o arquivo `index.html` em qualquer navegador moderno (Chrome, Edge, Firefox).
2. Preencha os **Dados do Documento** (Cliente, Projeto, etc.).
3. Na tabela de itens:
   - Digite o TIPO (ex: Material), QNT e VALOR.
   - O TOTAL é calculado sozinho.
   - Use `Enter` para ir para a próxima linha ou `Tab` para navegar.
4. Clique em **Exportar Word** para baixar o arquivo `.doc` pronto.

## 📂 Estrutura do Projeto

- `index.html`: Estrutura da página.
- `style.css`: Estilização completa (Tema "Corporate Excel").
- `app.js`: Lógica da aplicação (Cálculos, Autosave, Geração do HTML do Word).

## 🎨 Personalização

O design de exportação é definido dentro do arquivo `app.js` (na função `buildWordHTML`). O CSS da página web está em `style.css`.

## 📝 Licença

Livre para uso e modificação.
