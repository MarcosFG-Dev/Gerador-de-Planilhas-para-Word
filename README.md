# Relatório Pro - Gerador de Orçamentos e Relatórios

Aplicação web simples e eficiente para gerar relatórios e orçamentos com visual profissional (estilo Dashboard/Excel) e exportá-los diretamente para o Microsoft Word (`.doc`).

Totalmente desenvolvida com **HTML, CSS e JavaScript puro**, sem necessidade de instalação, banco de dados ou backend. Funciona diretamente no navegador e pode ser usada offline.

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
