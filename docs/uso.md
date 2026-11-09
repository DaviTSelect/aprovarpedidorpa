# Manual de Uso — Aprovar Pedidos RPA

Este guia explica como usar o robô de aprovação de pedidos passo a passo, sem precisar saber programação.

---

## O que o robô faz

Ele abre o sistema Conveniar automaticamente, busca cada pedido da sua planilha e executa todo o processo de aprovação — conta bancária, plano de trabalho, INSS Patronal e aprovação final — salvando o resultado na planilha em tempo real.

---

## Antes de começar

- Feche a planilha no Excel antes de rodar o robô (ele precisa salvar nela)
- Certifique-se de que o Chrome portátil está em `C:\Users\SeuUsuário\Chrome\`
- Sua planilha deve estar no formato `.xlsx`

---

## Preparando a planilha

1. Abra o arquivo `dados\planilha_exemplo.xlsx` como modelo
2. Preencha uma linha por pedido com as seguintes colunas:

| Coluna | O que colocar | Exemplo |
|--------|---------------|---------|
| **Nº Pedido** | Número do pedido exatamente como aparece no Conveniar | `54320/2026` |
| **Item PT** | Rubrica do plano de trabalho | `Fiscal` ou `Limpeza` |
| **Favorecido** | Nome da pessoa (só para controle) | `João da Silva` |
| **Data Nascimento** | Data de nascimento no formato dd/mm/aaaa (opcional) | `15/03/1990` |

> ⚠️ **Importante:** Não altere o nome das colunas. O robô lê pelo nome exato.

---

## Executando o robô

1. Abra o terminal (PowerShell) na pasta do projeto
2. Execute:
   ```powershell
   .\venv\Scripts\Activate.ps1
   python aprovar_pedidos.py
   ```
3. Uma janela abrirá. Clique em **Selecionar Planilha** e escolha seu arquivo `.xlsx`
4. Preencha:
   - **Data de Pagamento** — data em que os pagamentos serão realizados
   - **Competência** — mês/ano de referência (ex: `11/2026`)
5. Clique em **Iniciar**

---

## Acompanhando o progresso

O painel mostra cada pedido sendo processado em tempo real. Você verá mensagens como:

- `✅ Pedido 54320/2026 aprovado com sucesso!` — aprovado normalmente
- `❌ Erro ao processar 54321/2026: Conta principal não encontrada` — movido para aba Erros

---

## Resultado na planilha

Ao final, sua planilha terá três abas:

| Aba | O que contém |
|-----|--------------|
| **Aba 1 (original)** | Pedidos com falha que podem ser retentados |
| **Processados** | Pedidos aprovados ou não encontrados no sistema |
| **Erros** | Pedidos com erro (conta bancária não encontrada, etc.) |

A coluna **Resultado Aprovação** é preenchida automaticamente para cada pedido.

---

## Problemas comuns

| Problema | Solução |
|----------|---------|
| "Selecione uma planilha .xlsx" | Converta o arquivo de `.xls` para `.xlsx` antes de usar |
| "A planilha precisa ter a coluna Nº Pedido" | Verifique se o cabeçalho está escrito exatamente assim |
| "file is not a zip" | Feche a planilha no Excel e tente novamente |
| Pedido vai para aba Erros | Verifique se a conta bancária principal está cadastrada no Conveniar |
| Nenhum pedido processado | Verifique se a coluna `Status` tem o valor `ok` para as linhas desejadas |

---

## Dicas

- **Dados de nascimento:** se já estiver preenchido no Conveniar, o robô não altera. Só preenche se estiver vazio.
- **Paralelo:** o robô abre até 5 abas do Chrome ao mesmo tempo. Não feche as janelas enquanto estiver rodando.
- **Screenshots de debug:** se precisar investigar um erro, edite `aprovar_pedidos.py` e mude `SALVAR_SCREENSHOTS = True`. As imagens ficam na pasta `screenshots/`.

---

## Suporte

Em caso de dúvidas ou erros não cobertos por este manual, contate a equipe de automações.
