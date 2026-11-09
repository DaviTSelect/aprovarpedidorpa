# Contexto do Projeto para Agente IA

## O que este projeto faz

Automação (RPA) para aprovação em lote de pedidos de pagamento de Pessoa Física no sistema Conveniar (facto.conveniar.com.br). Lê pedidos de uma planilha Excel, abre cada um no Chrome via Selenium e executa o fluxo completo de aprovação: dados cadastrais → conta bancária → datas → valores → plano de trabalho → INSS Patronal → aprovação final. Salva o resultado na planilha original em tempo real.

## Tecnologias

- Python 3.12
- Selenium 4.46 (Chrome, ChromeDriver portátil em `~/Chrome/`)
- openpyxl 3.1.5 (leitura/escrita de planilhas)
- pandas 3.0.5 (leitura inicial da planilha)
- tkinter (interface gráfica mínima para seleção de arquivo e painel de progresso)
- threading (5 sessões paralelas do Chrome)

## Arquitetura

```
aprovar-pedidos-rpa/
├── aprovar_pedidos.py   # ponto de entrada — GUI tkinter + funcao_aprovar + executar_parte
├── conveniar.py         # sessão Selenium independente; login via iniciar_sessao()
├── dados/
│   └── planilha_exemplo.xlsx
├── docs/
│   └── uso.md
├── screenshots/         # debug; controlado por SALVAR_SCREENSHOTS
├── README.md
├── requirements.txt
└── .gitignore
```

**Módulos principais:**
- `aprovar_pedidos.py` — toda a lógica de automação e GUI. Função central: `funcao_aprovar(navegador, pedido, numero, rubrica, data_pagamento, competencia, data_nascimento)`. Paralelismo via `executar_parte` + `ThreadPoolExecutor`.
- `conveniar.py` — cria sessão Chrome isolada com perfil temporário. `iniciar_sessao(usuario, senha)` retorna o WebDriver já logado.

## Fluxo principal de aprovação (30 passos)

1. Busca o pedido pelo número na tela de consulta
2. Abre a aba de edição do pedido
3. Clica em "Atualizar Dados Cadastrais" → preenche data de nascimento (só se vazio no site e preenchido na planilha) → OK
4. Procura conta bancária → seleciona conta principal (coluna "Sim") → confirma checkbox
5. Preenche data de pagamento e competência
6. Salva o pedido
7. Abre o passo de aprovação, preenche data/competência novamente
8. Marca rubrica (Item PT) no plano de trabalho → OK (trata diálogo sim/não se aparecer)
9. Marca INSS Patronal no grid e no plano → OK (mesmo tratamento)
10. Clica em "Aprovar" → trata diálogo de condições (sim/não) → aguarda mensagem "Pedido Nº XXXXX aprovado." → clica OK
11. Registra resultado na planilha: Aprovado → aba "Processados"; Erro → aba "Erros"; Falha temporária → permanece na aba 1

## Como executar

```powershell
.\venv\Scripts\Activate.ps1
python aprovar_pedidos.py
```

## Padrões do projeto

- Código em português (variáveis, funções, prints)
- `scroll_js_clicar(elem)` antes de TODA interação com elemento (scrollIntoView + click via JS + aguarda loader)
- `aguardar_loader_sumir()` com `poll_frequency=0.1` após cada ação
- `eid(id)` para localizar elementos por ID com WebDriverWait
- Escrita na planilha protegida por `threading.Lock()` — busca linha pelo número do pedido (não por índice) para sobreviver a deleções concorrentes
- `SALVAR_SCREENSHOTS = False` em produção; `True` apenas para debug

## O que NÃO fazer

- Não commitar planilhas com dados reais de CPF, dados bancários ou datas de nascimento
- Não commitar a pasta `venv/` ou `screenshots/`
- Não alterar `conveniar.py` sem testar o login — ele é compartilhado com outros scripts
- Não aumentar `NUM_NAVEGADORES` além de 5 sem verificar se o servidor do Conveniar aguenta
- Não remover o `threading.Lock()` da função de escrita na planilha

## Dados sensíveis

- Credenciais do Conveniar hardcoded em `realizar_login()` — não commitar sem ofuscação
- Planilhas de entrada contêm CPF, dados bancários e datas de nascimento de pessoas físicas — nunca versionar
- A pasta `dados/` deve conter apenas a planilha exemplo com dados fictícios

## Referência de padrão corporativo

[PADRAO_EMPRESA/specs/00_PADRAO_GERAL.spec.md](../PADRAO_EMPRESA/specs/00_PADRAO_GERAL.spec.md)
