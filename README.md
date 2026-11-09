# Aprovar Pedidos RPA

Automação para aprovação em lote de pedidos de pagamento de Pessoa Física no sistema **Conveniar** (facto.conveniar.com.br). Lê uma planilha Excel com os pedidos, abre cada um no navegador e executa o fluxo completo de aprovação, registrando o resultado na própria planilha.

## Pré-requisitos

- Python 3.10+
- Google Chrome (portátil em `~/Chrome/`) ou Chrome instalado no sistema
- ChromeDriver compatível com a versão do Chrome (em `~/Chrome/chromedriver.exe` ou no PATH)
- Credenciais de acesso ao Conveniar com permissão de aprovação

## Instalação

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Configuração

As credenciais de acesso estão em `aprovar_pedidos.py` na função `realizar_login()`. Altere o usuário e senha conforme necessário.

```python
# aprovar_pedidos.py
def realizar_login():
    from conveniar import iniciar_sessao
    return iniciar_sessao("seu.usuario", "SuaSenha@")
```

## Uso

```powershell
.\venv\Scripts\Activate.ps1
python aprovar_pedidos.py
```

1. Uma janela abre pedindo para selecionar a planilha `.xlsx`
2. Preencha a **Data de Pagamento** e a **Competência**
3. Clique em **Iniciar** — o robô processa todos os pedidos automaticamente
4. Acompanhe o progresso no painel; ao término, os resultados já estão salvos na planilha

## Entrada e Saída

| Item | Local | Descrição |
|------|-------|-----------|
| Planilha de entrada | Selecionada pelo usuário | Deve ter as colunas obrigatórias (ver `dados/planilha_exemplo.xlsx`) |
| Resultados | Mesma planilha de entrada | Coluna `Resultado Aprovação` atualizada em tempo real |
| Aprovados / Não encontrados | Aba `Processados` | Movidos automaticamente após conclusão |
| Erros | Aba `Erros` | Movidos automaticamente; podem ser retentados |
| Screenshots de debug | `screenshots/` | Geradas apenas se `SALVAR_SCREENSHOTS = True` |

## Planilha de Entrada — Colunas

| Coluna | Obrigatória | Descrição |
|--------|-------------|-----------|
| `Nº Pedido` | ✅ | Número do pedido no formato `12345/2026` |
| `Item PT` | ✅ | Rubrica do plano de trabalho (ex: `Fiscal`, `Limpeza`) |
| `Favorecido` | — | Nome do favorecido (informativo) |
| `Data Nascimento` | — | Data no formato `dd/mm/aaaa`; preenchida automaticamente se o campo estiver vazio no sistema |
| `datanascim` | — | Aceito como variante de `Data Nascimento` |

> Veja o arquivo `dados/planilha_exemplo.xlsx` para um modelo completo.

## Estrutura do Projeto

```
aprovar-pedidos-rpa/
├── aprovar_pedidos.py   # script principal — interface gráfica + fluxo de aprovação
├── conveniar.py         # sessão Selenium e helpers de interação com o Conveniar
├── dados/
│   └── planilha_exemplo.xlsx  # modelo de planilha para o usuário
├── docs/
│   └── uso.md           # manual de uso passo a passo
├── screenshots/         # capturas de tela de debug (ignoradas pelo git)
├── README.md
├── requirements.txt
├── agente.md
└── .gitignore
```

## Configurações Avançadas

Edite as constantes no topo de `aprovar_pedidos.py`:

| Constante | Padrão | Descrição |
|-----------|--------|-----------|
| `NUM_NAVEGADORES` | `5` | Número de sessões paralelas do Chrome |
| `SALVAR_SCREENSHOTS` | `False` | `True` para salvar captura após cada passo (debug) |

## Testes

Execute a suíte automatizada com Python 3.10 ou superior:

```powershell
python -m unittest discover -v
```

A suíte cobre testes unitários de consulta à release, comparação de versões,
detecção do modo executável e instalação; também cobre o fluxo integrado de
consulta, download e preparação da substituição. Chamadas HTTP e a execução do
`.bat` são simuladas, então os testes não acessam o GitHub nem substituem ou
reiniciam o executável real.

## Padrão Corporativo

Este projeto segue a categoria **AUTO** do [Padrão de Engenharia da Empresa](../PADRAO_EMPRESA/specs/00_PADRAO_GERAL.spec.md).
