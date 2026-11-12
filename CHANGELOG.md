# Changelog

Formato baseado em [Keep a Changelog](https://keepachangelog.com/pt-BR/1.1.0/).
Versões seguem [SemVer](https://semver.org/lang/pt-BR/).

## [Não lançado]

---

## [1.2.1] - 2026-11-12

### Corrigido
- `AprovarPedidosPF.spec`: `python312.dll`, `vcruntime140.dll` excluídas do UPX — compressão corrompía a DLL causando crash "Failed to load Python DLL" ao tentar executar; causava loop de atualização (auto-update baixava novo exe que crashava imediatamente)

## [1.2.0] - 2026-11-10

### Adicionado
- Versão estável em produção
