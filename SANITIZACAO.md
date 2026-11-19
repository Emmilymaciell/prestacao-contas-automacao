# Sanitização desta versão pública

Esta é uma cópia sanitizada de um sistema em uso real num contrato de infraestrutura. Nenhum dado do cliente está aqui; tudo o que aparece como exemplo é inventado.

## O que foi removido

- Arquivos de proveniência/governança interna (`PROVENANCE.md`, `OWNERSHIP.md`). A origem não versionava dado real.

## O que foi trocado

- Rodovias reais → `RD101`, `RD102A`, `RD103`, `RD104`.
- Nome do programa → `Programa`.
- Caminhos locais → variáveis `PRESTACAO_BASE_DIR` (padrão `./examples/engenharia`) e `PRESTACAO_SAIDA_DIR` (padrão `./saida`).
- Mantido: só o mês de início de cada contrato, do qual depende o cálculo de trimestre.

## O que é fictício

- `examples/gerar_exemplo.py` gera PDFs em branco de nomes inventados.

## Como verificar

Sem suíte automatizada. Smoke:

```bat
py -m pip install -r requirements_prestacao_contas.txt
py examples\gerar_exemplo.py
py prestacao_contas_relatorios.py --roads RD101 RD103
```

Varreduras da versão pública (2026-11-19): gitleaks `dir` sem achados; busca por caminho local, e-mail, nomes do cliente e links privados sem ocorrências.
