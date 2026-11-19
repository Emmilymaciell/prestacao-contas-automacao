# Automação de prestação de contas - relatórios RM e RO

## O que este script faz

1. Varre as pastas de cada rodovia para localizar arquivos de relatório com nomenclatura `RM` e `RO`.
2. Gera um **Word por mês** em:
   `saida/Word` (ou `PRESTACAO_SAIDA_DIR`)
3. Gera um **PDF mensal por rodovia** com a junção dos PDFs RM e RO do mês.
4. Gera um **PDF trimestral por rodovia** com a junção de todos os PDFs dos 3 meses do trimestre.
5. Organiza tudo dentro da pasta da rodovia, criando pasta por trimestre e por mês.
6. Tenta liberar o espaço local dos arquivos do OneDrive após ler cada PDF, sem apagar nada da nuvem.

## Bibliotecas necessárias

```bash
pip install -r requirements_prestacao_contas.txt
```

## Como rodar

```bash
python prestacao_contas_relatorios.py
```

Para processar só algumas rodovias:

```bash
python prestacao_contas_relatorios.py --roads RD101 RD103
```

Para mais detalhes no log:

```bash
python prestacao_contas_relatorios.py --log-level DEBUG
```

## Observações importantes

- O script preserva a estrutura original das páginas dos PDFs porque apenas concatena as páginas existentes.
- A organização trimestral é feita em blocos de 3 meses a partir do início de cada contrato:
  - RD101, RD102A e RD103: início em setembro/2025
  - RD104: início em outubro/2025
- Se algum arquivo não tiver mês/ano claro no nome, o script usa a pasta, o sequencial do RM ou a data de modificação como fallback.
- Os logs ficam em:
  `saida/_logs`

## Ponto de atenção

Se houver muitos arquivos com nomenclaturas muito fora do padrão, vale ajustar as expressões regulares do script para refletir exatamente como esses nomes aparecem nas suas pastas.
