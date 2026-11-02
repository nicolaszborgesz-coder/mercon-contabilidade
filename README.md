# Mercon Contabilidade - Automação Fiscal V5.1

Sistema em Python para escritórios contábeis que automatiza conciliação de NF-e.

### O que o sistema faz
- Lê 100+ XMLs de NF-e (ElementTree com namespace Sefaz)
- Lê extratos bancários CSV com detecção automática de separador `; , \t` e encoding `utf-8 / latin-1`
- Identifica clientes por CNPJ (16 clientes simultâneos)
- Classifica: `OK-CONCILIADO`, `PENDENTE PARCIAL/TOTAL`, `DIVERGENCIA`, `SOBRA BANCARIA`
- Gera Excel profissional com 4 abas: CONCILIACAO, RESUMO, ERROS_XML, PASTAS_NAO_ENCONTRADAS
- Estilo visual Mercon: Azul #17365D, Verde #C6EFCE, Amarelo #FFEB9C, Vermelho #FFC7CE

### Stack
Python • Decimal (precisão fiscal) • OpenPyXL • XML • CSV • Pathlib

### Como rodar
```bash
pip install openpyxl
python src/conciliacao.py
