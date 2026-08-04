# Automação — Ata do Comitê de Cadastros

Automação em Python que transforma as planilhas de itens enviadas pelos compradores em uma **ata de reunião em Word (.docx)** completa, formatada e pronta para envio — eliminando o trabalho manual de montar a ata item por item toda semana.

Desenvolvido para uso interno no processo do Comitê de Cadastros da rede de supermercados onde trabalho, como Assistente de Cadastro.

## Contexto

Toda semana, cada comprador envia uma planilha com os produtos que vai apresentar no comitê. Antes desta automação, a ata era montada manualmente: copiar item por item, formatar, organizar por comprador. Esse processo tinha dois problemas — tomava tempo e era sujeito a erro humano (item esquecido, formatação inconsistente).

A automação resolve isso em duas etapas.

## Como funciona

O processo é composto por dois scripts que rodam em sequência:

1. **`Gerar_Planilha_Unificada.py`**
   Lê todas as planilhas `.xlsx` recebidas dos compradores (pasta de entrada), junta os dados em uma única planilha organizada por comprador, e salva o resultado com o nome da data do comitê.

2. **`Gerar_Ata.py`**
   Lê a planilha unificada gerada no passo anterior e monta automaticamente um documento Word (`.docx`) formatado como ata oficial — com título, participantes, itens agrupados por comprador (em seções recolhíveis), checkboxes clicáveis para aprovação por tipo de loja, e seções de encaminhamentos e deliberações.

O `.docx` é gerado manipulando diretamente o XML do formato OOXML do Word (sem depender de bibliotecas como `python-docx`), o que permite recursos avançados como seções recolhíveis nativas do Word e checkboxes interativos.

## Uso

```bash
# 1. Colocar as planilhas dos compradores na pasta de entrada
# 2. Rodar o unificador
python "Gerar_Planilha_Unificada.py"

# 3. Rodar o gerador de ata
python "Gerar_Ata.py"
```

Cada script pergunta a data da reunião no terminal (ou usa a data de hoje, se você apertar Enter direto).

### Dependências

```bash
pip install openpyxl
```

## Sobre este projeto

Este script nasceu de um problema real do meu dia a dia como Assistente de Cadastro, e foi meu ponto de entrada para estudar Python e TI de forma mais estruturada. Ele foi desenvolvido com apoio de IA (Claude, da Anthropic) — usada tanto para escrever partes do código quanto para explicar conceitos ao longo do processo. Optei por essa transparência porque parte do meu aprendizado tem sido justamente revisar o código gerado, entender cada trecho e me tornar capaz de manter e evoluir a automação sozinho.

As planilhas de entrada e as atas geradas contêm dados internos da empresa e não fazem parte deste repositório (ver `.gitignore`).
