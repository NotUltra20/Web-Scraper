# Web Scraper — ONGs da Paraíba (ONGs Brasil)

Extrai todas as ONGs cadastradas na Paraíba no site [ONGs Brasil](https://www.ongsbrasil.com.br/default.asp?Pag=1&Destino=Instituicoes&Estado=PB) e gera uma planilha Excel.

## Fluxo de coleta

1. **Estado (PB)** — lista de cidades
2. **Cidade** — lista de ONGs (com paginação quando houver)
3. **Detalhe** — página com endereço, telefone, e-mail, CNPJ, etc.
4. **Excel** — arquivo `ongs_paraiba.xlsx`

## Requisitos

- Python 3.10+ (no Windows use `py`)

## Instalação

```powershell
cd "c:\Users\Adm\Desktop\Web-Scraper de ONGs"
py -m pip install -r requirements.txt
```

## Uso

### Extrair todas as ONGs da Paraíba

```powershell
py scraper.py
```

Isso pode levar **várias horas** (centenas de ONGs × ~1 s por requisição). O progresso é salvo em `checkpoint.json`; se interromper, execute de novo para continuar.

### Opções úteis

```powershell
# Uma cidade só (teste rápido)
py scraper.py --cidade "Alagoa Nova" -o teste.xlsx

# Intervalo maior entre requisições (mais educado com o servidor)
py scraper.py --delay 2

# Reiniciar do zero
py scraper.py --no-resume

# Log detalhado
py scraper.py -v
```

## Saída

| Coluna | Descrição |
|--------|-----------|
| Código | ID no site |
| Nome | Razão social / nome principal |
| Cidade (listagem) | Cidade pela qual a ONG foi encontrada |
| Bairro (listagem) | Bairro na listagem da cidade |
| Endereço, Bairro, CEP, Cidade, Estado, País | Dados da página de detalhe |
| Telefone, Nome Fantasia, E-mail, Site, CNPJ | Contato e cadastro |
| URL | Link da página de detalhe |

## Arquivos gerados

- `ongs_paraiba.xlsx` — planilha final
- `checkpoint.json` — progresso (pode apagar após concluir)

## Observações

- Alguns campos podem vir vazios se a ONG não os preencheu no cadastro.
- A codificação do site é Latin-1; o scraper trata isso automaticamente.
