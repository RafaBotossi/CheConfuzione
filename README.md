# Che Confuzione

Roteiro de viagem responsivo, criado para consultar a programação no computador e no celular.

## Como abrir

O jeito mais simples é iniciar um servidor local na pasta:

```powershell
python -m http.server 8000
```

Depois, abra `http://localhost:8000`.

## Recursos

- Carrega automaticamente `Roteiro Principal.xlsx`.
- Importa novas planilhas `.xlsx`, `.xls` ou `.csv` no navegador.
- Detecta a aba e as colunas do roteiro automaticamente.
- Une data e horário e ignora a coluna de valor e as totalizações.
- Filtra por dia e pesquisa em todos os campos.
- Abre endereços no Google Maps.
- Permite preencher ou editar endereços, salvos apenas no aparelho.
- Layout de tabela no computador e cartões no celular.
- Pode ser instalado como app e mantém os arquivos principais disponíveis offline.

O arquivo importado e os endereços editados ficam no `localStorage` do navegador e não são enviados para servidores.
