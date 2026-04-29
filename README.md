# Dashboard de Anomalias REMO

Dashboard HTML para acompanhamento das anomalias pendentes e executadas encaminhadas pela Equatorial para a REMO.

## Arquivos

- `index.html`: dashboard pronto para abrir no navegador ou publicar pelo GitHub Pages.
- `dashboard_anomalias_remo.html`: copia local do dashboard gerado.
- `generate_dashboard.py`: gerador usado para recriar o HTML a partir da planilha `Pasta1.xlsx` e da logo REMO.

## Como visualizar

Abra `index.html` no navegador.

O mapa interativo usa base cartografica online, entao a maquina precisa estar conectada a internet para carregar o fundo geografico.

## Como atualizar a base

1. Abra o dashboard.
2. Clique em `Importar Excel`.
3. Selecione a planilha atualizada enviada pela Equatorial.
4. O dashboard atualiza os KPIs, graficos, mapa e tabela automaticamente.

A nova base fica salva no proprio navegador utilizado. Nao e necessario banco de dados.

## Como exportar CSV

Use os filtros desejados no painel e clique em `Exportar tabela CSV`. O arquivo exportado contem somente os registros filtrados no momento.
