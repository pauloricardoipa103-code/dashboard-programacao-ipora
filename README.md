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

1. Abra `admin.html`.
2. Selecione a planilha atualizada enviada pela Equatorial.
3. Confira os totais exibidos na tela.
4. Informe um token GitHub com permissao de escrita no repositorio.
5. Clique em `Publicar atualizacao`.

A base publicada fica em `dados/anomalias.json`. Todos que acessarem o link do painel passam a carregar essa mesma base central.

## Como exportar CSV

Use os filtros desejados no painel e clique em `Exportar tabela CSV`. O arquivo exportado contem somente os registros filtrados no momento.
