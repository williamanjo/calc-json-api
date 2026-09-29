# JSON API Importer / Importador de JSON de API

Extensão do LibreOffice Calc que chama uma API JSON (GET ou POST) e transforma cada array da resposta numa tabela nas células. A interface fica em português quando o LibreOffice está em pt e em inglês nos outros idiomas.

Licença MIT, disponível em [LICENSE](LICENSE).

## Gerar e instalar

```
"C:\Program Files\LibreOffice\program\python.exe" build.py
```

O comando gera `dist/calc-json-api.oxt`. Para instalar, dê dois cliques no arquivo ou use **Ferramentas > Gerenciador de extensões > Adicionar**. Depois reinicie o LibreOffice.

Também funciona instalado para todos os usuários da máquina:

```
unopkg add --shared dist/calc-json-api.oxt
```

Para gerar de novo o logo (`assets/logo.png`, usado no site) e o ícone (`images/icon.png`), rode `tools/make_logo.ps1`.

## Uso

O menu **Dados da API** (em inglês, **API Data**) aparece no Calc:

- **Importar da API...**: preencha:
  - **Método**: GET ou POST.
  - **URL**.
  - **Cabeçalhos**: o botão **+ Cabeçalho** adiciona uma linha com Nome e Valor, e o **−** remove a linha (até 15 cabeçalhos). Marque **Sensível** para esconder o valor e não salvá-lo no documento.
  - **Body**: só aparece com o método POST. Se você não informar `Content-Type`, vai `application/json`, e nesse caso o body é validado como JSON antes do envio.

  A barra de status mostra o progresso do download (KB baixados de total; quando a API não informa o tamanho, só os KB baixados) e depois o da gravação nas células ("Importando voos: 4000 de 12000 linhas"). Entre as duas etapas, a extensão lista os arrays encontrados, e você escolhe quais importar e onde colocar:
  - lado a lado, a partir da célula selecionada, com uma coluna vazia entre as tabelas;
  - uma aba para cada array, começando em A1.
- **Atualizar dados da API**: busca de novo e reescreve cada tabela na mesma posição. O conteúdo antigo é limpo, mas a formatação que você aplicou é mantida.
- **Remover vínculos com a API**: faz o documento esquecer as posições salvas. Os dados continuam nas células.

## Como o JSON vira tabela

- Cada array, em qualquer nível, vira uma tabela (`voos`, `dados.passageiros`...). A primeira linha traz o caminho do array e a segunda, o cabeçalho.
- Arrays de objetos: uma coluna por chave. Objetos aninhados são achatados (`origem.iata`), e arrays aninhados ficam na célula como texto JSON.
- Arrays de valores simples viram uma única coluna, `valor`.
- `true`/`false` viram texto e `null` vira célula vazia. Inteiros acima de 2^53 são gravados como texto para não perderem dígitos.

## Segurança

O método, a URL, os cabeçalhos e o body são gravados nas propriedades do documento (`JsonApiImportConfig`). O valor de um cabeçalho marcado como **Sensível** não é gravado: fica só na memória enquanto o LibreOffice está aberto, e o "Atualizar" pede de novo depois que o programa é reiniciado. Os cabeçalhos **não marcados são gravados com o valor**, mesmo que sejam um token.

O **body é gravado inteiro**. Não coloque senhas no body de um documento que você vai compartilhar.

Documentos criados nas versões 1.0 a 1.2 continuam funcionando. Os cabeçalhos que eles não gravavam passam a aparecer marcados como Sensível.
