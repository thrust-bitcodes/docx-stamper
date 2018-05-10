docx-stamper [![Build Status](https://travis-ci.org/thrust-bitcodes/docx-stamper.svg?branch=master)](https://travis-ci.org/thrust-bitcodes/docx-stamper)
===============

docx-stamper é um *bitcode* de renderização de docx baseado em templates para [thrust](https://github.com/thrustjs/thrust).

# Instalação

Posicionado em um app [thrust](https://github.com/thrustjs/thrust), no seu terminal:

```bash
thrust install docx-stamper
```

## Tutorial

Primeiro vamos configurar um template em docx, com o seguinte texto:

```
{{nome}}

{{#responsaveis}}
  - {{razao_social}}
    - {{telefone}}
{{/responsaveis}}

{{#contato}}
  {{contato}}
{{/contato}}
```
A sintaxe é semelhante a da espeficicação do mustache, onde `{{nome}}` será substituido pelo valor do objeto.

`{{#responsaveis}}` e `{{/responsaveis}}` sinalizam o inicio e o fim de um bloco, caso `responsaveis` exista no objeto, a expressão contida será printada, se não, será deletada.
Caso o valor da propriedade em questão seja um array, este bloco será repetido para cada item do array.

Inicios e fim de blocos devem estar sempre sozinhos em uma linha, conforme mostrado no exemplo acima, não se preocupe, essas linhas não permanecerão no documento.

Em seguida vamos criar o arquivo de inicialização *startup.js*, nele devemos fazer *require* do *docx-stamper*, e usar o mesmo.

```javascript
//Realizamos o require dos bitcodes
var docxStamper = require('docx-stamper');

var Files = Java.type('java.nio.file.Files');
var Paths = Java.type('java.nio.file.Paths');

var fileBytes = getFileBytes('/template.docx');
var record = {
  nome: 'Bruno',
  responsaveis: [{
    razao_social: 'João',
    telefone: '321'
  }, {
    razao_social: 'Maria',
    telefone: '123'
  }]
};

var bytes = docxStamper.parseDocument(
  fileBytes,
  record,
  {
    parseToPdf: true //Default: false
  }
);

saveFile('/output.pdf', bytes);

function getFileBytes(file) {
  var path = Paths.get(file);
  return Files.readAllBytes(path);
}

function saveFile(file, bytes) {
  var outPath = Paths.get(file);
  Files.write(outPath, bytes);
}


```

## API

```javascript
 /**
  * Método utilizado para parsear um template docx, usando um objeto
  * para preenchimento dos placeholders.
  * @param byteArray Array de bytes do arquivo
  * @param record Objeto que será usado como contexto para substituição dos placeholders
  * @param options Objeto de opções para modificar o funcionamento do parser. 
  *   parseToPdf <Boolean> Default: false
  */
parseDocument(byteArray, record, opt)
```