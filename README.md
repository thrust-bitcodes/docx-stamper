docx-stamper
===============

docx-stamper é um *bitcode* de renderização de docx baseado em templates para [thrust](https://github.com/thrustjs/thrust).

# Instalação

Posicionado em um app [thrust](https://github.com/thrustjs/thrust), no seu terminal:

```bash
thrust install docx-stamper
```

## Tutorial

Primeiro vamos configurar nosso arquivo de inicialização *startup.js*, nele devemos fazer *require* do *docx-stamper*, e usar o mesmo.

```javascript
//Realizamos o require dos bitcodes
var docxStamper = require('docx-stamper');

var Files = Java.type('java.nio.file.Files');
var Paths = Java.type('java.nio.file.Paths');

var fileBytes = getFileBytes('/template.docx');
var record = getRecord();

var bytes = docxStamper.parseDocument(
  fileBytes,
  record,
  {
    parseToPdf: true //Default: false
  }
);

saveFile('/output.pdf', bytes);

function getRecord() {
  return {
    cod_contrato: 1,
    teste: 'Uma string',
    nomes: ['Bruno', 'Rodrigo'],
    responsaveis: [{
      razao_social: 'João',
      telefone: '321'
    }, {
      razao_social: 'Maria',
      telefone: '123'
    }, {
      razao_social: 'Jorge',
      telefone: '654'
    }]
  };
}

function getFileBytes(file) {
  var path = Paths.get(file);
  return Files.readAllBytes(path);
}

function saveFile(file, bytes) {
  var outPath = Paths.get();
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