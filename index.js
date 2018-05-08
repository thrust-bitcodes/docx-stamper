/*
 * @Author: Bruno Machado
 * @Date: 2018-05-07
 */

var FileInputStream = Java.type('java.io.FileInputStream');
var FileOutputStream = Java.type('java.io.FileOutputStream');
var ByteArrayInputStream = Java.type('java.io.ByteArrayInputStream');
var ByteArrayOutputStream = Java.type('java.io.ByteArrayOutputStream');

var XWPFDocument = Java.type('org.apache.poi.xwpf.usermodel.XWPFDocument');

var ConverterRegistry = Java.type('fr.opensagres.xdocreport.converter.ConverterRegistry');
var ConverterTypeTo = Java.type('fr.opensagres.xdocreport.converter.ConverterTypeTo');
var Options = Java.type('fr.opensagres.xdocreport.converter.Options');
var DocumentKind = Java.type('fr.opensagres.xdocreport.core.document.DocumentKind');

var SIMPLE_PLACEHOLDER_PATTERN = /\{\{(\w+|\.)\}\}/;
var BEGIN_PLACEHOLDER_PATTERN = /\{\{#(\w+)\}\}/;
var END_PLACEHOLDER_PATTERN = /\{\{\/(\w+)\}\}/;

 /**
  * Método utilizado para parsear um template docx, usando um objeto
  * para preenchimento dos placeholders.
  * @param byteArray Array de bytes do arquivo
  * @param record Objeto que será usado como contexto para substituição dos placeholders
  * @param options Objeto de opções para modificar o funcionamento do parser. 
  *   parseToPdf <Boolean> Default: false
  */
function parseDocument(byteArray, record, opt) {
  if (!byteArray) {
    throw new Error('Os bytes do arquivo são obrigatórios.');
  }

  if (!record) {
    throw new Error('Os JSON com os valores a serem renderizados é obrigatório.')
  }

  var options = Object.assign({
    parseToPdf: false
  }, opt);

  var fis = new ByteArrayInputStream(byteArray);
  var doc = new XWPFDocument(fis);

  var blocks = [];
  var currentBlock = null;

  doc.getParagraphs().forEach(function (paragraph) {
    var paragraphText = paragraph.getParagraphText();

    if (currentBlock && END_PLACEHOLDER_PATTERN.exec(paragraphText)) {
      currentBlock.endParagraph = paragraph;
      currentBlock = null;
    } else if (currentBlock != null) {
      currentBlock.paragraphs.unshift(paragraph);
    } else {
      var m = BEGIN_PLACEHOLDER_PATTERN.exec(paragraphText);

      if (m) {
        currentBlock = {
          startParagraph: paragraph,
          name: m[1],
          context: record[m[1]],
          paragraphs: []
        };

        blocks.unshift(currentBlock);
      } else {
        replaceIfItExists(paragraph, record, null);
      }
    }
  });

  blocks.forEach(function (block) {
    if (!block.endParagraph) {
      throw new Error("Bloco " + block.name + " foi iniciado porém não foi finalizado.");
    }

    doc.removeBodyElement(doc.getPosOfParagraph(block.startParagraph));
    doc.removeBodyElement(doc.getPosOfParagraph(block.endParagraph));

    if (block.context) {
      if (block.context.constructor.name != 'Array') {
        block.context = [block.context];
      }

      var listCtx = block.context.slice(0).reverse();

      var cursor;
      var lastParagraph = null;

      listCtx.forEach(function (ctx) {
        block.paragraphs.forEach(function (paragraph) {
          if (lastParagraph == null) {
            cursor = paragraph.getCTP().newCursor();
          } else {
            cursor = lastParagraph.getCTP().newCursor();
          }

          var newParagraph = cloneParagraph(doc, paragraph, cursor);
          replaceIfItExists(newParagraph, ctx, block.name);

          lastParagraph = newParagraph;
        });
      });
    }

    block.paragraphs.forEach(function (paragraph) {
      doc.removeBodyElement(doc.getPosOfParagraph(paragraph));
    });
  });

  var outputStream = new ByteArrayOutputStream();
  doc.write(outputStream);

  if (options.parseToPdf) {
    var inputStream = new ByteArrayInputStream(outputStream.toByteArray());
    outputStream.reset();

    var converterOptions = Options.getFrom(DocumentKind.DOCX).to(ConverterTypeTo.PDF);
		var converter = ConverterRegistry.getRegistry().getConverter(converterOptions);
		converter.convert(inputStream, outputStream, converterOptions);
  }

  return outputStream.toByteArray();
}

function replaceIfItExists(paragraph, context, ctxName) {
  var runs = paragraph.getRuns();

  if (runs) {
    runs.forEach(function (r) {
      var text = r.getText(0);

      if (text && !text.isEmpty()) {
        replaceTextIfItExists(r, text, context, ctxName);
      }
    });
  }
}

function replaceTextIfItExists(run, text, context, ctxName) {
  var m = SIMPLE_PLACEHOLDER_PATTERN.exec(text);

  if (m) {
    var key = m[1];
    var newValue = m[0];

    if ('.'.equals(key) || (ctxName != null && ctxName.equals(key))) {
      newValue = context.toString();
    } else {
      if (context.constructor.name == 'Object') {
        if (context[key]) {
          newValue = context[key].toString();
        }
      }
    }

    var replaceText = text.replace(m[0], newValue);
    run.setText(replaceText, 0);
  }
}

function cloneParagraph(doc, paragraph, cursor) {
  var newParagraph = doc.insertNewParagraph(cursor);

  var pPr = newParagraph.getCTP().isSetPPr() ? newParagraph.getCTP().getPPr()
    : newParagraph.getCTP().addNewPPr();

  pPr.set(paragraph.getCTP().getPPr());

  paragraph.getRuns().forEach(function (r) {
    var nr = newParagraph.createRun();
    cloneRun(nr, r);
  });

  return newParagraph;
}

function cloneRun(clone, source) {
  var rPr = clone.getCTR().isSetRPr() ? clone.getCTR().getRPr() : clone.getCTR().addNewRPr();
  rPr.set(source.getCTR().getRPr());

  clone.setText(source.getText(0));
}

exports = {
  parseDocument: parseDocument
}