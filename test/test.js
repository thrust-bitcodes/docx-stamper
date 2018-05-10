let majesty = require('majesty')
var docxStamper = require('../index.js');

var Files = Java.type('java.nio.file.Files');
var Paths = Java.type('java.nio.file.Paths');
var XWPFDocument = Java.type('org.apache.poi.xwpf.usermodel.XWPFDocument');
var ByteArrayInputStream = Java.type('java.io.ByteArrayInputStream');

function exec(describe, it, beforeEach, afterEach, expect, should, assert) {
    describe("Testes de parser de docx", function () {
        it("Deve parsear corretamente o arquivo", function () {
            var fileBytes = getFileBytes('./template.docx');

            var bytes = docxStamper.parseDocument(
                fileBytes,
                {
                    cod_contrato: 1,
                    nome: 'Bruno',
                    idade: '25',
                    nomes: ['Bruno', 'Rodrigo'],
                    responsaveis: [{
                        razao_social: 'João',
                        telefone: '321'
                    }, {
                        razao_social: 'Maria',
                        telefone: '123'
                    }]
                }
            );

            var inputStream = new ByteArrayInputStream(bytes);
            var doc = new XWPFDocument(inputStream);

            var paragraphs = Java.from(doc.getParagraphs()).map(function (paragraph) {
                return paragraph.getParagraphText();
            });

            let expected = ["", "Contrato - 1", "", "", "Primeira linha de texto", "Bruno texto Bruno 25", "Terceira linha de texto", "João", "321", "Maria", "123", "Uma linha de texto", "Outra linha de texto", "Bruno", "Rodrigo", "Mais outra linha de texto", "Ainda outra linha de texto", "Última linha de texto"];
            paragraphs.forEach(function(paragraph, index) {
                expect(paragraph, 'Falha no parágrafo: ' + index).to.equal(expected[index]);
            });
        });
    });
}

function getFileBytes(file) {
    var path = Paths.get(file);
    return Files.readAllBytes(path);
}

function saveFile(file, bytes) {
    var outPath = Paths.get();
    Files.write(outPath, bytes);
}

let res = majesty.run(exec)

print(res.success.length, " scenarios executed with success and")
print(res.failure.length, " scenarios executed with failure.\n")

res.failure.forEach(function (fail) {
    print("[" + fail.scenario + "] =>", fail.execption)
})