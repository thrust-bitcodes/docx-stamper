let majesty = require('majesty')

function exec(describe, it, beforeEach, afterEach, expect, should, assert) {
    describe("Testes de parser de docx", function () {
        it(" 1 deve ser igual a 1", function () {
            expect(1).to.equal(1);
        });
    });
}

let res = majesty.run(exec)

print(res.success.length, " scenarios executed with success and")
print(res.failure.length, " scenarios executed with failure.\n")

res.failure.forEach(function (fail) {
    print("[" + fail.scenario + "] =>", fail.execption)
})