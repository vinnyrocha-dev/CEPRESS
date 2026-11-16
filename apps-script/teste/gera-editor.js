// Gera o JavaScript que troca o conteúdo de um arquivo no editor do Apps Script (Monaco), conferindo o arquivo certo.
// Uso: node teste/gera-editor.js <arquivo local> <modelo no editor, ex. file_1.js> <saída>
const fs = require('fs');
const [arq, modelo, saida] = process.argv.slice(2);
const t = fs.readFileSync(arq, 'utf8');
const js = '(() => { const ed = monaco.editor.getEditors()[0], m = ed.getModel(); if (!m.uri.path.endsWith(' + JSON.stringify(modelo) +
  ')) return "ARQUIVO ERRADO " + m.uri; const t = ' + JSON.stringify(t) +
  '; ed.executeEdits("claude", [{ range: m.getFullModelRange(), text: t }]); return m.uri.path + " " + m.getValueLength() + " ok=" + (m.getValue() === t); })()';
new Function(js); // confere a sintaxe antes de usar
fs.writeFileSync(saida, js);
