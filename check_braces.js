const fs = require('fs');
const code = fs.readFileSync('script.js','utf8');
const stack = [];
let mode = null;
for (let i = 0; i < code.length; i += 1) {
  const ch = code[i];
  const prev = code[i - 1];
  if (mode) {
    if (mode === 'single' && ch === "'" && prev !== '\\') { mode = null; }
    else if (mode === 'double' && ch === '"' && prev !== '\\') { mode = null; }
    else if (mode === 'back' && ch === '' && prev !== '\\') { mode = null; }
    continue;
  }
  if (ch === "'") { mode = 'single'; continue; }
  if (ch === '"') { mode = 'double'; continue; }
  if (ch === '') { mode = 'back'; continue; }
  if (ch === '{') {
    stack.push(i);
  } else if (ch === '}') {
    if (!stack.length) {
      console.log('Unmatched } at', i);
      process.exit(0);
    }
    stack.pop();
  }
}
if (stack.length) {
  console.log('Unmatched { at positions', stack);
} else {
  console.log('Braces balanced');
}
