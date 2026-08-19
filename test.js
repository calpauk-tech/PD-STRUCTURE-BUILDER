import { readFileSync } from 'fs';
// Check if App.tsx uses XLSX
const code = readFileSync('App.tsx', 'utf8');
console.log(code.includes('XLSX.utils') || code.includes('XLSX.write'));
