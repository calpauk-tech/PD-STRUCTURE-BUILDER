const fs = require('fs');
let content = fs.readFileSync('services/plandayService.ts', 'utf8');

content = content.replace(
    'const last = parts[parts.length - 1];',
    'const last = parts[parts.length - 1] || parts[parts.length - 2];'
);

fs.writeFileSync('services/plandayService.ts', content);
