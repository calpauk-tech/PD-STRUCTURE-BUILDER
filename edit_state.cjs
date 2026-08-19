const fs = require('fs');
let code = fs.readFileSync('App.tsx', 'utf8');

const target1 = `    const [parsedHeaders, setParsedHeaders] = useState<string[]>([]);`;
const replace1 = `    const [includeWageColumns, setIncludeWageColumns] = useState(false);\n    const [parsedHeaders, setParsedHeaders] = useState<string[]>([]);`;

const target2 = `    const [groupColumn, setGroupColumn] = useState<string>('');`;
const replace2 = `    const [groupColumn, setGroupColumn] = useState<string>('');
    const [deptNumberColumn, setDeptNumberColumn] = useState<string>('');
    const [wageTypeColumn, setWageTypeColumn] = useState<string>('');
    const [rateColumn, setRateColumn] = useState<string>('');
    const [salaryCodeColumn, setSalaryCodeColumn] = useState<string>('');
    const [salaryCode2Column, setSalaryCode2Column] = useState<string>('');`;

if (code.includes(target1) && code.includes(target2)) {
    code = code.replace(target1, replace1).replace(target2, replace2);
    fs.writeFileSync('App.tsx', code);
    console.log("Success");
} else {
    console.log("Failed to find targets");
}
