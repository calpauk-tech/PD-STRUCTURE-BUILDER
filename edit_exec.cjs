const fs = require('fs');
let code = fs.readFileSync('App.tsx', 'utf8');

const targetExec = `                        const created = await createEmployeeGroup(item.name);
                        item.status = 'success';
                        item.createdId = created.id;
                        item.resultMessage = \`Created (ID: \${created.id})\`;
                        
                        // Apply default rates if enabled
                        if (applyGroupDefaults) {
                            try {
                                await updateEmployeeGroupDefaultRate(created.id, {
                                    defaultWageType: globalDefaultWageType,
                                    hourlyRate: globalHourlyRate !== '' ? Number(globalHourlyRate) : undefined,
                                    shiftRate: globalShiftRate !== '' ? Number(globalShiftRate) : undefined,
                                    salaryCode: globalSalaryCode,
                                    salaryCode2: globalSalaryCode2
                                });
                                item.resultMessage += ' + Rates applied';
                            } catch (rateErr: any) {
                                item.resultMessage += \` (Warning: Rates failed - \${rateErr.message})\`;
                            }
                        }`;

const replaceExec = `                        const created = await createEmployeeGroup(item.name);
                        item.status = 'success';
                        item.createdId = created.id;
                        item.resultMessage = \`Created (ID: \${created.id})\`;
                        
                        // Apply rates from file upload or global defaults
                        let payloadToApply = item.updatePayload;
                        
                        // If no file payload but global is enabled, use global
                        if (!payloadToApply && applyGroupDefaults) {
                            payloadToApply = {
                                defaultWageType: globalDefaultWageType,
                                hourlyRate: globalHourlyRate !== '' ? Number(globalHourlyRate) : undefined,
                                shiftRate: globalShiftRate !== '' ? Number(globalShiftRate) : undefined,
                                salaryCode: globalSalaryCode,
                                salaryCode2: globalSalaryCode2
                            };
                        }

                        if (payloadToApply && Object.keys(payloadToApply).length > 0) {
                            try {
                                await updateEmployeeGroupDefaultRate(created.id, payloadToApply);
                                item.resultMessage += ' + Rates applied';
                            } catch (rateErr: any) {
                                item.resultMessage += \` (Warning: Rates failed - \${rateErr.message})\`;
                            }
                        }`;

function removeWhitespace(str) {
    return str.replace(/\s+/g, '');
}

const codeNoSpace = removeWhitespace(code);
const targetNoSpace = removeWhitespace(targetExec);

if (codeNoSpace.includes(targetNoSpace)) {
    const regexStr = targetExec.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
    const regex = new RegExp(regexStr);
    code = code.replace(regex, replaceExec);
    fs.writeFileSync('App.tsx', code);
    console.log("Success replacing execution logic");
} else {
    console.log("Failed replacing execution logic");
}
