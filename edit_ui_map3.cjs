const fs = require('fs');
let code = fs.readFileSync('App.tsx', 'utf8');

const targetUIMap = `                {/* --- STEP 3: FIELD MAPPING --- */}
                {step === 2 && (appMode === 'import' || (appMode === 'update' && deleteInputMethod === 'file')) && (
                    <div className="bg-white p-8 rounded-2xl shadow-sm border border-gray-200 space-y-8">`;

const replaceUIMap = `                {/* --- STEP 3: FIELD MAPPING --- */}
                {step === 2 && (appMode === 'import' || deleteInputMethod === 'file') && (
                    <div className="bg-white p-8 rounded-2xl shadow-sm border border-gray-200 space-y-8">`;

function removeWhitespace(str) {
    return str.replace(/\s+/g, '');
}

if (removeWhitespace(code).includes(removeWhitespace(targetUIMap))) {
    const regexStr = targetUIMap.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
    code = code.replace(new RegExp(regexStr), replaceUIMap);
    fs.writeFileSync('App.tsx', code);
    console.log("Success replacing Field Mapping condition");
} else {
    console.log("Failed replacing Field Mapping condition");
}

const targetButton = `                            <button 
                                onClick={appMode !== 'import' ? generateActionReviewList : generateReviewList}`;

const replaceButton = `                            <button 
                                onClick={appMode === 'import' || (appMode === 'update' && deleteInputMethod === 'file') ? generateReviewList : generateActionReviewList}`;

if (removeWhitespace(code).includes(removeWhitespace(targetButton))) {
    const regexStr = targetButton.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
    code = code.replace(new RegExp(regexStr), replaceButton);
    fs.writeFileSync('App.tsx', code);
    console.log("Success replacing mapping button");
} else {
    console.log("Failed replacing mapping button");
}

const targetUpdateConfig = `                {/* --- STEP 3 (UPDATE MODE): CONFIGURE DEFAULTS --- */}
                {step === 2 && appMode === 'update' && (
                    <div className="bg-white p-8 rounded-2xl shadow-sm border border-gray-200 space-y-8">`;

const replaceUpdateConfig = `                {/* --- STEP 3 (UPDATE MODE): CONFIGURE DEFAULTS --- */}
                {step === 2 && appMode === 'update' && deleteInputMethod !== 'file' && (
                    <div className="bg-white p-8 rounded-2xl shadow-sm border border-gray-200 space-y-8">`;

if (removeWhitespace(code).includes(removeWhitespace(targetUpdateConfig))) {
    const regexStr = targetUpdateConfig.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
    code = code.replace(new RegExp(regexStr), replaceUpdateConfig);
    fs.writeFileSync('App.tsx', code);
    console.log("Success replacing Update Config condition");
} else {
    console.log("Failed replacing Update Config condition");
}
