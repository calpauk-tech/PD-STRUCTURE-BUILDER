const fs = require('fs');
let code = fs.readFileSync('App.tsx', 'utf8');

const targetStr = `        setReviewItems(items);
        setStep(appMode === 'update' ? 2 : 3); // Config or Review
    };

    // Add custom item in Review`;

const replaceStr = `        setReviewItems(items);
        setStep(appMode === 'update' && deleteInputMethod !== 'file' ? 2 : 3); // Config or Review
    };

    // Add custom item in Review`;

function removeWhitespace(str) {
    return str.replace(/\s+/g, '');
}

if (removeWhitespace(code).includes(removeWhitespace(targetStr))) {
    const regexStr = targetStr.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
    code = code.replace(new RegExp(regexStr), replaceStr);
    fs.writeFileSync('App.tsx', code);
    console.log("Success replacing setStep in generateActionReviewList");
} else {
    console.log("Failed replacing setStep in generateActionReviewList");
}
