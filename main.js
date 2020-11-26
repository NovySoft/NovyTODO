const fs = require('fs');
var userData;

function loadUserData() {
    let rawData = fs.readFileSync("config.json", "utf-8");
    let data = JSON.parse(rawData);
    return data;
}

function main() {
    userData = loadUserData();
}

main();