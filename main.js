const puppeteer = require('puppeteer');
const fs = require('fs');
var userData;

function loadUserData() {
    let rawData = fs.readFileSync("config.json", "utf-8");
    let data = JSON.parse(rawData);
    return data;
}

async function main() {
    userData = loadUserData();
    await loginAndGetAssignments(userData);
}

async function loginAndGetAssignments(userdata) {
    const browser = await puppeteer.launch({ headless: false });
    const page = await browser.newPage();
    //START OF LOGIN
    await page.goto('https://office.com', { waitUntil: 'networkidle0' });
    await page.click("#hero-banner-sign-in-to-office-365-link");
    console.log("alma");
    await page.waitForNavigation({ waitUntil: 'networkidle2' });
    await page.type('#i0116', userdata.teams.username);
    await page.click("#idSIButton9");
    await page.waitForNavigation({ waitUntil: 'networkidle0' });
    await page.type('#i0118', userdata.teams.password);
    await page.click("#idSIButton9");
    await page.waitForNavigation({ waitUntil: 'networkidle0' });
    await page.click("#idSIButton9");
    await page.waitForNavigation({ waitUntil: 'networkidle2' });
    await page.goto('https://teams.microsoft.com/', { waitUntil: 'networkidle0' });
    await page.click(".use-app-lnk");
    //END OF LOGIN
    console.log("Ads");
    await browser.close();
}

main();