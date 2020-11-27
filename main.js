const puppeteer = require('puppeteer');
const fs = require('fs');
var html2json = require('html2json').html2json;
var userData;
var homeworkObjectList = [];

function loadUserData() {
    console.log("Loading UserData");
    let rawData = fs.readFileSync("config.json", "utf-8");
    let data = JSON.parse(rawData);
    return data;
}

async function main() {
    userData = loadUserData();
    homeworkObjectList = [];
    let temp = await loginAndGetTeamsAssignments(userData);
    homeworkObjectList.push.apply(homeworkObjectList, temp);
}

//This function returns a list with the objects of the teams assignments
async function loginAndGetTeamsAssignments(userdata) {
    const browser = await puppeteer.launch({
        headless: false,
        args: ['--no-sandbox', '--disable-web-security', '--disable-features=site-per-process']
    });
    const page = await browser.newPage();
    console.log("Launched Puppeteer");
    //START OF LOGIN
    await page.goto('https://office.com', { waitUntil: 'networkidle0' });
    await page.click("#hero-banner-sign-in-to-office-365-link"); //Press login
    await page.waitForNavigation({ waitUntil: 'networkidle2' });
    await page.type('#i0116', userdata.teams.username);
    await page.click("#idSIButton9"); //Press next
    await page.waitForNavigation({ waitUntil: 'networkidle0' }); //Wait if automatic login is available
    await page.type('#i0118', userdata.teams.password);
    await page.click("#idSIButton9"); //Press login
    await page.waitForNavigation({ waitUntil: 'networkidle2' });
    await page.click("#idSIButton9"); //Press stayed login
    await page.waitForNavigation({ waitUntil: 'networkidle2' });
    console.log("Office login complete");
    await page.goto('https://teams.microsoft.com/', { waitUntil: 'networkidle2' });
    await page.click(".use-app-lnk"); //Navigate to teams and click on I rather use the webapp
    console.log("Opened teams webapp");
    await page.waitForNavigation({ waitUntil: 'networkidle2' });
    //END OF LOGIN
    //START OF GETTING ASSIGMENTS
    await page.click("#teams-app-bar > ul > li:nth-child(4)"); //Click on assignments
    await page.reload({ waitUntil: "networkidle2" });
    await timeout(5000);
    await page.evaluate(() => {
        document.querySelector("embedded-page-container > div > iframe").contentWindow.document.body.querySelector("div > .desktop-list-padding__3ShvT > div:nth-child(1) > button").click();
        return null;
    });
    console.log("Opened assignments");
    await timeout(5000);
    var iframe = await page.evaluate((sel) => {
        let elements = Array.from(document.querySelector("embedded-page-container > div > iframe").contentWindow.document.body.querySelector(".desktop-list-padding__3ShvT > div:nth-child(2) >div").children);
        let links = elements.map(element => {
            return element.innerHTML;
        });
        return links;
    });
    console.log("Got iFrame content");
    let iframeJsonList = iframe.map(element => {
        var json = html2json(element);
        return json;
    });
    console.log("Converted iFrame content to JSON");
    let tempHomeworkObjectList = iframeJsonList.map(element => {
        return {
            title: element.child[0].child[0].child[0].child[1].child[0].child[0].text,
            details: null,
            class: element.child[0].child[0].child[0].child[2].child[0].text,
            due: element.child[0].child[0].child[0].child[3].child[0].child[2].child[0].text.split("Határidő: ")[1],
        };
    });
    console.log("Made JS objects from JSON");
    return tempHomeworkObjectList;
}

function timeout(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

main();