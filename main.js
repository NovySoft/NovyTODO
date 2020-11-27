const puppeteer = require('puppeteer');
const fs = require('fs');
var html2json = require('html2json').html2json;
const requiem = require("requiem-http");
var qs = require("querystring");
var userData;
var homeworkObjectList = [];

function loadUserData() {
    let rawData = fs.readFileSync("config.json", "utf-8");
    let data = JSON.parse(rawData);
    console.log("Loaded UserData");
    return data;
}

async function main() {
    userData = loadUserData();
    homeworkObjectList = [];
    let temp = await loginAndGetTeamsAssignments(userData);
    homeworkObjectList.push.apply(homeworkObjectList, temp);
    temp = await getKretaAssignments(userData);
    homeworkObjectList.push.apply(homeworkObjectList, temp);
    console.log(homeworkObjectList);
    console.log("DONE");
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
    console.log("Opened office");
    await page.click("#hero-banner-sign-in-to-office-365-link"); //Press login
    console.log("Opened Login Page");
    await page.waitForNavigation({ waitUntil: 'networkidle2' });
    await page.type('#i0116', userdata.teams.username);
    await page.click("#idSIButton9"); //Press next
    console.log("Entered Username");
    await timeout(2000); //Wait if automatic login is available
    await page.type('#i0118', userdata.teams.password);
    await page.click("#idSIButton9"); //Press login
    console.log("Entered Password");
    await page.waitForNavigation({ waitUntil: 'networkidle2' });
    console.log("Pressed Stay Logged In");
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
    console.log("Navigated To Assignments Page");
    await timeout(5000);
    await page.evaluate(() => {
        document.querySelector("embedded-page-container > div > iframe").contentWindow.document.body.querySelector("div > .desktop-list-padding__3ShvT > div:nth-child(1) > button").click();
        return null;
    });
    console.log("Opened all assignments that are not handed in");
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

async function getKretaAssignments(userdata) {
    let today = new Date();
    today.setDate(today.getDate() - 7);
    let todayString = today.getFullYear() + "-" + today.getMonth() + "-" + today.getDate();
    // POST JSON body
    var postData = qs.stringify({
        userName: userdata.kreta.username,
        password: userdata.kreta.password,
        institute_code: userdata.kreta.school,
        grant_type: "password",
        client_id: "kreta-ellenorzo-mobile",
    });
    const postJsonOptions = {
        url: "https://idp.e-kreta.hu/connect/token",
        headers: {
            "User-Agent": "NovyTODO",
            "Content-Type": "application/x-www-form-urlencoded",
        },
        method: "POST",
        body: postData,
    };
    let res = await requiem.requestBody(postJsonOptions);
    let token = parseResponse(res.body.toString("utf8")).access_token;
    console.log("Got Kreta Token");
    res = await requiem
        .requestBody({
            url: "https://" +
                userdata.kreta.school +
                ".e-kreta.hu/ellenorzo/V3/Sajat/HaziFeladatok?datumTol=" +
                todayString,
            headers: {
                "User-Agent": "NovyTODO",
                "Authorization": "Bearer " + token,
            },
        });
    let responseJson = parseResponse(res.body.toString("utf8"));
    console.log("Got Kreta Assignments");
    let tempHomeworkObjectList = responseJson.map(element => {
        let tempTitle = removeTags(element.Szoveg);
        let tempDetail = null;
        if (tempTitle.length > 255) {
            tempTitle = tempTitle.substring(0, 255);
            tempDetail = tempTitle.substring(255);
        }
        return {
            title: tempTitle,
            details: tempDetail,
            class: element.TantargyNeve,
            due: element.HataridoDatuma,
        };
    });
    console.log("Created Kreta Assignments Object");
    return tempHomeworkObjectList;
}

function parseResponse(string) {
    try {
        var o = JSON.parse(string);
        if (o && typeof o === "object") {
            return o;
        }
    } catch (err) {
        console.error(err);
    }
    return string;
}

function timeout(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function removeTags(str) {
    //Remove html tags from text
    if ((str === null) || (str === ''))
        return false;
    else
        str = str.toString();
    return str.replace(/(<([^>]+)>)/ig, '');
}


main();