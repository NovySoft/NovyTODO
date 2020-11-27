const puppeteer = require('puppeteer');
const fs = require('fs');
const html2json = require('html2json').html2json;
const requiem = require("requiem-http");
const qs = require("querystring");
const express = require("express");
const msal = require('@azure/msal-node');
const { table } = require('console');
const SERVER_PORT = 3000;
var microsoftApiToken;
var rawData = fs.readFileSync("config.json", "utf-8");
var userData = JSON.parse(rawData);

const config = {
    auth: {
        clientId: userData.api.clientId,
        authority: "https://login.microsoftonline.com/common",
        clientSecret: userData.api.clientSecret,
    },
    system: { loggerOptions: { loggerCallback(loglevel, message, containsPii) { console.log(message); }, piiLoggingEnabled: false, logLevel: msal.LogLevel.Verbose, } }
};

// Create msal application object
const pca = new msal.ConfidentialClientApplication(config);
const app = express();
app.disable("x-powered-by");

app.get('/', (req, res) => {
    const authCodeUrlParameters = {
        scopes: ["user.read", "tasks.readwrite"],
        redirectUri: "http://localhost:3000/redirect",
    };

    // get url to sign user in and consent to scopes needed for application
    pca.getAuthCodeUrl(authCodeUrlParameters).then((response) => {
        res.redirect(response);
    }).catch((error) => console.log(JSON.stringify(error)));
});

app.get('/redirect', (req, res) => {
    const tokenRequest = {
        code: req.query.code,
        scopes: ["user.read"],
        redirectUri: "http://localhost:3000/redirect",
    };

    pca.acquireTokenByCode(tokenRequest).then((response) => {
        microsoftApiToken = response.accessToken;
        res.sendStatus(200);
    }).catch((error) => {
        console.log(error);
        res.status(500).send(error);
    });
});

async function main() {
    let homeworkObjectList = [];
    let temp = await loginAndGetTeamsAssignments(userData);
    //Todo: run code async to save some time
    homeworkObjectList.push.apply(homeworkObjectList, temp);
    temp = await getKretaAssignments(userData);
    homeworkObjectList.push.apply(homeworkObjectList, temp);
    console.log("The fetching part is done");
    await getMSApiToken(userData); //Sets microsoftApiToken variable
    await insertHomeworkTodos([]);
    console.log("Operation done, have great day");
    console.log("Complete operation took: "); //Todo: impletement timers
}

//This function returns a list with the objects of the teams assignments
async function loginAndGetTeamsAssignments(userdata) {
    let browser = await puppeteer.launch({
        headless: false,
        //UnSafe args used to get crossSite iframe content
        args: ['--no-sandbox', '--disable-web-security', '--disable-features=site-per-process']
    });
    let page = await browser.newPage();
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
    await page.close();
    await browser.close();
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

async function getMSApiToken(userdata) {
    let server = app.listen(SERVER_PORT, () => console.log('Express Login Server Started'));
    let browser = await puppeteer.launch({
        headless: false,
    });
    let page = await browser.newPage();
    await page.goto('http://localhost:3000/', { waitUntil: 'networkidle0' });
    await timeout(300);
    await page.type('#i0116', userdata.todo.username);
    await page.click("#idSIButton9"); //Press next
    console.log("Entered Username");
    await page.waitForNavigation({ waitUntil: 'networkidle2' });
    await page.type('#i0118', userdata.todo.password);
    await page.click("#idSIButton9"); //Press login
    console.log("Entered Password");
    await page.waitForNavigation({ waitUntil: 'networkidle0' });
    //!NOTICE FIRST YOU SHOULD MANUALLY ALLOW YOUR APPLICATION
    //!You musn't click stay signed in!
    //*The code cannot decide whether that is necessary
    server.close();
    await page.close();
    await browser.close();
    if (microsoftApiToken == null) {
        throw Error("Token wasn't found");
    } else {
        console.log("Express Login Server Closed + Api Login Successful");
    }
}

async function insertHomeworkTodos(homeworks) {
    let listId = await getTODOListId();
    console.log(listId);
}

//This function grabs the id of the todo list we are going to push into
async function getTODOListId() {
    let res = await requiem
        .requestBody({
            url: "https://graph.microsoft.com/v1.0/me/todo/lists",
            headers: {
                "User-Agent": "NovyTODO",
                "Authorization": "Bearer " + microsoftApiToken,
            },
        });
    let responseJson = parseResponse(res.body.toString("utf8"));
    let returnIdList = responseJson.value.filter(element => {
        return element.wellknownListName == "defaultList";
    });
    if (returnIdList.length == 0) {
        throw new Error("Couldn't find default list");
    } else if (returnIdList.length > 1) {
        console.warn("More than one defult list was found, using the first one");
    }
    return returnIdList[0].id;
}


main();