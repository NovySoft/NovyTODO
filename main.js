const puppeteer = require('puppeteer');
const fs = require('fs');
const html2json = require('html2json').html2json;
const requiem = require("requiem-http");
const qs = require("querystring");
const express = require("express");
const msal = require('@azure/msal-node');
const sqlite3 = require('sqlite3').verbose();
const sqlite = require('sqlite');
const { table } = require('console');
const colors = require('colors');
const SERVER_PORT = 1235;
var microsoftApiToken;
var totalStart = new Date();
var start = new Date();
var rawData = fs.readFileSync("config.json", "utf-8");
var userData = JSON.parse(rawData);
const databasePath = "./db/novytodo.db";
var db;
var headless = true;
var newItems = {
    kreta: 0,
    teams: 0,
};

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
        redirectUri: `http://localhost:${SERVER_PORT}/redirect`,
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
        redirectUri: `http://localhost:${SERVER_PORT}/redirect`,
    };

    pca.acquireTokenByCode(tokenRequest).then((response) => {
        microsoftApiToken = response.accessToken;
        res.sendStatus(200);
    }).catch((error) => {
        console.log(error);
        res.status(500).send(error);
    });
});

function addCustomFunction() {
    //We use hash functions to store id of teams assignments
    Object.defineProperty(String.prototype, 'hashCode', {
        value: function() {
            var hash = 0,
                i, chr;
            for (i = 0; i < this.length; i++) {
                chr = this.charCodeAt(i);
                hash = ((hash << 5) - hash) + chr;
                hash |= 0; // Convert to 32bit integer
            }
            return hash;
        }
    });
    Array.prototype.removeIf = function(callback) {
        var i = this.length;
        while (i--) {
            if (callback(this[i], i)) {
                this.splice(i, 1);
            }
        }
    };
}

async function main() {
    start = new Date();
    let homeworkObjectList = [];
    addCustomFunction();
    await openDatabase();
    let databaseTime = Math.abs(start - new Date());
    start = new Date();
    let temp = await loginAndGetTeamsAssignments(userData);
    homeworkObjectList.push.apply(homeworkObjectList, temp);
    let teamsTime = Math.abs(start - new Date());
    start = new Date();
    temp = await getKretaAssignments(userData);
    homeworkObjectList.push.apply(homeworkObjectList, temp);
    let kretaTime = Math.abs(start - new Date());
    start = new Date();
    console.log("The fetching part is done".yellow);
    await getMSApiToken(userData); //Sets microsoftApiToken variable
    await insertHomeworkTodos(homeworkObjectList);
    let todoTime = Math.abs(start - new Date());
    start = new Date();
    let fullTime = Math.abs(totalStart - new Date());
    console.log("\n");
    console.log("Operation done, have great day".green);
    console.log(("Inserted new items, in total: " + (newItems.kreta + newItems.teams)).magenta);
    console.log(("Inserted kreta items: " + newItems.kreta).magenta);
    console.log(("Inserted teams items: " + newItems.teams).magenta);
    console.log(("Complete operation took: " + fullTime / 60 + "s").blue);
    console.log(("Database time: " + databaseTime / 60 + "s").blue);
    console.log(("Teams time: " + teamsTime / 60 + "s").blue);
    console.log(("Kréta time: " + kretaTime / 60 + "s").blue);
    console.log(("Todo time: " + todoTime / 60 + "s").blue);
}

//This function returns a list with the objects of the teams assignments
async function loginAndGetTeamsAssignments(userdata) {
    let browser = await puppeteer.launch({
        headless: headless,
        //UnSafe args used to get crossSite iframe content
        args: ['--no-sandbox', '--disable-web-security', '--disable-features=site-per-process']
    });
    let page = await browser.newPage();
    console.log("Launched Puppeteer".magenta);
    //START OF LOGIN
    await page.goto('https://office.com', { waitUntil: 'networkidle0', timeout: 0 });
    console.log("Opened office".magenta);
    await page.click("#hero-banner-sign-in-to-office-365-link"); //Press login
    console.log("Opened Login Page".magenta);
    try {
        await page.waitForNavigation({ waitUntil: 'networkidle2' });
    } catch (e) {
        console.log("Navigation timeout continuing".red);
    }
    await page.type('#i0116', userdata.teams.username);
    await page.click("#idSIButton9"); //Press next
    console.log("Entered Username".magenta);
    await timeout(5000); //Wait if automatic login is available
    await page.type('#i0118', userdata.teams.password);
    await page.click("#idSIButton9"); //Press login
    console.log("Entered Password".magenta);
    try {
        await page.waitForNavigation({ waitUntil: 'networkidle2' });
    } catch (e) {
        console.log("Navigation timeout continuing".red);
    }
    console.log("Pressed Stay Logged In".magenta);
    await page.click("#idSIButton9"); //Press stayed login
    try {
        await page.waitForNavigation({ waitUntil: 'networkidle2' });
    } catch (e) {
        console.log("Navigation timeout continuing".red);
    }
    console.log("Office login complete".magenta);
    await page.goto('https://teams.microsoft.com/', { waitUntil: 'networkidle2', timeout: 0 });
    await page.click(".use-app-lnk"); //Navigate to teams and click on I rather use the webapp
    console.log("Opened teams webapp".magenta);
    try {
        await page.waitForNavigation({ waitUntil: 'networkidle2' });
    } catch (e) {
        console.log("Navigation timeout continuing".red);
    }
    //END OF LOGIN
    //START OF GETTING ASSIGMENTS
    await page.click("#teams-app-bar > ul > li:nth-child(4)"); //Click on assignments
    await page.reload({ waitUntil: "networkidle2", timeout: 0 });
    console.log("Navigated To Assignments Page".magenta);
    await timeout(5000);
    let doesExist = await page.evaluate(() => {
        if (document.querySelector("embedded-page-container > div > iframe").contentWindow.document.body.querySelector("div > .desktop-list-padding__3ShvT > div:nth-child(1) > button") != null) {
            document.querySelector("embedded-page-container > div > iframe").contentWindow.document.body.querySelector("div > .desktop-list-padding__3ShvT > div:nth-child(1) > button").click();
            return true;
        } else {
            return false;
        }
    });
    if (doesExist) {
        console.log("Opened all assignments that are not handed in".magenta);
    } else {
        console.log("Seems like you have got no older assignments button, continuing".magenta);
    }
    await timeout(5000);
    var iframe = await page.evaluate((sel) => {
        if (document.querySelector("embedded-page-container > div > iframe").contentWindow.document.body.querySelector(".desktop-list-padding__3ShvT > div:nth-child(2) >div") == null) {
            return true;
        }
        let elements = Array.from(document.querySelector("embedded-page-container > div > iframe").contentWindow.document.body.querySelector(".desktop-list-padding__3ShvT > div:nth-child(2) >div").children);
        let links = elements.map(element => {
            return element.innerHTML;
        });
        return links;
    });
    console.log("Got iFrame content".magenta);
    if (iframe == true) {
        console.log("You have got no teams assignments :)".green);
        return [];
    }
    let iframeJsonList = iframe.map(element => {
        var json = html2json(element);
        return json;
    });
    console.log("Converted iFrame content to JSON".magenta);
    let tempHomeworkObjectList = iframeJsonList.map(element => {
        let tempTitle = element.child[0].child[0].child[0].child[1].child[0].child[0].text;
        let tempClass = element.child[0].child[0].child[0].child[2].child[0].text;
        let due = element.child[0].child[0].child[0].child[3].child[0].child[2].child[0].text;
        if ((due.split("Határidő: ")[1] != null || due.split("Határidő: ")[1] != undefined) &&
            !(due.includes("holnap") || due.includes("tomorrow") ||
                (due.includes("ma") || due.includes("today")))) {
            due = due.split("Határidő: ")[1];
            due = Date.parse(due);
            due = new Date(due).toISOString();
        } else {
            //sometimes teams uses english
            //sometimes it uses hungarion
            //Can't figure out why
            if (due.includes("holnap") || due.includes("tomorrow")) {
                const today = new Date();
                const tomorrow = new Date(today);
                tomorrow.setDate(tomorrow.getDate() + 1);
                let hourminutes = null;
                if (due.includes("holnap")) {
                    hourminutes = due.split("Határidő holnap ekkor: ")[1]; //Időpont, de csak óra és perc
                } else if (due.includes("tomorrow")) {
                    hourminutes = getTwentyFourHourTime(due.split("Due tomorrow at ")[1]);
                }
                if (hourminutes == null) {
                    console.error(`Could determinate due date of this (${"t" + (tempClass + tempTitle).hashCode()}) object, continuing with other objects`);
                    return;
                }
                tomorrow.setHours(hourminutes.split(":")[0]);
                tomorrow.setMinutes(hourminutes.split(":")[1]);
                tomorrow.setSeconds(0);
                tomorrow.setMilliseconds(0);
                due = tomorrow.toISOString();
            } else if (due.includes("ma") || due.includes("today")) {
                const today = new Date();
                let hourminutes = null;
                if (due.includes("ma")) {
                    hourminutes = due.split("Határidő ma ekkor: ")[1]; //Időpont, de csak óra és perc
                } else if (due.includes("tomorrow")) {
                    hourminutes = getTwentyFourHourTime(due.split("Due today at ")[1]);
                }
                if (hourminutes == null) {
                    console.error(`Could determinate due date of this (${"t" + (tempClass + tempTitle).hashCode()}) object, continuing with other objects`);
                    return;
                }
                today.setHours(hourminutes.split(":")[0]);
                today.setMinutes(hourminutes.split(":")[1]);
                today.setSeconds(0);
                today.setMilliseconds(0);
                due = today.toISOString();
            }
            //FIXME: Lehet tegnap érték is?
        }
        return {
            id: "t" + (tempClass + tempTitle).hashCode(), //t for teams
            title: tempTitle,
            details: '',
            class: tempClass,
            due: due,
        };
    });
    console.log("Made JS objects from JSON".green);
    await page.close();
    await browser.close();
    return tempHomeworkObjectList;
}

function getTwentyFourHourTime(input) {
    var time = input;
    var hours = Number(time.match(/^(\d+)/)[1]);
    var minutes = Number(time.match(/:(\d+)/)[1]);
    var AMPM = time.match(/\s(.*)$/)[1];
    if (AMPM == "PM" && hours < 12) hours = hours + 12;
    if (AMPM == "AM" && hours == 12) hours = hours - 12;
    var sHours = hours.toString();
    var sMinutes = minutes.toString();
    if (hours < 10) sHours = "0" + sHours;
    if (minutes < 10) sMinutes = "0" + sMinutes;
    return (sHours + ":" + sMinutes);
}

async function getKretaAssignments(userdata) {
    let today = new Date();
    today.setDate(today.getDate() - 7);
    let todayString = today.toISOString();
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
    console.log("Got Kreta Token".magenta);
    res = await requiem
        .requestBody({
            url: `https://${userdata.kreta.school}.e-kreta.hu/ellenorzo/V3/Sajat/HaziFeladatok?datumTol=${todayString}`,
            headers: {
                "User-Agent": "NovyTODO",
                "Authorization": "Bearer " + token,
            },
        });
    let responseJson = parseResponse(res.body.toString("utf8"));
    console.log("Got Kreta Assignments".magenta);
    let tempHomeworkObjectList = responseJson.map(element => {
        let tempTitle = removeTags(element.Szoveg);
        let tempDetail = '';
        if (tempTitle.length + element.TantargyNeve > 250) {
            tempTitle = tempTitle.substring(0, 250 - element.TantargyNeve - 1);
            tempDetail = tempTitle.substring(250 - element.TantargyNeve - 1);
        }
        return {
            id: "k" + element.Uid, //K for kréta
            title: tempTitle,
            details: tempDetail,
            class: element.TantargyNeve,
            due: element.HataridoDatuma,
        };
    });
    console.log("Created Kreta Assignments Object".green);
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
    let server = app.listen(SERVER_PORT, () => console.log('Express Login Server Started'.magenta));
    let browser = await puppeteer.launch({
        headless: headless,
    });
    let page = await browser.newPage();
    await page.goto(`http://localhost:${SERVER_PORT}/`, { waitUntil: 'networkidle0', timeout: 0 });
    await timeout(300);
    await page.type('#i0116', userdata.todo.username);
    await page.click("#idSIButton9"); //Press next
    console.log("Entered Username".magenta);
    try {
        await page.waitForNavigation({ waitUntil: 'networkidle2' });
    } catch (e) {
        console.log("Navigation timeout continuing".red);
    }
    await page.type('#i0118', userdata.todo.password);
    await page.click("#idSIButton9"); //Press login
    console.log("Entered Password".magenta);
    try {
        await page.waitForNavigation({ waitUntil: 'networkidle0' });
    } catch (e) {
        console.log("Navigation timeout continuing".red);
    }
    //!NOTICE FIRST YOU SHOULD MANUALLY ALLOW YOUR APPLICATION
    //!You musn't click stay signed in!
    //*The code cannot decide whether that is necessary
    server.close();
    await page.close();
    await browser.close();
    if (microsoftApiToken == null) {
        throw Error("Token wasn't found");
    } else {
        console.log("Express Login Server Closed + Api Login Successful".green);
    }
}

async function insertHomeworkTodos(homeworks) {
    let listId = await getTODOListId();
    //let items = await getTODOitems(listId);
    let alreadyInsertedItems = await getAlreadyInsertedItems();
    for (var item of homeworks) {
        let existList = alreadyInsertedItems.filter(e => e.todoId === item.id);
        if (existList.length > 1) {
            throw Error("Exist List Contains more than one match");
        }
        if (existList.length > 0) {
            console.log(`Element ${item.id} (${item.class}) already inserted, should update`.yellow);
            //FIXME make updates work
        } else {
            let result = await insertNewMsTODO(item, listId);
            if (result != true) {
                console.log("");
                console.error("TODO request failed. This is what I got in response:".red);
                console.error(result);
                console.log("");
            } else {
                await insertIntoDB(item);
                console.log(`Inserted ${item.id} (${item.class}) into todo and database`.green);
            }
        }
    }
}

async function insertNewMsTODO(todo, taskListId) {
    let dueDate = Date.parse(todo.due);
    var reminderDateString = new Date(dueDate);
    reminderDateString.setDate(reminderDateString.getDate() - 3); //Get notified 3 days prior
    reminderDateString.setHours(12); //I like getting notfications at 12 o'clock
    reminderDateString = reminderDateString.toISOString();
    var postData = JSON.stringify({
        title: todo.class + " " + todo.title,
        body: {
            content: todo.details + "   -Made with novyTODO",
            contentType: "text"
        },
        dueDateTime: {
            dateTime: todo.due,
            timeZone: "UTC",
        },
        isReminderOn: true,
        reminderDateTime: {
            dateTime: reminderDateString,
            timeZone: "UTC",
        },
    });
    const postJsonOptions = {
        url: `https://graph.microsoft.com/v1.0/me/todo/lists/${taskListId}/tasks`,
        headers: {
            "User-Agent": "NovyTODO",
            "Authorization": "Bearer " + microsoftApiToken,
            "Content-Type": "application/json",
        },
        method: "POST",
        body: postData,
    };
    let res = await requiem.requestBody(postJsonOptions);
    if (res.statusCode == 200 || res.statusCode == 201) {
        return true;
    } else {
        return res.body.toString();
    }
}

async function getAlreadyInsertedItems() {
    let result = await db.all("SELECT * FROM todo_data");
    return result;
}

//Takes an hwObject as input
async function insertIntoDB(input) {
    if (input.id.startsWith("k")) {
        newItems.kreta += 1;
    } else {
        newItems.teams += 1;
    }
    await db.run("INSERT INTO todo_data (todoId, title, details, class, due) VALUES (:id, :title, :details, :class, :due)", {
        ':id': input.id,
        ':title': input.title,
        ':details': input.details,
        ':class': input.class,
        ':due': input.due,
    });
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

async function getTODOitems(id) {
    //TODO: Add counters for how many items you had done since last runtime
    let res = await requiem
        .requestBody({
            url: `https://graph.microsoft.com/v1.0/me/todo/lists/${id}/tasks`,
            headers: {
                "User-Agent": "NovyTODO",
                "Authorization": "Bearer " + microsoftApiToken,
            },
        });
    let responseJson = parseResponse(res.body.toString("utf8"));
    console.log("Got current todos in default list");
    return responseJson;
}

async function openDatabase() {
    if (!fs.existsSync(databasePath)) {
        console.warn("Database not found, making one");
    }
    db = new sqlite3.Database(databasePath, (err) => {
        if (err) {
            console.error(err.message);
            throw Error(err);
        }
    });
    db.close();
    db = await sqlite.open({
        filename: databasePath,
        driver: sqlite3.cached.Database,
    });
    await db.run("CREATE TABLE IF NOT EXISTS todo_data (databaseId INTEGER PRIMARY KEY,todoId TEXT,title TEXT, details TEXT, class TEXT, due TEXT)");
    console.log("Opened Databased".green);
}


main();