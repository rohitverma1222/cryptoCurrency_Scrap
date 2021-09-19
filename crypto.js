let puppeteer = require("puppeteer");
let xlsx = require("xlsx");
let fs = require("fs");
let path = require("path");
let page;

(async function fn() {
    let browser = await puppeteer.launch({
        headless: false,
        defaultViewport: null,
        args: ["--start-maximized"],
    });

    page = await browser.newPage();

    await page.goto("https://gadgets.ndtv.com/finance/crypto-currency-price-in-india-inr-compare-bitcoin-ether-dogecoin-ripple-litecoin");
    // await page.waitForSelector("div p")
    await delay(5000);
    await page.waitForSelector('._cptbltr td ._flx._cpnm');

    let array = await page.$$eval("._cptbltr td ._flx._cpnm", el => el.map(x => x.getAttribute("href")));

    for (let i = 0; i < array.length; i++) {

        await page.keyboard.down('Control');
        let newpage = await page.click('a[href="' + array[i] + '"]')

        await page.keyboard.up('Control');
        let p = await browser.pages();
        let n = p[p.length - 1];
        await n.bringToFront();
        await n.waitForSelector("h1");//new tabs opened and wait for h1 tags
        //task is find name and file it into a excel file
        let waitforName = await n.waitForSelector('[class="h1"]')
        let Nameof = await n.evaluate(ele => ele.textContent, waitforName);

        let excelNameOfFile = Nameof;//use this variable for use
        //make a folder//
        let FolderName = path.join(__dirname, Nameof);

        if (fs.existsSync(FolderName) == false)
            fs.mkdirSync(FolderName);

        let excelFile = path.join(FolderName, Nameof + ".xlsx");
        let playerArray = [];


        Nameof = FolderName + '/' + Nameof + '.png';
        let chartwait = await n.waitForSelector("._cptblwrp");
        await n.evaluate(() => {
            document.querySelector("h2._hd").scrollIntoView();
        }, chartwait);
        await delay(2000);


        await n.screenshot({                      // Screenshot the website using defined options

            path: Nameof,
            fullPage: true,
            // Save the screenshot in current directory                             // take a fullpage screenshot
        });
        let ary = await n.$$('._cptbl._cptblm');



        let data = await ary[1].evaluate(ele => ele.textContent, "._flx ._cpnm");
        // let newArray=[];
        let arrayOfdata = data.split(" ");
        let someary = [];
        for (let i = 0; i < arrayOfdata.length; i++) {
            if (arrayOfdata[i] != "" && arrayOfdata[i] != "\n" && arrayOfdata[i] != "₹" && arrayOfdata[i] != 'Date' &&
                arrayOfdata[i] != 'Open' && arrayOfdata[i] != 'High' && arrayOfdata[i] != 'Low' && arrayOfdata[i] != 'Close' && arrayOfdata[i] != 'Volume' && arrayOfdata[i] != 'Change' && arrayOfdata[i] != '(%)') {
                someary.push(arrayOfdata[i]);
            }
        }
        let again = [];
        while(someary.length)
        {
            again.push(someary.splice(0,7));
        }
        

        for (let i = 0; i < again.length; i++) {
            let date = again[i][0];

            let open = again[i][1];

            let high = again[i][2];

            let low = again[i][3];
            let close = again[i][4];
            let volume = again[i][5];
            let change = again[i][6];


            let obj = {
                date,
                open,
                high,
                low,
                close,
                volume,
                change,
            }

            if (fs.existsSync(excelFile) == false) {
                playerArray.push(obj)
            }
            else {
                playerArray = excelReader(excelFile, excelNameOfFile);
                playerArray.push(obj);
            }
            excelWriter(excelFile, playerArray, excelNameOfFile);
        }
        console.log(excelNameOfFile+" Work done")
        await n.close();
    }
})();


const delay = ms => new Promise(res => setTimeout(res, ms));

function excelWriter(filePath, json, sheetName) {
    // workbook create
    let newWB = xlsx.utils.book_new();
    // worksheet
    let newWS = xlsx.utils.json_to_sheet(json);
    xlsx.utils.book_append_sheet(newWB, newWS, sheetName);
    // excel file create 
    xlsx.writeFile(newWB, filePath);
}

function excelReader(filePath, sheetName) {
    // player workbook
    let wb = xlsx.readFile(filePath);
    // get data from a particular sheet in that wb
    let excelData = wb.Sheets[sheetName];
    // sheet to json 
    let ans = xlsx.utils.sheet_to_json(excelData);
    return ans;
}