const { next } = require('cheerio/lib/api/traversing');
var express = require('express');
var router = express.Router();
var opn = require('opn');
XLSXParser = require("xlsx-to-json");
const { chromium } = require('playwright');
const puppeteer = require('puppeteer');
let XLSX = require('xlsx');

// 
require('dotenv').config();

/* GET home page. */
router.get('/', function(req, res, next) {
  res.render('index', { title: 'Express' });
});

router.post('/search', function(req, res, next){
  // form情報
  var search_condition = {
    company_name: req.body.company_name,
    booking_id: req.body.booking_id,
    container_id: req.body.container_id
  };
  
  console.log(search_condition);

  let sampleFile;
  let uploadPath;

  // booking number
  var url = "";
  switch (search_condition.company_name) {
    case "OOCL":
      url = "https://www.oocl.com/jpn/Pages/default.aspx";
      break;
  
    default:
      break;
  }

  if (search_condition.booking_id) {
    console.log(process.env.BOOKING_ID);
    //const { chromium } = require('playwright');
    /*
    (
      async()=>{
        //const userDataDir = 'public';
        //const context = await chromium.launchPersistentContext(userDataDir, { headless: false });
        //const page = await context.newPage();
        
        const browser = await chromium.launch({
          channel: 'msedge',
          headless: false
        });
        const context = await browser.newContext();
        const page = await context.newPage();
        //var url_new = "https://moc.oocl.com/party/cargotracking/ct_search_from_other_domain.jsf?ANONYMOUS_BEHAVIOR=BUILD_UP&domainName=PARTY_DOMAIN&ENTRY_TYPE=OOCL&ENTRY=MCC&ctSearchType=BC&ctShipmentNumber=";
        //url_new = url_new + search_condition.booking_id;
        //await page.goto(url_new, {waitUntil: 'domcontentloaded'});

        await page.goto(url, {waitUntil: 'domcontentloaded'});
        
        
        const browser = await puppeteer.launch({
          executablePath: 'C:/Program Files/Google/Chrome/Application/chrome',
          headless: false
        });
        
        //const context = await browser.newContext();
        const page = await browser.newPage();
        await page.goto(url, {waitUntil: 'domcontentloaded'});
        
      }
    )();
    */
    
    (
      async()=>{

        const browser = await chromium.launch({
          headless: false
        });
        const page = await browser.newPage();
        await page.goto(url, {waitUntil: 'domcontentloaded'});
        await page.waitForSelector('#SEARCH_NUMBER');
        // Click button[role="button"]:has-text("ブッキング番号")
        //await page.click("button[role=\"button\"]:has-text(\"ブッキング番号\")")
        // Click text=ブッキング番号
        await page.click("text=ブッキング番号")

        //await page.type('#SEARCH_NUMBER', search_condition.booking_id);
        await page.type('#SEARCH_NUMBER', search_condition.booking_id);
        
        await page.click('#container_btn');
        
        //const userDataDir = 'public';
        //const context = await chromium.launchPersistentContext(userDataDir, { headless: false });
    }
    )();
    
  }

  // booking No. file upload
  if (url != "" && req.files && Object.keys(req.files).length !== 0) {
    /*
    console.log(`public/images/search-${search_condition.company_name}.png`);
    await page.goto(url);
    //await page.screenshot({path: `public/images/search-${search_condition.company_name}.png`});
    // Click #bottomCookieNotice button:has-text("Accept")
    await page.click('#bottomCookieNotice button:has-text("Accept")');

    // Click input[name="SEARCH_NUMBER"]
    await page.click('input[name="SEARCH_NUMBER"]');

    // Fill input[name="SEARCH_NUMBER"]
    await page.fill('input[name="SEARCH_NUMBER"]', '2673572460');

    // Press Enter
    const [page1] = await Promise.all([
      page.waitForEvent('popup'),
      page.waitForNavigation({ url: 'https://moc.oocl.com/party/cargotracking/ct_search_from_other_domain.jsf?ANONYMOUS_BEHAVIOR=BUILD_UP&domainName=PARTY_DOMAIN&ENTRY_TYPE=OOCL&ENTRY=MCC&ctSearchType=BC&ctShipmentNumber=2673572460' }),
      page.press('input[name="SEARCH_NUMBER"]', 'Enter')
    ]);

    await browser.close();
    */
    // The name of the input field (i.e. "excel_booking") is used to retrieve the uploaded file
    sampleFile = req.files.excel_booking;
    uploadPath = 'public/' + sampleFile.name;
    console.log(sampleFile.data);

    sampleFile.mv(uploadPath, function(err) {
      if (err)
        return res.status(500).send(err);

        // cellDates（日付セルの保持形式を指定）
        // false：数値（シリアル値）[default]
        // true ：日付フォーマット
        let container_excel = XLSX.readFile(uploadPath, {cellDates:true});
        let sheet1, sheet2, sheet3;
        container_excel.SheetNames.forEach(sheet => {
          if (sheet == "booking_info") {
            sheet1 = XLSX.utils.sheet_to_json(container_excel.Sheets[sheet]);
          }
        });
        console.log(sheet1);
        //const sheet1_obj = JSON.parse(sheet1);
        var booking_Nums = [];
        for (let i = 0; i < sheet1.length; i++) {
          booking_Nums.push(sheet1[i]['No.']);  
        }
        var comp_booking_Nums = {
          'company_name': search_condition.company_name,
          'booking_Nums': booking_Nums
        };
        return res.render('booking_result', {table_booking_Nums: comp_booking_Nums});
      });
    
    /*
    var booking_Nums = ['2672869920', '2672870460', '2672869960', '2673572460', '2674150930', '2675457610', '2676042740', '2676302540', '2677976770', '2677977380'];
    var comp_booking_Nums = {
      'company_name': search_condition.company_name,
      'booking_Nums': booking_Nums
    };
    */
    //url = "https://moc.oocl.com/party/cargotracking/ct_search_from_other_domain.jsf?ANONYMOUS_BEHAVIOR=BUILD_UP&domainName=PARTY_DOMAIN&ENTRY_TYPE=OOCL&ENTRY=MCC&ctSearchType=BC&ctShipmentNumber=";
    //url = url + search_condition.booking_id;
    //url = url + item;
    //res.redirect(url);

    //return res.render('booking_result', {table_booking_Nums: comp_booking_Nums}); 
  }
    //console.log(search_condition.container_id);
  if (search_condition.container_id) {
    var container_trackers=[];
    var container_id = search_condition.container_id;
    (async()=>{
      const browser = await chromium.launch({
        headless: false
      });
      const page = await browser.newPage();
        // Go to https://www.msc.com/track-a-shipment?agencyPath=jpn
      //await page.goto('https://www.msc.com/track-a-shipment?agencyPath=jpn', {waitUntil: 'domcontentloaded'});
      await page.goto('https://www.msc.com/track-a-shipment?agencyPath=gbr', {waitUntil: 'domcontentloaded'});
      //await page.goto('https://www.msc.com/track-a-shipment?lang=fr-fr', {waitUntil: 'domcontentloaded'});
      // Click text=Booking number / Container number / Bill of lading Please input a tracking numbe >> [placeholder="Booking number / Container number / Bill of lading"]
      //await page.click('text=Booking number / Container number / Bill of lading Please input a tracking numbe >> [placeholder="Booking number / Container number / Bill of lading"]');
      await page.click('#ctl00_ctl00_plcMain_plcMain_TrackSearch_txtBolSearch_TextField');
      // Fill text=Booking number / Container number / Bill of lading Please input a tracking numbe >> [placeholder="Booking number / Container number / Bill of lading"]
      //await page.fill('text=Booking number / Container number / Bill of lading Please input a tracking numbe >> [placeholder="Booking number / Container number / Bill of lading"]', container_id);
      await page.fill('#ctl00_ctl00_plcMain_plcMain_TrackSearch_txtBolSearch_TextField', container_id);
      // Press Enter
      await Promise.all([
        page.waitForNavigation({waitUntil: 'domcontentloaded'}/*{ url: 'https://www.msc.com/track-a-shipment?link=9f1cf13f-6f91-4062-9053-8496fd971398' }*/),
        //page.press('text=Booking number / Container number / Bill of lading Please input a tracking numbe >> [placeholder="Booking number / Container number / Bill of lading"]', 'Enter')
        await page.press('#ctl00_ctl00_plcMain_plcMain_TrackSearch_txtBolSearch_TextField', 'Enter')
      ]);
      
      let table_2 = await page.waitForSelector('table:nth-child(4)');

      //console.log(future_tr);
      //var table2_string = await table_2.innerHTML();
      /*
      await future_tr.$$eval('.responsiveTd', (nodes) => {
        nodes.forEach(node => {
          info.push(node.innerText);
        })
      });*/
      //console.log(table2_string);
      
      var table_info = await table_2.$$eval('tr', nodes => nodes.map(n => n.innerText));
      var table_info_arr = table_info.map(ti => ti.split('\t'));
      console.log(table_info_arr);
      
      let container_tracker = {
        'container_id': container_id,
        'track_info': table_info_arr
      };
      container_trackers.push(container_tracker);
      // Close page
      await page.close();
      console.log("success!");

      return res.render('container_result', {table_tarck_info: container_trackers});
    })();
  }
  if (req.body.company_name == "" && req.files && Object.keys(req.files).length !== 0) {
    //return res.status(400).send('No files were uploaded.');
  
    // The name of the input field (i.e. "excel_container") is used to retrieve the uploaded file
    sampleFile = req.files.excel_container;
    uploadPath = 'public/' + sampleFile.name;
    console.log(sampleFile.data);

    /*
    const jsonFromXLSX = new XLSXParser()
    .xlsx_to_json(
      sampleFile.data, // file/blob/base64/string/array/buffer as xlsx to convert to json
      null // options from require('xlsx').xlsx_to_json()
    );
    console.log(jsonFromXLSX);
      */
    // Use the mv() method to place the file somewhere on your server
    sampleFile.mv(uploadPath, function(err) {
      if (err)
        return res.status(500).send(err);

      //res.send('File uploaded!');
      // cellDates（日付セルの保持形式を指定）
      // false：数値（シリアル値）[default]
      // true ：日付フォーマット
      let container_excel = XLSX.readFile(uploadPath, {cellDates:true});
      let sheet1, sheet2, sheet3;
      container_excel.SheetNames.forEach(sheet => {
        if (sheet == "container_info") {
          sheet1 = XLSX.utils.sheet_to_json(container_excel.Sheets[sheet]);
        }
      });
      console.log(sheet1);
      var container_trackers=[];
      (async()=>{
        await Promise.all(
          sheet1.map(async(item)=>{
          let container_id = item['No.'];
          try {
              
            const browser = await chromium.launch({
              headless: true
            });
            const page = await browser.newPage();
              // Go to https://www.msc.com/track-a-shipment?agencyPath=jpn
            await page.goto('https://www.msc.com/track-a-shipment?agencyPath=jpn', {waitUntil: 'domcontentloaded'});
            //await page.goto('https://www.msc.com/track-a-shipment?agencyPath=gbr', {waitUntil: 'domcontentloaded'});
            //await page.goto('https://www.msc.com/track-a-shipment?lang=fr-fr', {waitUntil: 'domcontentloaded'});
            // Click text=Booking number / Container number / Bill of lading Please input a tracking numbe >> [placeholder="Booking number / Container number / Bill of lading"]
            //await page.click('text=Booking number / Container number / Bill of lading Please input a tracking numbe >> [placeholder="Booking number / Container number / Bill of lading"]');
            await page.click('#ctl00_ctl00_plcMain_plcMain_TrackSearch_txtBolSearch_TextField');
            // Fill text=Booking number / Container number / Bill of lading Please input a tracking numbe >> [placeholder="Booking number / Container number / Bill of lading"]
            //await page.fill('text=Booking number / Container number / Bill of lading Please input a tracking numbe >> [placeholder="Booking number / Container number / Bill of lading"]', container_id);
            await page.fill('#ctl00_ctl00_plcMain_plcMain_TrackSearch_txtBolSearch_TextField', container_id);
            // Press Enter
            await Promise.all([
              page.waitForNavigation({waitUntil: 'domcontentloaded'}/*{ url: 'https://www.msc.com/track-a-shipment?link=9f1cf13f-6f91-4062-9053-8496fd971398' }*/),
              //page.press('text=Booking number / Container number / Bill of lading Please input a tracking numbe >> [placeholder="Booking number / Container number / Bill of lading"]', 'Enter')
              await page.press('#ctl00_ctl00_plcMain_plcMain_TrackSearch_txtBolSearch_TextField', 'Enter')
            ]);
            
            let table_2 = await page.waitForSelector('table:nth-child(4)');

            //console.log(future_tr);
            //var table2_string = await table_2.innerHTML();
            /*
            await future_tr.$$eval('.responsiveTd', (nodes) => {
              nodes.forEach(node => {
                info.push(node.innerText);
              })
            });*/
            //console.log(table2_string);
            
            var table_info = await table_2.$$eval('tr', nodes => nodes.map(n => n.innerText));
            var table_info_arr = table_info.map(ti => ti.split('\t'));
            //console.log(table_info_arr);
            let container_tracker = {
              'container_id': container_id,
              'track_info': table_info_arr
            };
            
            // vesselとvoyage情報ない行を削除
            const eta_table = table_info_arr.filter(word => word[3].length > 1);
            console.log(eta_table);
            const length = eta_table.length;
            console.log(length);
            for (let i = 0; i < eta_table.length; i++) {
              switch (eta_table[i][1]) {
                case "Export Loaded on Vessel":
                  item['POL'] = eta_table[i][0];
                  item['ETD'] = eta_table[i][2];
                  break;
                case "Import Discharged from Vessel":
                  item['POD'] = eta_table[i][0];
                  item['ETA'] = eta_table[i][2];
                  break;
                case "Estimated Time of Arrival":
                  item['POD'] = eta_table[i][0];
                  item['ETA'] = eta_table[i][2];
                  break;
                default:
                  break;
              }
            }
            /*
            if (length > 1) {
              item['POL'] = eta_table[length-1][0];
              item['ETD'] = eta_table[length-1][2];
              item['POD'] = eta_table[1][0];
              item['ETA'] = eta_table[1][2];
            }*/
            item['Latest Movement'] = table_info_arr[1][0] + " - " +  table_info_arr[1][1] + " - " + table_info_arr[1][2];
            container_trackers.push(container_tracker);
            // Close page
            await page.close();
            //console.log(container_tracker);
            //console.log("trackers:" + container_trackers);
            console.log("success!");
          } catch (error) {
              throw error;
          }
        }));

        console.log("#############################");
        console.log(container_trackers.length);
        
        // 更新したsheet(PODとETAを追加する)
        console.log(sheet1);
        let exportExcel = XLSX.utils.book_new();
        let exportSheet = XLSX.utils.json_to_sheet(sheet1);
        XLSX.utils.book_append_sheet(exportExcel, exportSheet, "tracking_info");
        XLSX.writeFile(exportExcel, "public/container_detail_new.xlsx");

        return res.render('container_result', {table_tarck_info: container_trackers});
      })();

    });
}
});

router.get('/search/:booking_id', (req, res, next) => {
  const booking_id = req.params.booking_id;

  var url = "https://moc.oocl.com/party/cargotracking/ct_search_from_other_domain.jsf?ANONYMOUS_BEHAVIOR=BUILD_UP&domainName=PARTY_DOMAIN&ENTRY_TYPE=OOCL&ENTRY=MCC&ctSearchType=BC&ctShipmentNumber=";
  url = url + booking_id;
  res.redirect(url);
});

module.exports = router;
