const puppeteer = require('puppeteer');
const XLSX = require('xlsx-color');
const os = require('os');
const fs = require('fs');
const ExcelJS = require('exceljs');
const { request } = require('http');
const { error } = require('console');
const { join } = require('path');

const homeDir = os.homedir();
const excelPath = homeDir + "\\Desktop\\HeadlessBrowser\\AutomateSendingOutTaxReturnsV3.xlsx";
console.log("Excel Path: " + excelPath);
let clientBook = XLSX.readFile(excelPath);

let clientText = "ClientList"
let clientSheet = clientBook.Sheets[clientText];
let clientArray = XLSX.utils.sheet_to_json(clientSheet);
console.log(clientArray);

let summaryText = "ClientSummary";
delete clientBook.Sheets[summaryText];
let summarySheet = clientBook.Sheets[summaryText];
XLSX.writeFile(clientBook, excelPath);




const breakPoint = "----------------------------------------------------------------------------------";
let emailTemplate = "Send Out Returns, Bill, and Efile Sigs - InvoiceInTd";
// let emailTemplate = "Test CRM Code";
const timeout = 3000;
const year = 2022;


const estimate = "Estimate";
const taxReturn = "Tax Return";
const signature = "Signature";
const invoice = "Invoice";
const clientUpload = "Client uploaded documents";

const sealExcel = "Seal"
const estimateExcel = "Estimate";
const taxReturnExcel = "TaxReturn";
const signatureExcel = "Signature";
const invoiceExcel = "Invoice";
const closerExcel = "Closer";
const sealYearExcel = "YearToSeal";


const screenResolution = { width: 1500, height: 800 }; // Replace with your screen's resolution

async function startAndLogin(){
    const browser = await puppeteer.launch({ headless: false, args: ['--start-maximized'] }); // Launch a headless browser
    const page = await browser.newPage(); // Open a new page'
    await page.setViewport(screenResolution);


    await page.goto('https://portal.greenoakfinancial.com/login'); // Navigate to Login Page
  
    try { //Login to GOF
      await page.type('input#email', 'will2828@purdue.edu');
      await page.type('input#password', '$zM{8$;@Z<+8@O=UU;><');
      await page.click('button[type="submit"]');
    } catch (error) { //If already logged in, skip and login
      console.error(error);
      console.log("Already Logged In.");
      console.log("Continuing to Webpage...");
    }
    await page.waitForNavigation( {timeout: 120000} );  
    return {browser, page};
}

async function sendOutEmail(page, i){

  console.log("SEND OUT EMAIL");

  const setup = await precheck(page, i);

  for (j = 1; j < 3; j++){
    emailTempName = 'EmailTemp' + j;
    // emailSentName = 'EmailSent' + j;
    console.log("Checking Column " + emailTempName + " for template...");
    emailTemplate = clientArray[i][emailTempName];
    if(emailTemplate.toLocaleLowerCase() != 'none'){
      console.log(emailTemplate + " Selected as Template.");
      try{ //Send Email
        if(setup === clientArray[i].ClientName){ //check names match
          console.log("Names Match.");
          console.log("Sending Email...");
          await page.goto('https://portal.greenoakfinancial.com/app/clients/' + clientArray[i].ClientID + '/mailbox/inbox'); // Navigate to Client Inbox
          await page.waitForSelector('div.block-container .btn');
          await page.click('div.block-container .btn');
    
          await page.waitForSelector('input#react-select-4-input');
          await page.type('input#react-select-4-input', emailTemplate);
    
          await new Promise(resolve => setTimeout(resolve, 1000)); //1 sec delay
          await page.keyboard.press('Enter');
          await new Promise(resolve => setTimeout(resolve, 1000)); //1 sec delay
    
    
          if(emailTemplate.toLocaleLowerCase() === "test crm code"){
            await page.waitForSelector('input#subject-input');
            await page.type('input#subject-input', " Email " + i);
            await new Promise(resolve => setTimeout(resolve, 2000)); //2 sec delay
          }
          
          await page.keyboard.press('Enter');
          await new Promise(resolve => setTimeout(resolve, 5000)); // Delay for 5000 milliseconds (5 seconds)
          
          console.log("Email Sent.");
          clientArray[i][emailTempName] = "Y";
        } else {
          console.log(clientArray[i].ClientName + " does not match ClientID.");
          console.log("Moving to next client...");
          clientArray[i][emailTempName] = "N";
          continue;
        }
      } catch (error) { //catch case if email is not send or error occurs
        console.log("Email Send Failed.");
        console.error(error);
        clientArray[i][emailTempName] = "N";
        continue;
      }
    } else {
      console.log("Email Template " + j + " does not exist.");
      clientArray[i][emailTempName] = "N";
      continue;
    }
    
  }
  console.log("EMAIL SEND COMPLETE");

}

async function sendOutDocuments(page, i){

  console.log("SEND OUT DOCUMENTS");
  
  const setup = await precheck(page, i);

  if(setup === clientArray[i].ClientName){ //check names match

    console.log("Names Match.");

    //Begin Standard Three Document Selection
    //---------------------------------------------------------------------------------------------------------//

    let documentsToShare = [
      {
        docName: estimate,
        docExcel: estimateExcel,
      },
      {
        docName: taxReturn,
        docExcel: taxReturnExcel,
      },
    ];

    const navigateToYear = await findYearDocuments(page, i, year, documentsToShare);
    const documentSend = await shareMultipleDocuments(page, i, navigateToYear, documentsToShare);
    if(!documentSend){
      for(let j = 0; j < documentsToShare.length; j++){
        clientArray[i][documentsToShare[j].docExcel] = "N";
      }
    }

  

    // //Begin Estimate Document Send 
    // const estimateArray = await findYearCreateArray(page, i, year, estimateExcel);
    // const estimateDocumentShare = await shareDocument(page, i, estimateArray, estimate);
    // if(estimateDocumentShare){
    //   clientArray[i][estimateExcel] = "Y";
    // } else {
    //   clientArray[i][estimateExcel] = "N";
    // }
    // //Estimate File Send Complete



    // //Begin Tax Return Document Send
    // const taxReturnArray = await findYearCreateArray(page, i, year, taxReturnExcel);
    // const taxReturnDocumentShare = await shareDocument(page, i, taxReturnArray, taxReturn);
    // if(taxReturnDocumentShare){
    //   clientArray[i][taxReturnExcel] = "Y";
    // } else {
    //   clientArray[i][taxReturnExcel] = "N";
    // }
    // //Tax Returns File Send Complete



    //Begin Signature Page Request
    const signatureArray = await findYearCreateArray(page, i, year, signatureExcel);
    const signatureRequest = await requestSignatures(page, i, signatureArray, signature);
    if(signatureRequest){
      clientArray[i][signatureExcel] = "Y";
    } else {
      clientArray[i][signatureExcel] = "N";
    }

    console.log("DOCUMENT SEND COMPLETE");
    return true;

  } else { // Names don't match, so continue to next client

    console.log(clientArray[i].ClientName + " does not match ClientID.");
    console.log("Moving to next client...");
    return false;

  }
  
}

async function sendOutInvoices(page, i){
  
  console.log("SEND OUT INVOICES");
  
  const setup = await precheck(page, i);

  if(setup === clientArray[i].ClientName){ //check names match

    console.log("Names Match.");

    const invoiceArray = await findYear(page, i, year);

    const linkedInvoice = await linkInvoice(page, i, invoiceArray);

    const createdInvoice = await createInvoice(page, i, linkedInvoice);

    if(createdInvoice){
      clientArray[i][invoiceExcel] = "Y";
    } else {
      clientArray[i][invoiceExcel] = "N";
    }

    console.log("INVOICE SEND COMPLETE");
    return true;

  } else { // Names don't match, so continue to next client

    console.log(clientArray[i].ClientName + " does not match ClientID.");
    console.log("Moving to next client...");
    return false;

  }
}

async function closer(page, i){
  console.log("BEGIN CLOSER");
  completeCloser = clientArray[i][closerExcel]
  if(completeCloser == null || completeCloser.toLocaleLowerCase() != 'none'){
    try{ //Ensure Page domain exists
      console.log("Navigate to Jobs Page.");
      await page.goto('https://portal.greenoakfinancial.com/app/jobs/'); // Navigate to Jobs 
    } catch (error) { //catch case if no job page exists
      console.log("Job Page does not exist.");
      console.error(error);
      clientArray[i][closerExcel] = "N";
      return false;
    }
    try {
      console.log("Linking Signature Document...");

      //type into search bar
      await page.waitForSelector('input.search__input');
      await page.type('input.search__input', clientArray[i].ClientName);
      await page.keyboard.press('Enter');
      

      await new Promise(resolve => setTimeout(resolve, 1000)); //1 sec delay


      //click name
      const name = xPathClick(page, `//button[.//span[@class="btn__text" and text()="1040, ${clientArray[i].ClientName}"]]`);
      if(!name){
        clientArray[i][closerExcel] = "N";
        return false;
      }


      //click link button
      const link = xPathClick(page, "//button[.//span[contains(@class, 'btn__text') and text()='Link']]");
      if(!link){
        clientArray[i][closerExcel] = "N";
        return false;
      }


      //click document button
      const documents = xPathClick(page, "//li[contains(@class, 'rc-dropdown-menu-item')]//span[contains(@class, 'btn__text') and contains(text(), 'Documents')]");
      if(!documents){
        clientArray[i][closerExcel] = "N"; 
        return false;
      }


      //navigate to firm docs page
      const firmDocs = xPathClick(page, "//span[contains(text(), 'Firm docs shared with client')]");
      if(!firmDocs){
        clientArray[i][closerExcel] = "N"; 
        return false;
      }


      //navigate to year page
      const yearSelect = xPathClick(page, `//li[contains(@class, 'sub-folder')]//span[contains(text(), '${year}')]`)
      if(!yearSelect){
        clientArray[i][closerExcel] = "N"; 
        return false;
      }


      //click signature document
      const signatureSelect = xPathClick(page, `//span[contains(@class, 'info-block__text') and contains(text(), '${signature}')]`);
      if(!signatureSelect){
        clientArray[i][closerExcel] = "N"; 
        return false;
      }
      await new Promise(resolve => setTimeout(resolve, 3000)); //3 sec delay
      

      //submit signature as linked document
      await page.keyboard.press('Enter');


      //click save and exit button
      await page.waitForSelector('button._root_19cba_1._primary_19cba_59');
      await page.click('button._root_19cba_1._primary_19cba_59');


      console.log("Signature Page Linked.");
      console.log("CLOSER COMPLETE");
      

      clientArray[i][closerExcel] = "Y";
      return true;     
    } catch (error) {
      console.log("Closer Task Failed.");
      console.error(error);
      clientArray[i][closerExcel] = "N";
      return false;
    }
  } else {
    console.log("No closer selected.");
    clientArray[i][closerExcel] = "N";
    return false;
  }
}

async function seal(page, i){

  console.log("Begin Seal for " + clientArray[i].ClientName);
  sealClient = clientArray[i][sealExcel];
  

  if(sealClient == null || sealClient.toLocaleLowerCase() != "none"){
    try{
      // Navigate to Client Documents
      await page.goto('https://portal.greenoakfinancial.com/app/clients/' + clientArray[i].ClientID + '/documents');
            
      // Wait for and click the folder corresponding to the specified year
      await page.waitForSelector(`span[title='${clientUpload}']`);
      await page.click(`span[title='${clientUpload}']`);

      // Delay for 3 seconds (3000 milliseconds)
      await new Promise(resolve => setTimeout(resolve, 3000));
    
      const elementsDetails = await page.evaluate(() => {
        const buttons = Array.from(document.querySelectorAll('input.checkbox__input'));
        return buttons.map((button, index) => {
            const grandparent = button.parentElement ? button.parentElement.parentElement.parentElement : null;
            return {
                index: index,
                grandparentText: grandparent.textContent.trim(),
                grandparentHtml: grandparent ? grandparent.outerHTML : 'No grandparent'
            };
        });
      });

      console.log(elementsDetails);

      //click checkbox of Client Upload
      const clientUploadIndex = elementsDetails.findIndex(element => element.grandparentText.includes(clientArray[i][sealYearExcel] + " Client Upload"));
      console.log(clientUploadIndex);
      if (clientUploadIndex !== -1) {
          await page.evaluate((index) => {
              const buttons = Array.from(document.querySelectorAll('input.checkbox__input'));
              if (buttons[index]) {
                  buttons[index].click();
                  console.log(`Checkbox clicked.`);
              }
          }, clientUploadIndex);
          clientArray[i][sealExcel] = "Y";
      } else {
          console.log(`Checkbox not found`);
          clientArray[i][sealExcel] = "N";
          return false;
      }


      // Seal the folder
      await page.waitForSelector(`button[title='Seal documents to prevent the changes from a client']`);
      await page.click(`button[title='Seal documents to prevent the changes from a client']`);
      
    } catch (error) {
      console.error(error);
      return false;
    }
  }
}

async function findYearCreateArray(page, i, year, excelText) {
  documentShare = clientArray[i][excelText];
  if(documentShare == null || documentShare.toLocaleLowerCase() != 'none'){
    try {
      // Navigate to Client Documents
      await page.goto('https://portal.greenoakfinancial.com/app/clients/' + clientArray[i].ClientID + '/documents');
      
      // Wait for and click the first folder
      await page.waitForSelector('div.-even span.truncate');
      await page.click('div.-even span.truncate');

      // Wait for and click the folder corresponding to the specified year
      await page.waitForSelector('span[title=' + "\"" + year + "\"" + ']');
      await page.click('span[title=' + "\"" + year + "\"" + ']');

      // Delay for 3 seconds (3000 milliseconds)
      await new Promise(resolve => setTimeout(resolve, 3000));

      // Create an array of titles
      const titles = await page.$$eval('span.truncate', spans => spans.map(span => span.getAttribute('title')));

      return titles; // Return the array of titles

    } catch (error) {
      console.log("Navigation to year folder/creation of title array failed.");
      console.error(error);
      return false; // Return false if there was an error
    }
  } else {
    console.log("No " + excelText + " document selected.");
    return false;
  }
}

async function findYearDocuments(page, i, year, documentArray) {
  let tempArray = [];
  for(let j = 0; j < documentArray.length; j++){
    tempArray.push(clientArray[i][documentArray[j].docExcel]);
  }
  
  if(tempArray.every(element => element.toLocaleLowerCase() === 'none')){
    return false;
  }

  try {
    // Navigate to Client Documents
    await page.goto('https://portal.greenoakfinancial.com/app/clients/' + clientArray[i].ClientID + '/documents');
    
    // Wait for and click the first folder
    await page.waitForSelector('div.-even span.truncate');
    await page.click('div.-even span.truncate');

    // Wait for and click the folder corresponding to the specified year
    await page.waitForSelector('span[title=' + "\"" + year + "\"" + ']');
    await page.click('span[title=' + "\"" + year + "\"" + ']');

    // Delay for 3 seconds (3000 milliseconds)
    await new Promise(resolve => setTimeout(resolve, 3000));

    return true;

  } catch (error) {
    console.log("Navigation to year folder failed.");
    console.error(error);
    return false; // Return false if there was an error
  }
}

async function findYear(page, i, year) {
  invoiceExists = clientArray[i][invoiceExcel];
  if(invoiceExists == null || invoiceExists.toLocaleLowerCase() !== 'none'){
    try {
      // Navigate to Client Documents
      await page.goto('https://portal.greenoakfinancial.com/app/clients/' + clientArray[i].ClientID + '/documents');
      
      // Wait for and click the first folder
      await page.waitForSelector('div.-even span.truncate');
      await page.click('div.-even span.truncate');

      // Wait for and click the folder corresponding to the specified year
      await page.waitForSelector('span[title=' + "\"" + year + "\"" + ']');
      await page.click('span[title=' + "\"" + year + "\"" + ']');

      // Delay for 3 seconds (3000 milliseconds)
      await new Promise(resolve => setTimeout(resolve, 3000));

      return true;

    } catch (error) {
      console.log("Navigation to year folder failed.");
      console.error(error);
      return false; // Return false if there was an error
    }
  } else {
    console.log("Invoice not selected.");
    return false;
  }
}

async function shareMultipleDocuments(page, i, navigated, clickList){
  if(navigated){  
      console.log('Locating document checkboxes...');
      try{
        for(let j = 0; j < clickList.length; j++){
          // Retrieve buttons and their details
          const elementsDetails = await page.evaluate(() => {
            const buttons = Array.from(document.querySelectorAll('input.checkbox__input'));
            return buttons.map((button, index) => {
                const grandparent = button.parentElement ? button.parentElement.parentElement.parentElement : null;
                return {
                    index: index,
                    grandparentText: grandparent.textContent.trim(),
                    grandparentHtml: grandparent ? grandparent.outerHTML : 'No grandparent'
                };
            });
          });
          //click vertical ellipsis of signature page
          const signatureButtonIndex = elementsDetails.findIndex(element => element.grandparentText.includes(clickList[j].docName));
          if (signatureButtonIndex !== -1) {
              await page.evaluate((index) => {
                  const buttons = Array.from(document.querySelectorAll('input.checkbox__input'));
                  if (buttons[index]) {
                      buttons[index].click();
                      console.log(`Checkbox clicked.`);
                  }
              }, signatureButtonIndex);
              clientArray[i][clickList[j].docExcel] = "Y";
          } else {
              console.log(`Checkbox not found`);
              continue;
          }
          
        }



        console.log("Checkbox(es) located and selected.");
        console.log("Sharing document(s)...");

        await new Promise(resolve => setTimeout(resolve, 1000)); // Delay for 1 sec

        const share = xPathClick(page, "//button[.//span[contains(@class, 'btn__text') and text()='Share']]");
        if(!share){
          console.log("Share button not found.");
          return false;
        }

        const send = xPathClick(page, "//button[.//span[contains(@class, 'btn__text') and text()='Send']]");
        if(!send){
          console.log("Send button not found.");
          return false;
        }
      

        await new Promise(resolve => setTimeout(resolve, 5000)); // Delay for 3000 milliseconds (3 seconds)
        console.log("Documents Shared.");
        return true;


      } catch (error) {
        console.log("File Send Failed.");
        console.error(error);
        return false;
      }
  } else {
    return false;
  }
}

async function shareDocument(page, i, titles, keyword){
  if(titles){  
    try{
      for (j = 0; j < titles.length; j++){
        titleContains = titles[j].includes(keyword);
        if (titleContains){
          console.log("Title Located: " + titles[j]);
          // await page.waitForSelector('div.checkbox__box:nth-of-type(1)');
          // await page.click('div.checkbox__box:nth-of-type(1)');

          await page.waitForSelector('span[title=' + "\"" + titles[j] +"\"" + ']');
          await page.click('span[title=' + "\""+ titles[j] +"\"" + ']');

          console.log("Title Clicked.");
          break;
        }
      }
      if(titleContains){
        await page.waitForSelector('button.btn.btn_icon[data-test="option-vertical"]');
        await page.click('button.btn.btn_icon[data-test="option-vertical"]');

        await page.waitForSelector('a.btn.btn_menu-item[data-test="share-link"]');
        await page.click('a.btn.btn_menu-item[data-test="share-link"]');

        await page.waitForSelector('button.btn[type="submit"]');
        await page.click('button.btn[type="submit"]');

        await new Promise(resolve => setTimeout(resolve, 5000)); // Delay for 5000 milliseconds (5 seconds)
        console.log(keyword + " File Sent.");        //Estimate File Send Completed
      } else{
        console.log("No " + keyword + " File.")
        return false;
      }
    } catch (error) {
      console.log(keyword + " File Send Failed.");
      console.error(error);
      return false;
    }
    await new Promise(resolve => setTimeout(resolve, 5000)); // Delay for 3000 milliseconds (3 seconds)
    return true;
  } else {
    return false;
  }
}

async function requestSignatures(page, i, titles, keyword, excelText){
  if(titles){ 
    try{ //request signatures from title array with specified keyword
      for (j = 0; j < titles.length; j++){
        titleContains = titles[j].includes(keyword);
        if (titleContains){
          console.log("Title Located: " + titles[j]);
          await page.waitForSelector('span[title=' + "\"" + titles[j] +"\"" + ']');
          await page.click('span[title=' + "\""+ titles[j] +"\"" + ']');
          console.log("Title Clicked.");
          break;
        }
      }
      if(!titleContains){
        console.log(keyword + " does not match any elements.")
        return false;
      }

      await page.waitForSelector('button.btn.btn_icon[data-test="option-vertical"]');
      await page.click('button.btn.btn_icon[data-test="option-vertical"]');

      await page.waitForSelector('a.btn_menu-item[href*="/signature/new?"]');
      await page.click('a.btn_menu-item[href*="/signature/new?"]');

      await new Promise(resolve => setTimeout(resolve, 5000)); //5 sec delay
      
      signatureTemp = clientArray[i].SignatureTemplate
      if(signature == null || signatureTemp.toLocaleLowerCase() != "none"){
        console.log(keyword + " Template: " + clientArray[i].SignatureTemplate);
        await page.waitForSelector('input.react-select__input');
        await page.type('input.react-select__input', clientArray[i].SignatureTemplate);

        await new Promise(resolve => setTimeout(resolve, 3000));

        await page.keyboard.press('Enter');
      } else {
        console.log("No signature template selected");
      }
    
      await new Promise(resolve => setTimeout(resolve, 3000)); //3 sec delay

      const requireKBA = xPathClick(page, "//span[@class='toggle__label']");
      if(!requireKBA){
        console.log("Require KBA button not located.");
        return false;
      }

      // Wait for the element with the specified class and text to be available on the page
      await page.waitForSelector('button.btn.m-t-30.full-width.document-page__button'); // Replace with the actual text if needed
      // Click the element
      await page.click('button.btn.m-t-30.full-width.document-page__button'); // Replace with the actual text if needed


      // Delay for 5000 milliseconds (5 seconds) 
      await new Promise(resolve => setTimeout(resolve, 5000)); //5 sec delay

      //Complete Signature Page Request
      console.log(keyword + " Request Sent.");

      return true;

    } catch (error) { //catch case if signature request fails
      console.log(keyword + " Request Failed.");
      console.error(error);
      return false;
    }
  } else {
    return false;
  }
}

async function linkInvoice(page, i, invoiceArrayCreated){  
  if(invoiceArrayCreated){
    try{
      console.log("Linking Invoice...");
      // Retrieve buttons and their details
      const elementsDetails = await page.evaluate(() => {
        const buttons = Array.from(document.querySelectorAll('button[data-test="option-vertical"]'));
        return buttons.map((button, index) => {
            const grandparent = button.parentElement ? button.parentElement.parentElement : null;
            return {
                index: index,
                grandparentText: grandparent.textContent.trim(),
                grandparentHtml: grandparent ? grandparent.outerHTML : 'No grandparent'
            };
        });
      });

      //click vertical ellipsis of signature page
      const signatureButtonIndex = elementsDetails.findIndex(element => element.grandparentText.includes(signature));
      if (signatureButtonIndex !== -1) {
          await page.evaluate((index) => {
              const buttons = Array.from(document.querySelectorAll('button[data-test="option-vertical"]'));
              if (buttons[index]) {
                  buttons[index].click();
                  console.log(`Vertical Ellipsis of Signature clicked.`);
              }
          }, signatureButtonIndex);
      } else {
          console.log('Button with text "Signature" not found');
          return false;
      }

      //click edit on dropdown
      const edit = xPathClick(page, `//li[@data-menu-id[contains(translate(., 'EDIT', 'edit'), 'edit')]]`);
      if(!edit){
        console.log("Edit Button not located.");
        return false;
      }
      
      //wait for Link Invoice
      await new Promise(resolve => setTimeout(resolve, timeout));
  
      //click link invoice
      const linkUpInvoice = xPathClick(page, "//button[.//span[contains(@class, 'btn__text') and text()='Link invoice']]");
      if(!linkUpInvoice){
        console.log("Link Invoice not located.")
        return false;
      }

      //wait for create invoice
      await new Promise(resolve => setTimeout(resolve, timeout));

      //create invoice
      const makeInvoice = xPathClick(page, "//button[.//span[contains(@class, 'btn__text') and text()='Create invoice']]");
      if(!makeInvoice){
        console.log("Make Inoivce not located.");
        return false;
      }

      console.log("Invoice Link Complete.");
      return true;

    } catch (error) {
      console.error(error);
      return false;
    }
  } else {
    return false;
  }
}

async function createInvoice(page, i, invoiceLinked){
  if(invoiceLinked){
    try{
      console.log("Creating Invoice...");
      console.log("Choosing " + clientArray[i].InvoiceTemplate + " as invoice template...");

      //input invoice tempate
      await page.waitForSelector('input#react-select-5-input.react-select__input');
      await page.type('input#react-select-5-input.react-select__input', clientArray[i].InvoiceTemplate);
      await page.keyboard.press('Enter');

      console.log("Template Chosen.");

      //navigate to invoice amount
      await page.waitForSelector('._textareaField_zwyj7_57');
      await page.click('._textareaField_zwyj7_57');
      await page.keyboard.press('Tab');

      //enter invoice amount
      await page.type('input.simple-input[type="text"][autocapitalize="off"][autocomplete="off"][value="0.00"]', clientArray[i].InvoiceAmount.toString());
      
      //create invoice
      await page.keyboard.press('Enter');
      await new Promise(resolve => setTimeout(resolve, timeout)); //3 sec delay

      //Click save
      const save = xPathClick(page, "//button[@data-test='submit-button' and //span[text()='Save']]");
      if(!save){
        console.log("Save button not located.");
        return false;
      }

      //wait for save
      await new Promise(resolve => setTimeout(resolve, timeout)); //3 sec delay

      console.log('Invoice Created.');
      return true;

    } catch (error) {
      console.error(error);
      return false;
    }
  } else {
    return false;
  }
}

async function xPathClick(page, xpath){
  try{
    //click button
    await page.waitForXPath(xpath);
    const xPathArray = await page.$x(xpath);
    if (xPathArray.length > 0) {
        await xPathArray[0].click();
        return true;
    } else {
        console.log(xpath + " not found.");
        return false;
    }    
  } catch (error) {
    console.log("xPathClick failed.");
    console.error(error);
    return false;
  }
}

async function precheck(page, i){
  try{ //Ensure Page domain exists
    await page.goto('https://portal.greenoakfinancial.com/app/clients/' + clientArray[i].ClientID + '/profile'); // Navigate to Client Profile
  } catch (error) { //catch case if no domain exists
    console.log("Client Profile does not exist.");
    console.error(error);
    return false;
  }

  try{ //Check that page exists
    await page.waitForSelector('.avatar__name', { timeout });
  } catch (error) { // catch case if page doesnt exist
    console.log('https://portal.greenoakfinancial.com/app/clients/' + clientArray[i].ClientID + '/profile' + ' does not exist');
    console.log("Page Check Failed.");
    console.error(error);
    return false;
  }

  try{ //Get the name of page name
    console.log("Checking Names...");
    checkedName = await page.evaluate(() => document.querySelector('.avatar__name').textContent);
    console.log("Check Complete.")
    return checkedName;
  } catch (error) { //catch case if not able to access page name
    console.log("Name Check Failed. ")
    console.error(error);
    return false;
  }
}

function setColorBasedOnValue(cell) {
  if (cell.v === 'X') {
      cell.s = { fill: { patternType: "solid", fgColor: { rgb: "FFFF00" } } }; // Yellow
      console.log("Filling with yellow...");
  } else if (cell.v === 'Y') {
      cell.s = { fill: { patternType: "solid", fgColor: { rgb: "00FF00" } } }; // Green
      console.log("Filling with green...");
  } else if (cell.v === 'N') {
      cell.s = { fill: { patternType: "solid", fgColor: { rgb: "FF0000" } } }; // Red
      console.log("Filling with red...");
  }
}

async function main() {

  const { browser, page } = await startAndLogin();

  try {
    console.log(breakPoint);
    for(let i = 0; i < clientArray.length; i++){ //loop through all clients in excel file
      console.log(clientArray[i].ClientName + ": ")

      const emailCheck = await sendOutEmail(page, i);
      const documentsCheck = await sendOutDocuments(page, i);
      const invoiceCheck = await sendOutInvoices(page, i);
      const closerCheck = await closer(page, i);
      const sealCheck = await seal(page, i);

      // if(emailCheck){
      //   clientArray[i].EmailSent = "Y";
      // } else {
      //   clientArray[i].EmailSent = "N";
      // }


      summarySheet = XLSX.utils.sheet_add_json(summarySheet, clientArray);
      clientBook.Sheets[summaryText] = summarySheet;
      XLSX.writeFile(clientBook, excelPath);

      console.log(breakPoint + " " + (i+1) +" iterations.");
    }
    const range = XLSX.utils.decode_range(summarySheet['!ref']);
    // Loop through each cell in the range
    for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
      for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
        // Construct the cell reference
        const cellRef = XLSX.utils.encode_cell({ r: rowNum, c: colNum });

        // Access the cell
        const cell = summarySheet[cellRef];

        // Perform operations with the cell
        if (cell) {
            setColorBasedOnValue(cell);
        }
      }
    }
    XLSX.writeFile(clientBook, excelPath);
  } catch (error) {
      console.error(error);
  } finally { 
      XLSX.writeFile(clientBook, excelPath);
      await browser.close(); // Close the browser when done
  }
}

main().catch(console.error);
