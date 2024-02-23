const cheerio = require('cheerio');
const rp = require("request-promise");
const ExcelJS = require('exceljs');

const reqlink = 'https://www.flipkart.com/mobile-phones-store';

rp(reqlink)
  .then(function (html) {
    const $ = cheerio.load(html);
    const products = $("a._2rpwqI"); // Selector for product links

    const saveData = [];
    
    products.each((index, element) => {
      const title = $(element).attr('title'); // Extracting the title attribute
      const link = $(element).attr('href'); // Extracting the href attribute
      const image = $(element).find('img._396cs4').attr('src'); // Extracting the src attribute of the image

      saveData.push({ title, link, image });
    });
// console.log(saveData);
    // Create a new Excel workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Mobile Phones');

    // Define the columns in the worksheet
    worksheet.columns = [
        { header: 'Title', key: 'title' },
        { header: 'Link', key: 'link' },
        { header: 'Image', key: 'image' }
    ];

    // Add the scraped data to the worksheet
    saveData.forEach(mobileData => {
        worksheet.addRow(mobileData);
    });

    // Save the workbook to a file
    workbook.xlsx.writeFile('mobileData.xlsx')
      .then(() => {
        console.log('Excel file saved successfully!');
      })
      .catch(err => {
        console.error('Error saving Excel file:', err);
      });
  })
  .catch(function (err) {
    console.log(err);
  });
