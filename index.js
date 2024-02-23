const cheerio = require('cheerio');
const rp = require("request-promise");
const ExcelJS = require('exceljs');

const reqlink = 'https://en.wikipedia.org/wiki/List_of_presidents_of_the_United_States';

rp(reqlink)
  .then(function (html) {
    const $ = cheerio.load(html);
    const titles = $('.wikitable tbody tr th');
    const saveData = [];
    
    titles.each((index, element) => {
      const title = $(element).text().trim();
      const name = $(element).find('a').attr('title'); 
      const link = $(element).find('a').attr('href');
      saveData.push({ name, link });
    });

    // Create a new Excel workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Presidents');

    // Define the columns in the worksheet
    worksheet.columns = [
        { header: 'Name', key: 'name' },
        { header: 'Link', key: 'link' }
    ];

    // Add the scraped data to the worksheet
    saveData.forEach(president => {
        worksheet.addRow(president);
    });

    // Save the workbook to a file
    workbook.xlsx.writeFile('presidents.xlsx')
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
