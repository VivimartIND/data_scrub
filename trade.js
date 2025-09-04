const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');

const searchUrl = 'https://www.tradeindia.com/search.html?keyword=supermarket%20in%20Chennai';

async function saveToExcel(products, filename = 'tradeindia_supermarkets.xlsx', append = false) {
  const workbook = new ExcelJS.Workbook();
  
  if (append) {
    try {
      await workbook.xlsx.readFile(filename);
    } catch (error) {
      console.log(`No existing file found or error reading ${filename}. Creating new workbook.`);
    }
  }

  const worksheet = workbook.getWorksheet('Products') || workbook.addWorksheet('Products');

  if (!append) {
    worksheet.columns = [
      { header: 'Product Name', key: 'productName', width: 40 },
      { header: 'Price', key: 'price', width: 15 },
      { header: 'MOQ', key: 'moq', width: 15 },
      { header: 'Seller Name', key: 'sellerName', width: 30 },
      { header: 'Location', key: 'location', width: 20 },
      { header: 'Business Type', key: 'businessType', width: 30 },
      { header: 'Member Since', key: 'memberSince', width: 15 },
      { header: 'GST', key: 'gst', width: 20 },
      { header: 'Specifications', key: 'specifications', width: 50 },
      { header: 'Product Overview', key: 'productOverview', width: 50 },
      { header: 'Company Details', key: 'companyDetails', width: 50 },
      { header: 'Address', key: 'address', width: 50 },
      { header: 'Contact Person', key: 'contactPerson', width: 30 },
      { header: 'Image', key: 'image', width: 50 },
      { header: 'Link', key: 'link', width: 60 }
    ];
  }

  products.forEach(product => {
    worksheet.addRow(product);
  });

  try {
    await workbook.xlsx.writeFile(filename);
    console.log(`Data saved to ${filename} with ${products.length} products.`);
  } catch (error) {
    console.error(`Error writing to ${filename}: ${error.message}`);
    throw error;
  }
}

async function scrapeTradeIndia() {
  console.log('Launching browser...');
  const browser = await puppeteer.launch({
    headless: false, // Visible for debugging
    defaultViewport: { width: 1280, height: 800 },
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
    slowMo: 50 // Slight slowMo for reliability
  });

  const page = await browser.newPage();
  console.log('Navigating to search URL...');
  await retryGoto(page, searchUrl);

  const allProducts = [];
  let currentPage = 1;
  let hasNext = true;

  while (hasNext) {
    console.log(`Scraping page ${currentPage}...`);

    // Extract product cards on current page
    const productCards = await page.$$('div.col-md-3.mb-3');
    console.log(`Found ${productCards.length} product cards on page ${currentPage}.`);

    for (let i = 0; i < productCards.length; i++) {
      const card = productCards[i];
      let product = {
        productName: 'N/A',
        price: 'N/A',
        moq: 'N/A',
        sellerName: 'N/A',
        location: 'N/A',
        businessType: 'N/A',
        memberSince: 'N/A',
        gst: 'N/A',
        specifications: 'N/A',
        productOverview: 'N/A',
        companyDetails: 'N/A',
        address: 'N/A',
        contactPerson: 'N/A',
        image: 'N/A',
        link: 'N/A'
      };

      try {
        // Extract basic info from card
        product = await page.evaluate(card => {
          const nameElem = card.querySelector('h2.Body3R');
          const priceElem = card.querySelector('p.price');
          const moqElem = card.querySelector('p.moq');
          const sellerElem = card.querySelector('h3.Body4R.coy-name');
          const locationElem = card.querySelector('p.location');
          const yearsElem = card.querySelector('p.years');
          const imageElem = card.querySelector('img');
          const linkElem = card.querySelector('a');

          return {
            productName: nameElem ? nameElem.textContent.trim() : 'N/A',
            price: priceElem ? priceElem.textContent.trim() : 'N/A',
            moq: moqElem ? moqElem.textContent.trim() : 'N/A',
            sellerName: sellerElem ? sellerElem.textContent.trim() : 'N/A',
            location: locationElem ? locationElem.textContent.trim() : 'N/A',
            memberSince: yearsElem ? yearsElem.textContent.trim() : 'N/A',
            image: imageElem ? imageElem.src : 'N/A',
            link: linkElem ? linkElem.href : 'N/A'
          };
        }, card);

        // Navigate to detail page for more info
        if (product.link !== 'N/A') {
          const detailPage = await browser.newPage();
          try {
            await retryGoto(detailPage, product.link);
            await detailPage.waitForSelector('.product-detail-section', { timeout: 10000 }).catch(() => {});

            const details = await detailPage.evaluate(() => {
              const businessTypeElem = document.querySelector('.business-details p:nth-of-type(1)');
              const gstElem = document.querySelector('.gst p');
              const specsTable = document.querySelector('table.spec-table');
              const overviewElem = document.querySelector('.about-product div.seo-content');
              const companyDetailsElem = document.querySelector('.company-details div.seo-content');
              const addressElem = document.querySelector('.info-block p.title');
              const contactPersonElem = document.querySelector('.info-block p.title');

              let specifications = 'N/A';
              if (specsTable) {
                specifications = Array.from(specsTable.querySelectorAll('tr')).map(row => {
                  const key = row.querySelector('td:first-child')?.textContent.trim();
                  const value = row.querySelector('td:last-child')?.textContent.trim();
                  return `${key}: ${value}`;
                }).join('; ');
              }

              return {
                businessType: businessTypeElem ? businessTypeElem.textContent.trim() : 'N/A',
                gst: gstElem ? gstElem.textContent.trim() : 'N/A',
                specifications,
                productOverview: overviewElem ? overviewElem.textContent.trim() : 'N/A',
                companyDetails: companyDetailsElem ? companyDetailsElem.textContent.trim() : 'N/A',
                address: addressElem ? addressElem.textContent.trim() : 'N/A',
                contactPerson: contactPersonElem ? contactPersonElem.textContent.trim() : 'N/A'
              };
            });

            product.businessType = details.businessType;
            product.gst = details.gst;
            product.specifications = details.specifications;
            product.productOverview = details.productOverview;
            product.companyDetails = details.companyDetails;
            product.address = details.address;
            product.contactPerson = details.contactPerson;
          } catch (error) {
            console.error(`Failed to scrape detail page for product ${i + 1}: ${error.message}`);
          } finally {
            await detailPage.close();
          }
        }

        console.log(`Extracted data for product ${i + 1}: ${product.productName}`);
        allProducts.push(product);
      } catch (error) {
        console.error(`Error processing card ${i + 1}: ${error.message}`);
      }

      await new Promise(resolve => setTimeout(resolve, 2000)); // Polite delay
    }

    // Check for next page
    hasNext = await page.evaluate(() => {
      const nextButton = document.querySelector('li.last-link a.highlight_btn');
      if (nextButton) {
        nextButton.click();
        return true;
      }
      return false;
    });

    if (hasNext) {
      await page.waitForNavigation({ waitUntil: 'networkidle2' }).catch(() => {});
      currentPage++;
    }
  }

  // Save to Excel
  if (allProducts.length === 0) {
    console.error('No products found.');
  } else {
    try {
      await saveToExcel(allProducts);
      console.log(`Total products extracted: ${allProducts.length}. Data saved to Excel.`);
    } catch (error) {
      console.error(`Error saving data: ${error.message}`);
    }
  }

  await browser.close();
}

async function retryGoto(page, url, retries = 3) {
  for (let i = 0; i < retries; i++) {
    try {
      await page.goto(url, { waitUntil: 'networkidle2', timeout: 30000 });
      return;
    } catch (error) {
      console.error(`Attempt ${i + 1} failed: ${error.message}`);
      if (i === retries - 1) throw error;
      await new Promise(resolve => setTimeout(resolve, 2000));
    }
  }
}

scrapeTradeIndia().catch(error => {
  console.error('Error in scrapeTradeIndia:', error);
  process.exit(1);
});