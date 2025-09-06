const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');

const searchUrl = 'https://www.exportersindia.com/indian-suppliers/tea.htm';

async function saveToExcel(products, filename = 'exportersindia_tea_products.xlsx', append = false) {
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
      { header: 'Name', key: 'name', width: 40 },
      { header: 'Price', key: 'price', width: 15 },
      { header: 'Supplier Name', key: 'supplierName', width: 30 },
      { header: 'Location', key: 'location', width: 30 },
      { header: 'GST', key: 'gst', width: 20 },
      { header: 'Verified', key: 'verified', width: 15 },
      { header: 'Member Since', key: 'memberSince', width: 15 },
      { header: 'Phone', key: 'phone', width: 20 },
      { header: 'Image', key: 'image', width: 50 },
      { header: 'Link', key: 'link', width: 60 },
      { header: 'Specifications', key: 'specifications', width: 50 },
      { header: 'Nature of Business', key: 'natureOfBusiness', width: 30 }
    ];
  }

  products.forEach(product => {
    worksheet.addRow({
      name: product.name || 'N/A',
      price: product.price || 'N/A',
      supplierName: product.supplierName || 'N/A',
      location: product.location || 'N/A',
      gst: product.gst || 'N/A',
      verified: product.verified || 'N/A',
      memberSince: product.memberSince || 'N/A',
      phone: product.phone || 'N/A',
      image: product.image || 'N/A',
      link: product.link || 'N/A',
      specifications: product.specifications || 'N/A',
      natureOfBusiness: product.natureOfBusiness || 'N/A'
    });
  });

  try {
    await workbook.xlsx.writeFile(filename);
    console.log(`Data saved to ${filename} with ${products.length} products.`);
  } catch (error) {
    console.error(`Error writing to ${filename}: ${error.message}`);
    throw error;
  }
}

async function scrapeExportersIndia() {
  console.log('Launching browser...');
  const browser = await puppeteer.launch({
    headless: false, // Visible for debugging
    defaultViewport: { width: 1280, height: 800 },
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
    slowMo: 50
  });

  const page = await browser.newPage();
  console.log('Navigating to search URL...');
  await retryGoto(page, searchUrl);

  const allProducts = [];
  let lastHeight = 0;
  let sameHeightCount = 0;
  const maxSameHeight = 3; // Stop if page height doesn't change after 3 scrolls
  const maxProducts = 100; // Limit to avoid excessive scraping; adjust as needed

  while (allProducts.length < maxProducts) {
    console.log(`Scraping products (total collected: ${allProducts.length})...`);

    // Wait for product cards to load
    await page.waitForSelector('li div.l3Inn', { timeout: 10000 }).catch(() => console.log('No cards found.'));

    // Extract product cards
    const productCards = await page.$$('li div.l3Inn');
    console.log(`Found ${productCards.length} product cards.`);

    for (let i = 0; i < productCards.length; i++) {
      if (allProducts.length >= maxProducts) break;

      const card = productCards[i];
      let product = {
        name: 'N/A',
        price: 'N/A',
        supplierName: 'N/A',
        location: 'N/A',
        gst: 'N/A',
        verified: 'N/A',
        memberSince: 'N/A',
        phone: 'N/A',
        image: 'N/A',
        link: 'N/A',
        specifications: 'N/A',
        natureOfBusiness: 'N/A'
      };

      try {
        // Extract data from the card
        product = await page.evaluate(card => {
          const nameElem = card.querySelector('h3 a.prdclk');
          const priceElem = card.querySelector('div._price');
          const supplierElem = card.querySelector('div._company a.com_nam');
          const locationElem = card.querySelector('div._address span.title_tooltip');
          const verifiedElem = card.querySelector('ul._mebData li a[title="V-Trust Member"] span');
          const memberElem = card.querySelector('div.ms-yrs span');
          const imageElem = card.querySelector('div.classImg img');
          const linkElem = card.querySelector('h3 a.prdclk');
          const specList = card.querySelector('ul._attriButes');
          const specs = {};
          if (specList) {
            specList.querySelectorAll('li').forEach(li => {
              const key = li.querySelector('span.eipdt-lbl')?.textContent.trim();
              const value = li.querySelector('span.eipdt-val')?.textContent.trim();
              if (key && value) specs[key] = value;
            });
          }

          return {
            name: nameElem ? nameElem.textContent.trim() : 'N/A',
            price: priceElem ? priceElem.textContent.trim().replace(/\s+/g, ' ') : 'N/A',
            supplierName: supplierElem ? supplierElem.textContent.trim() : 'N/A',
            location: locationElem ? locationElem.getAttribute('data-tooltip') : 'N/A',
            verified: verifiedElem ? verifiedElem.textContent.trim() : 'N/A',
            memberSince: memberElem ? memberElem.textContent.trim() : 'N/A',
            image: imageElem ? imageElem.src : 'N/A',
            link: linkElem ? linkElem.href : 'N/A',
            specifications: Object.keys(specs).length > 0 ? JSON.stringify(specs) : 'N/A'
          };
        }, card);

        // Click "View Mobile" to reveal phone
        try {
          const viewMobileButton = await card.$('a._view_mobile');
          if (viewMobileButton) {
            await viewMobileButton.click();
            await new Promise(resolve => setTimeout(resolve, 1000)); // Wait for number to appear
            product.phone = await page.evaluate(card => {
              const phoneButton = card.querySelector('a._view_mobile');
              return phoneButton ? phoneButton.textContent.trim().replace('View Mobile', '').trim() : 'N/A';
            }, card);
          }
        } catch (error) {
          console.error(`Failed to reveal phone for product ${i + 1}: ${error.message}`);
        }

        // Visit product detail page
        if (product.link !== 'N/A') {
          const detailPage = await browser.newPage();
          try {
            await retryGoto(detailPage, product.link);

            // Extract GST and Nature of Business
            const details = await detailPage.evaluate(() => {
              const details = {};
              const listItems = document.querySelectorAll('ul.pdsd-od-list li');
              listItems.forEach(item => {
                const key = item.querySelector('img')?.nextSibling?.textContent.trim();
                const value = item.querySelector('span')?.textContent.trim();
                if (key && value) details[key] = value;
              });
              return details;
            });

            product.gst = details['GST No.'] || 'N/A';
            product.natureOfBusiness = details['Nature of Business'] || 'N/A';

            await detailPage.close();
          } catch (error) {
            console.error(`Error scraping product page ${product.link}: ${error.message}`);
            await detailPage.close();
          }
        }

        console.log(`Extracted data for product ${i + 1}: ${product.name}`);
        allProducts.push(product);
      } catch (error) {
        console.error(`Error processing card ${i + 1}: ${error.message}`);
      }
    }

    // Scroll to load more products
    const currentHeight = await page.evaluate(() => document.body.scrollHeight);
    await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
    await new Promise(resolve => setTimeout(resolve, 2000)); // Wait for new content to load

    // Check if page height changed
    const newHeight = await page.evaluate(() => document.body.scrollHeight);
    if (newHeight === currentHeight) {
      sameHeightCount++;
      if (sameHeightCount >= maxSameHeight) {
        console.log('No more products loaded after scrolling. Stopping.');
        break;
      }
    } else {
      sameHeightCount = 0;
      lastHeight = newHeight;
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

scrapeExportersIndia().catch(error => {
  console.error('Error in scraping:', error);
  process.exit(1);
});