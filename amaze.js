const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');

const searchUrl = 'https://www.amazon.in/s?k=Kids+furnishings&pf_rd_i=1380442031&pf_rd_m=A1VBAL9TL5WCBF&pf_rd_p=88bdd31e-821b-49b6-a8f9-73ff228c5098&pf_rd_r=1WMX76R5MK2S30BC6C3G&pf_rd_s=merchandised-search-7&ref=QAHzEditorial_en_IN_1';

// Custom delay function for compatibility
const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

async function saveToExcel(products, filename = 'amazonindia_kids_furnishings.xlsx', append = false) {
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
      { header: 'Name', key: 'name', width: 50 },
      { header: 'Price', key: 'price', width: 15 },
      { header: 'MRP', key: 'mrp', width: 15 },
      { header: 'Offer', key: 'offer', width: 15 },
      { header: 'Seller', key: 'seller', width: 30 },
      { header: 'Image', key: 'image', width: 50 },
      { header: 'Link', key: 'link', width: 60 }
    ];
  }

  products.forEach(product => {
    worksheet.addRow({
      name: product.name || 'N/A',
      price: product.price || 'N/A',
      mrp: product.mrp || 'N/A',
      offer: product.offer || 'N/A',
      seller: product.seller || 'N/A',
      image: product.image || 'N/A',
      link: product.link || 'N/A'
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

async function scrapeAmazonIndia() {
  console.log('Launching browser...');
  const browser = await puppeteer.launch({
    headless: false, // Visible for debugging
    defaultViewport: { width: 1280, height: 800 },
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
    slowMo: 100 // For dynamic content
  });

  const page = await browser.newPage();
  await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36');
  console.log('Navigating to search URL...');
  await retryGoto(page, searchUrl);

  const allProducts = [];
  const maxPages = 7; // Based on pagination data
  let currentPage = 1;

  while (currentPage <= maxPages) {
    console.log(`Scraping page ${currentPage} (total products collected: ${allProducts.length})...`);

    // Wait for product cards
    await page.waitForSelector('div.s-product-image-container', { timeout: 15000 }).catch(() => console.log('No cards found or timeout.'));

    // Extract product cards
    const productCards = await page.$$('div.s-product-image-container');
    console.log(`Found ${productCards.length} product cards on page ${currentPage}.`);

    for (let i = 0; i < productCards.length; i++) {
      let product = {
        name: 'N/A',
        price: 'N/A',
        mrp: 'N/A',
        offer: 'N/A',
        seller: 'N/A',
        image: 'N/A',
        link: 'N/A'
      };

      try {
        // Wait for dynamic content
        await page.waitForSelector('div.puis-card-container h2 a span', { timeout: 5000 }).catch(() => console.log(`Name element not found for card ${i + 1}.`));
        await delay(1000); // Increased delay

        // Extract data from the card
        product = await page.evaluate((card, index) => {
          const cardContainer = card.closest('div.s-card-container') || card.closest('div.puis-card-container');
          if (!cardContainer) {
            console.log(`Card ${index + 1}: No parent container found`);
            return {
              name: 'N/A',
              price: 'N/A',
              mrp: 'N/A',
              offer: 'N/A',
              image: 'N/A',
              link: 'N/A'
            };
          }

          // Updated selectors
          const nameElem = cardContainer.querySelector('h2 a span.a-text-normal') || cardContainer.querySelector('h2 span');
          const priceElem = cardContainer.querySelector('span.a-price > span.a-offscreen') || cardContainer.querySelector('span.a-price-whole');
          const mrpElem = cardContainer.querySelector('span.a-price.a-text-price span.a-offscreen');
          const offerElem = cardContainer.querySelector('span.savingPriceOverride') || cardContainer.querySelector('div.a-row.a-size-base.a-color-base span');
          const imageElem = card.querySelector('img.s-image');
          const linkElem = card.querySelector('a.a-link-normal');

          // Debugging logs
          console.log(`Card ${index + 1} - Name element: ${nameElem ? nameElem.textContent.trim() : 'Not found'}`);
          console.log(`Card ${index + 1} - Price element: ${priceElem ? priceElem.textContent.trim() : 'Not found'}`);
          console.log(`Card ${index + 1} - MRP element: ${mrpElem ? mrpElem.textContent.trim() : 'Not found'}`);
          console.log(`Card ${index + 1} - Offer element: ${offerElem ? offerElem.textContent.trim() : 'Not found'}`);
          console.log(`Card ${index + 1} - Image element: ${imageElem ? imageElem.src : 'Not found'}`);
          console.log(`Card ${index + 1} - Link element: ${linkElem ? linkElem.href : 'Not found'}`);

          return {
            name: nameElem ? nameElem.textContent.trim() : 'N/A',
            price: priceElem ? priceElem.textContent.trim() : 'N/A',
            mrp: mrpElem ? mrpElem.textContent.trim() : 'N/A',
            offer: offerElem ? offerElem.textContent.trim() : 'N/A',
            image: imageElem ? imageElem.src : 'N/A',
            link: linkElem ? linkElem.href : 'N/A'
          };
        }, productCards[i], i);

        // Visit product detail page for seller
        if (product.link !== 'N/A') {
          const detailPage = await browser.newPage();
          try {
            console.log(`Navigating to product page: ${product.link}`);
            await retryGoto(detailPage, product.link);

            await detailPage.waitForSelector('a#sellerProfileTriggerId', { timeout: 10000 }).catch(() => console.log(`Seller element not found for ${product.link}`));

            product.seller = await detailPage.evaluate(() => {
              const sellerElem = document.querySelector('a#sellerProfileTriggerId');
              return sellerElem ? sellerElem.textContent.trim() : 'N/A';
            });

            console.log(`Extracted seller for product ${i + 1}: ${product.seller}`);
            await detailPage.close();
          } catch (error) {
            console.error(`Error scraping product page ${product.link}: ${error.message}`);
            await detailPage.close();
          }
        }

        console.log(`Extracted data for product ${i + 1}: ${product.name}`);
        allProducts.push(product);
      } catch (error) {
        console.error(`Error processing card ${i + 1} on page ${currentPage}: ${error.message}`);
      }
    }

    // Navigate to next page
    if (currentPage < maxPages) {
      try {
        const nextButton = await page.$('a.s-pagination-next');
        if (nextButton) {
          console.log(`Navigating to page ${currentPage + 1}`);
          await Promise.all([
            page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 30000 }),
            nextButton.click()
          ]);
          currentPage++;
          await delay(3000);
        } else {
          console.log('No next page button found. Stopping pagination.');
          break;
        }
      } catch (error) {
        console.error(`Error navigating to page ${currentPage + 1}: ${error.message}`);
        break;
      }
    } else {
      console.log('Reached maximum page limit.');
      break;
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
      await delay(2000);
    }
  }
}

scrapeAmazonIndia().catch(error => {
  console.error('Error in scraping:', error);
  process.exit(1);
});