const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');

const baseUrl = "https://www.flipkart.com/clothing-and-accessories/topwear/pr?sid=clo,ash&p[]=facets.ideal_for%255B%255D%3DMen&p[]=facets.ideal_for%255B%255D%3Dmen&otracker=categorytree&fm=neo%2Fmerchandising&iid=M_48bdf622-562e-4fe9-8897-476a97588ff8_1_X1NCR146KC29_MC.RLI1MOY42WPG&otracker=hp_rich_navigation_1_1.navigationCard.RICH_NAVIGATION_Fashion~Men%27s%2BTop%2BWear~All_RLI1MOY42WPG&otracker1=hp_rich_navigation_PINNED_neo%2Fmerchandising_NA_NAV_EXPANDABLE_navigationCard_cc_1_L2_view-all&cid=RLI1MOY42WPG";

async function saveToExcel(products, filename = 'flipkart_products.xlsx', append = false) {
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
      { header: 'Brand', key: 'brand', width: 20 },
      { header: 'Price', key: 'price', width: 15 },
      { header: 'MRP', key: 'mrp', width: 15 },
      { header: 'Offer', key: 'offer', width: 15 },
      { header: 'Image', key: 'image', width: 50 },
      { header: 'Images', key: 'images', width: 60 },
      { header: 'Specifications', key: 'specifications', width: 100 },
      { header: 'Seller', key: 'seller', width: 30 },
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

async function scrapeProductData() {
  console.log('Launching browser...');
  const browser = await puppeteer.launch({
    headless: false,
    defaultViewport: { width: 1280, height: 800 },
    slowMo: 100
  });

  const page = await browser.newPage();
  console.log('Navigating to base URL...');
  await retryGoto(page, baseUrl);

  // Get total pages
  const totalPages = await page.evaluate(() => {
    const paginationText = document.querySelector('div._1G0WLw span')?.textContent || '';
    const match = paginationText.match(/Page \d+ of (\d+)/);
    return match ? parseInt(match[1]) : 1;
  });
  console.log(`Total pages found: ${totalPages}`);

  const allProductDetails = [];
  let isFirstSave = true;
  const maxPages = 5; // Limit to 5 pages for demo; change to totalPages for full scrape

  // Handle Ctrl+C to save data and exit
  process.on('SIGINT', async () => {
    console.log('\nCaught interrupt signal. Saving collected data...');
    try {
      if (allProductDetails.length > 0) {
        await saveToExcel(allProductDetails, 'flipkart_products.xlsx', !isFirstSave);
        console.log(`Saved ${allProductDetails.length} products to flipkart_products.xlsx`);
      } else {
        console.log('No products to save.');
      }
    } catch (error) {
      console.error(`Error saving on interrupt: ${error.message}`);
    }
    await browser.close();
    process.exit(0);
  });

  for (let currentPage = 1; currentPage <= maxPages; currentPage++) {
    const pageUrl = `${baseUrl}&page=${currentPage}`;
    console.log(`Navigating to page ${currentPage}: ${pageUrl}`);
    await retryGoto(page, pageUrl);

    console.log('Extracting products from current page...');
    const products = await page.evaluate(() => {
      const productCards = document.querySelectorAll('div[data-id^="SHT"]');
      return Array.from(productCards).map(card => {
        const brandElement = card.querySelector('div.syl9yP');
        const nameElement = card.querySelector('a.WKTcLC');
        const priceElement = card.querySelector('div.Nx9bqj');
        const mrpElement = card.querySelector('div.yRaY8j');
        const offerElement = card.querySelector('div.UkUFwK');
        const imageElement = card.querySelector('img._53J4C-');
        const linkElement = card.querySelector('a.rPDeLR');

        return {
          brand: brandElement ? brandElement.textContent.trim() : 'N/A',
          name: nameElement ? nameElement.textContent.trim() : 'N/A',
          price: priceElement ? priceElement.textContent.trim() : 'N/A',
          mrp: mrpElement ? mrpElement.textContent.trim() : 'N/A',
          offer: offerElement ? offerElement.textContent.trim() : 'N/A',
          image: imageElement ? imageElement.src : 'N/A',
          link: linkElement ? 'https://www.flipkart.com' + linkElement.getAttribute('href') : 'N/A'
        };
      });
    });

    const validProducts = products.filter(product => product.link !== 'N/A' && product.name !== 'N/A');
    console.log(`Found ${products.length} products on page ${currentPage}, ${validProducts.length} valid after filtering`);

    for (const product of validProducts) {
      let productPage;
      try {
        productPage = await browser.newPage();
        await retryGoto(productPage, product.link);
        await productPage.waitForSelector('div._4WELSP', { timeout: 60000 }); // Wait for main image

        const details = await productPage.evaluate(() => {
          // Images
          const imageElements = document.querySelectorAll('ul.ZqtVYK li img');
          const images = Array.from(imageElements)
            .map(img => img.src.replace(/\/128\/128\//, '/416/416/')) // Get higher res
            .filter(src => src)
            .join('; ');

          // Seller
          const sellerElement = document.querySelector('#sellerName span');
          const seller = sellerElement ? sellerElement.textContent.trim() : 'N/A';

          // Specifications / Product Details
          const specRows = document.querySelectorAll('div.sBVJqn._8vsVX1 .row');
          const specifications = {};
          specRows.forEach(row => {
            const key = row.querySelector('div._9NUIO9')?.textContent.trim();
            const value = row.querySelector('div.-gXFvC')?.textContent.trim();
            if (key && value) {
              specifications[key] = value;
            }
          });
          const specString = Object.entries(specifications)
            .map(([k, v]) => `${k}: ${v}`)
            .join('; ');

          return {
            images: images || 'N/A',
            seller: seller,
            specifications: specString || 'N/A'
          };
        });

        const detailedProduct = { ...product, ...details };
        allProductDetails.push(detailedProduct);

        // Save immediately after each product
        try {
          await saveToExcel([detailedProduct], 'flipkart_products.xlsx', !isFirstSave);
          console.log(`Extracted and saved details for ${product.name}`);
        } catch (error) {
          console.error(`Failed to save product ${product.name}: ${error.message}`);
        }
        isFirstSave = false;
      } catch (error) {
        console.error(`Error scraping details for ${product.name}: ${error.message}`);
        const detailedProduct = { ...product, images: 'N/A', seller: 'N/A', specifications: 'N/A' };
        allProductDetails.push(detailedProduct);
        try {
          await saveToExcel([detailedProduct], 'flipkart_products.xlsx', !isFirstSave);
          console.log(`Saved partial details for ${product.name}`);
        } catch (saveError) {
          console.error(`Failed to save partial product ${product.name}: ${saveError.message}`);
        }
        isFirstSave = false;
      } finally {
        if (productPage) await productPage.close();
      }
    }
  }

  if (allProductDetails.length === 0) {
    console.error('No products found.');
  } else {
    // Final save to ensure all data is written
    try {
      await saveToExcel(allProductDetails, 'flipkart_products.xlsx', false);
      console.log(`Final save: Total products extracted: ${allProductDetails.length}. Data saved to Excel.`);
    } catch (error) {
      console.error(`Error in final save: ${error.message}`);
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
    }
  }
}

scrapeProductData().catch(async (error) => {
  console.error('Error in scrapeProductData:', error);
  process.exit(1);
});