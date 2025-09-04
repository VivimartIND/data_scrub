const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');

const baseUrl = "https://www.amazon.in/alm/category/?_encoding=UTF8&almBrandId=ctnow&node=4859496031&ref_=cct_cg_sfz2s1tn_2a1&pd_rd_w=kdRBw&content-id=amzn1.sym.d8fd3b2a-4f1f-4142-933e-3a2fe109f290&pf_rd_p=d8fd3b2a-4f1f-4142-933e-3a2fe109f290&pf_rd_r=YYPN40GGA76ZJWN2CD7H&pd_rd_wg=MLkrd&pd_rd_r=1d887d87-9622-459c-ac8d-08d183ae1afa";

async function saveToExcel(products, filename = 'amazon_products.xlsx', append = false) {
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
      { header: 'MRP', key: 'mrp', width: 15 },
      { header: 'Offer', key: 'offer', width: 15 },
      { header: 'Image', key: 'image', width: 50 },
      { header: 'Images', key: 'images', width: 60 },
      { header: 'Description', key: 'description', width: 100 },
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
    headless: false, // Changed to false to open the browser visibly
    defaultViewport: { width: 1280, height: 800 },
    slowMo: 0 // No slowMo
  });

  const page = await browser.newPage();
  console.log('Navigating to base URL...');
  await retryGoto(page, baseUrl);

  console.log('Scrolling to load products...');
  let previousHeight = 0;
  let scrollAttempts = 0;
  const maxScrollAttempts = 20; // Adjust as needed to load sufficient products
  while (scrollAttempts < maxScrollAttempts) {
    await page.evaluate(() => window.scrollBy(0, window.innerHeight));
    await new Promise(resolve => setTimeout(resolve, 1000)); // Reduced delay
    const newHeight = await page.evaluate(() => document.body.scrollHeight);
    if (newHeight === previousHeight) break;
    previousHeight = newHeight;
    scrollAttempts++;
  }

  console.log('Extracting product links...');
  const allProductLinks = await page.evaluate(() => {
    const productCards = document.querySelectorAll('div[data-csa-c-type="item"]');
    return Array.from(productCards).map(card => {
      const linkElement = card.querySelector('a.a-link-normal');
      return linkElement ? 'https://www.amazon.in' + linkElement.getAttribute('href') : null;
    }).filter(link => link);
  });

  console.log(`Collected ${allProductLinks.length} product links.`);

  const allProductDetails = [];
  let isFirstSave = true;
  const batchSize = 10; // Increased batch size for more concurrency

  // Handle Ctrl+C to save data and exit
  process.on('SIGINT', async () => {
    console.log('\nCaught interrupt signal. Saving collected data...');
    try {
      if (allProductDetails.length > 0) {
        await saveToExcel(allProductDetails, 'amazon_products.xlsx', !isFirstSave);
        console.log(`Saved ${allProductDetails.length} products to amazon_products.xlsx`);
      } else {
        console.log('No products to save.');
      }
    } catch (error) {
      console.error(`Error saving on interrupt: ${error.message}`);
    }
    await browser.close();
    process.exit(0);
  });

  for (let i = 0; i < allProductLinks.length; i += batchSize) {
    const batchLinks = allProductLinks.slice(i, i + batchSize);
    const batchPromises = batchLinks.map(async (link) => {
      let productPage;
      try {
        productPage = await browser.newPage();
        await retryGoto(productPage, link);

        const product = await productPage.evaluate(() => {
          const nameElement = document.querySelector('span[id="productTitle"]');
          const priceElement = document.querySelector('span.a-price-whole');
          const mrpElement = document.querySelector('span.a-text-price');
          const offerElement = document.querySelector('span.savingsPercentage');
          const imageElement = document.querySelector('#landingImage');

          return {
            name: nameElement ? nameElement.textContent.trim() : 'N/A',
            price: priceElement ? '₹' + priceElement.textContent.trim() : 'N/A',
            mrp: mrpElement ? mrpElement.textContent.trim() : 'N/A',
            offer: offerElement ? offerElement.textContent.trim() : 'N/A',
            image: imageElement ? imageElement.src : 'N/A',
            link: window.location.href
          };
        });

        // Wait for images, description, seller
        await productPage.waitForSelector('#feature-bullets', { timeout: 10000 }).catch(() => {}); // Optional wait

        const details = await productPage.evaluate(() => {
          // Images
          const imageElements = document.querySelectorAll('#altImages img');
          const images = Array.from(imageElements)
            .map(img => img.src.replace(/_SX38_SY50_/, '_SX679_')) // Higher res
            .filter(src => src && !src.includes('transparent-pixel'))
            .join('; ');

          // Description
          const bulletElements = document.querySelectorAll('#feature-bullets ul li');
          const description = Array.from(bulletElements)
            .map(li => li.textContent.trim())
            .filter(text => text)
            .join('; ');

          // Seller
          const sellerElement = document.querySelector('#fresh-merchant-info a span');
          const seller = sellerElement ? sellerElement.textContent.trim() : 'N/A';

          return {
            images: images || 'N/A',
            description: description || 'N/A',
            seller: seller
          };
        });

        const detailedProduct = { ...product, ...details };
        console.log(`Extracted details for ${product.name}`);
        return detailedProduct;
      } catch (error) {
        console.error(`Error scraping details for link ${link}: ${error.message}`);
        return {
          name: 'N/A',
          price: 'N/A',
          mrp: 'N/A',
          offer: 'N/A',
          image: 'N/A',
          link: link,
          images: 'N/A',
          description: 'N/A',
          seller: 'N/A'
        };
      } finally {
        if (productPage) await productPage.close();
      }
    });

    const batchResults = await Promise.all(batchPromises);
    allProductDetails.push(...batchResults);

    // Save batch immediately
    try {
      await saveToExcel(batchResults, 'amazon_products.xlsx', !isFirstSave);
      isFirstSave = false;
    } catch (error) {
      console.error(`Failed to save batch: ${error.message}`);
    }
  }

  if (allProductDetails.length === 0) {
    console.error('No products found.');
  } else {
    // Final save
    try {
      await saveToExcel(allProductDetails, 'amazon_products.xlsx', false);
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