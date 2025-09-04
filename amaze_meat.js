const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');

const baseUrl = "https://www.amazon.in/aw/s?i=freshmeat&bbn=81249930031&rh=n%3A92566315031&s=featured-rank&_encoding=UTF8&pf_rd_p=72d668d4-f892-4690-8a3e-811b6049a0ae&pf_rd_r=EK0G40Y0FCTTZYTQSX52&ref=cct_cg_Meat_2a1";

async function saveToExcel(products, filename = 'amazon_meat_products.xlsx', append = false) {
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
    worksheet.addRow({
      name: product.name || 'N/A',
      price: product.price || 'N/A',
      mrp: product.mrp || 'N/A',
      offer: product.offer || 'N/A',
      image: product.image || 'N/A',
      images: product.images || 'N/A',
      description: product.description || 'N/A',
      seller: product.seller || 'N/A',
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

async function scrapeProductData() {
  console.log('Launching browser...');
  const browser = await puppeteer.launch({
    headless: false, // Visible browser for debugging; set to true for headless
    defaultViewport: { width: 1280, height: 800 },
    args: ['--no-sandbox', '--disable-setuid-sandbox'], // For compatibility in some environments
    slowMo: 0 // No slowMo
  });

  const page = await browser.newPage();
  console.log('Navigating to base URL...');
  await retryGoto(page, baseUrl);

  const allProductLinks = [];
  let currentPage = 1;
  const maxPages = 5; // Limit to 5 pages; adjust as needed
  let hasNext = true;

  while (hasNext && currentPage <= maxPages) {
    console.log(`Collecting links from page ${currentPage}`);

    const links = await page.evaluate(() => {
      const productCards = document.querySelectorAll('div.s-product-image-container');
      return Array.from(productCards).map(card => {
        const linkElement = card.querySelector('a.a-link-normal.s-no-outline');
        return linkElement ? 'https://www.amazon.in' + linkElement.getAttribute('href') : null;
      }).filter(link => link);
    });

    allProductLinks.push(...links);

    // Check for next page
    hasNext = await page.evaluate(() => {
      const nextButton = document.querySelector('a.s-pagination-item.s-pagination-next:not(.s-pagination-disabled)');
      if (nextButton) {
        nextButton.click();
        return true;
      }
      return false;
    });

    if (hasNext) {
      try {
        await page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 30000 });
        currentPage++;
      } catch (error) {
        console.error(`Navigation to next page failed: ${error.message}`);
        hasNext = false;
      }
    }
  }

  console.log(`Collected ${allProductLinks.length} product links.`);

  const allProductDetails = [];
  let isFirstSave = true;
  const batchSize = 5; // Smaller batch size to avoid overwhelming the server

  // Handle Ctrl+C to save data and exit
  process.on('SIGINT', async () => {
    console.log('\nCaught interrupt signal. Saving collected data...');
    try {
      if (allProductDetails.length > 0) {
        await saveToExcel(allProductDetails, 'amazon_meat_products.xlsx', !isFirstSave);
        console.log(`Saved ${allProductDetails.length} products to amazon_meat_products.xlsx`);
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
          const priceElement = document.querySelector('span.a-price[data-a-size="xl"] .a-offscreen');
          const mrpElement = document.querySelector('span.a-price.a-text-price[data-a-strike="true"] .a-offscreen');
          const offerElement = document.querySelector('span.a-badge-label-inner .a-badge-text');
          const imageElement = document.querySelector('img[id="landingImage"]');

          return {
            name: nameElement ? nameElement.textContent.trim() : 'N/A',
            price: priceElement ? priceElement.textContent.trim() : 'N/A',
            mrp: mrpElement ? mrpElement.textContent.trim() : 'N/A',
            offer: offerElement && offerElement.textContent.includes('Save') ? offerElement.textContent.trim() : 'N/A',
            image: imageElement ? imageElement.src : 'N/A',
            link: window.location.href
          };
        });

        // Wait for additional details to load
        await productPage.waitForSelector('#feature-bullets', { timeout: 10000 }).catch(() => {});
        await productPage.waitForSelector('#fresh-merchant-info', { timeout: 10000 }).catch(() => {});
        await productPage.waitForSelector('#altImages', { timeout: 10000 }).catch(() => {});

        const details = await productPage.evaluate(() => {
          // Images
          const imageElements = document.querySelectorAll('#altImages li.a-spacing-small.item.imageThumbnail img');
          const images = Array.from(imageElements)
            .map(img => img.src.replace(/_SX38_SY50_CR,0,0,38,50_/, '_SL1000_'))
            .filter(src => src && !src.includes('360_icon') && !src.includes('transparent-pixel'))
            .join('; ');

          // Description
          const bulletElements = document.querySelectorAll('#feature-bullets ul.a-unordered-list.a-vertical.a-spacing-mini li.a-spacing-mini');
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
      await saveToExcel(batchResults, 'amazon_meat_products.xlsx', !isFirstSave);
      isFirstSave = false;
    } catch (error) {
      console.error(`Failed to save batch: ${error.message}`);
    }
    // Polite delay between batches
    await new Promise(resolve => setTimeout(resolve, 2000));
  }

  if (allProductDetails.length === 0) {
    console.error('No products found.');
  } else {
    // Final save
    try {
      await saveToExcel(allProductDetails, 'amazon_meat_products.xlsx', false);
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
      console.error(`Attempt ${i + 1} failed for ${url}: ${error.message}`);
      if (i === retries - 1) throw error;
      await new Promise(resolve => setTimeout(resolve, 1000)); // Delay before retry
    }
  }
}

scrapeProductData().catch(async (error) => {
  console.error('Error in scrapeProductData:', error);
  process.exit(1);
});