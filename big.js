const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');

const searchUrl = 'https://www.bigbasket.com/cl/fashion/?nc=nb';

async function saveToExcel(products, filename = 'bigbasket_fashion_products.xlsx', append = false) {
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
      { header: 'Brand', key: 'brand', width: 30 },
      { header: 'Product Name', key: 'productName', width: 50 },
      { header: 'Price', key: 'price', width: 15 },
      { header: 'MRP', key: 'mrp', width: 15 },
      { header: 'Offer', key: 'offer', width: 15 },
      { header: 'Description', key: 'description', width: 100 },
      { header: 'Specification', key: 'specification', width: 100 },
      { header: 'Other Info', key: 'otherInfo', width: 100 },
      { header: 'Images', key: 'images', width: 100 },
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

async function scrapeBigBasket() {
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
  let previousCardCount = 0;
  let scrollCount = 0;
  const maxScrolls = 20; // Adjust as needed

  console.log('Starting infinite scroll...');
  while (scrollCount < maxScrolls) {
    await page.waitForSelector('div.SKUDeck___StyledDiv-sc-1e5d9gk-0.eA-dmzP', { timeout: 10000 }).catch(() => console.log('No cards found yet.'));

    // Scroll to bottom
    await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
    await new Promise(resolve => setTimeout(resolve, 3000)); // Wait for load

    const currentCardCount = await page.evaluate(() => document.querySelectorAll('div.SKUDeck___StyledDiv-sc-1e5d9gk-0.eA-dmzP').length);
    console.log(`Scroll ${scrollCount + 1}: Found ${currentCardCount} cards.`);

    if (currentCardCount === previousCardCount && currentCardCount > 0) {
      console.log('No new cards loaded. Stopping scroll.');
      break;
    }

    previousCardCount = currentCardCount;
    scrollCount++;
  }

  console.log('Extracting product data...');
  const productCards = await page.$$('div.SKUDeck___StyledDiv-sc-1e5d9gk-0.eA-dmzP');
  console.log(`Found ${productCards.length} product cards.`);

  for (let i = 0; i < productCards.length; i++) {
    const card = productCards[i];
    let product = {
      brand: 'N/A',
      productName: 'N/A',
      price: 'N/A',
      mrp: 'N/A',
      offer: 'N/A',
      description: 'N/A',
      specification: 'N/A',
      otherInfo: 'N/A',
      images: 'N/A',
      link: 'N/A'
    };

    try {
      // Extract basic info from card
      product = await page.evaluate(card => {
        const brandElem = card.querySelector('span.BrandName___StyledLabel2-sc-hssfrl-1.keQNWn');
        const nameElem = card.querySelector('h3');
        const priceElem = card.querySelector('.Pricing___StyledLabel-sc-pldi2d-1.AypOi');
        const mrpElem = card.querySelector('.Pricing___StyledLabel2-sc-pldi2d-2.hsCgvu');
        const offerElem = card.querySelector('.Tags___StyledLabel-sc-aeruf4-0.kkRHYp');
        const imageElem = card.querySelector('img.DeckImage___StyledImage-sc-1mdvxwk-3.cSWRCd');
        const linkElem = card.querySelector('a');

        return {
          brand: brandElem ? brandElem.textContent.trim() : 'N/A',
          productName: nameElem ? nameElem.textContent.trim() : 'N/A',
          price: priceElem ? priceElem.textContent.trim() : 'N/A',
          mrp: mrpElem ? mrpElem.textContent.trim() : 'N/A',
          offer: offerElem ? offerElem.textContent.trim() : 'N/A',
          image: imageElem ? imageElem.src : 'N/A',
          link: linkElem ? 'https://www.bigbasket.com' + linkElem.getAttribute('href') : 'N/A'
        };
      }, card);

      // Navigate to detail page for description, specs, other info, images
      if (product.link !== 'N/A') {
        const detailPage = await browser.newPage();
        try {
          await retryGoto(detailPage, product.link);
          await detailPage.waitForSelector('section.Image___StyledSection-sc-1nc1erg-0.lhmdrK', { timeout: 10000 }).catch(() => {});

          const details = await detailPage.evaluate(() => {
            const descElem = document.querySelector('.MoreDetails___StyledDiv-sc-1h9rbjh-0.dNIxde .bullets div');
            const specElem = document.querySelectorAll('.MoreDetails___StyledDiv-sc-1h9rbjh-0.dNIxde .bullets ul li');
            const otherElem = document.querySelectorAll('.MoreDetails___StyledDiv-sc-1h9rbjh-0.kIqWEi .bullets p');
            const imageElems = document.querySelectorAll('.thumbnail img');

            let description = descElem ? descElem.textContent.trim() : 'N/A';
            let specification = Array.from(specElem).map(li => li.textContent.trim()).join('; ');
            let otherInfo = Array.from(otherElem).map(p => p.textContent.trim()).join('; ');
            let images = Array.from(imageElems).map(img => img.src).join('; ');

            return {
              description,
              specification: specification || 'N/A',
              otherInfo: otherInfo || 'N/A',
              images: images || 'N/A'
            };
          });

          product.description = details.description;
          product.specification = details.specification;
          product.otherInfo = details.otherInfo;
          product.images = details.images;
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

scrapeBigBasket().catch(error => {
  console.error('Error in scrapeBigBasket:', error);
  process.exit(1);
});