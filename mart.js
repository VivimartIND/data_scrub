const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');

const searchUrl = "https://dir.indiamart.com/search.mp?ss=iphone&prdsrc=1&v=4&mcatid=&catid=&crs=xnh-city&trc=xim&cq=Punjaipuliampatti&tags=res:RC2|ktp:N0|mtp:Brn|wc:1|lcf:3|cq:punjaipuliampatti|qr_nm:gl-gd|cs:16544|com-cf:nl|ptrs:na|mc:179822|cat:750|qry_typ:P|lang:en|rtn:2-0-2-0-3-2-1|tyr:3|qrd:250903|mrd:250903|prdt:250903|msf:ls|pfen:1|gli:G1I1";

async function saveToExcel(products, filename = 'indiamart_iphone_products.xlsx', append = false) {
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
      { header: 'Model Number', key: 'modelNumber', width: 20 },
      { header: 'Color', key: 'color', width: 15 },
      { header: 'Internal Memory', key: 'internalMemory', width: 20 },
      { header: 'Battery Capacity', key: 'batteryCapacity', width: 20 },
      { header: 'Country of Origin', key: 'countryOfOrigin', width: 20 },
      { header: 'Supplier Name', key: 'supplierName', width: 30 },
      { header: 'Location', key: 'location', width: 30 },
      { header: 'GST', key: 'gst', width: 20 },
      { header: 'Verified', key: 'verified', width: 15 },
      { header: 'Member Since', key: 'memberSince', width: 15 },
      { header: 'Rating', key: 'rating', width: 15 },
      { header: 'Phone', key: 'phone', width: 20 },
      { header: 'Image', key: 'image', width: 50 },
      { header: 'Link', key: 'link', width: 60 }
    ];
  }

  products.forEach(product => {
    worksheet.addRow({
      name: product.name || 'N/A',
      price: product.price || 'N/A',
      modelNumber: product.modelNumber || 'N/A',
      color: product.color || 'N/A',
      internalMemory: product.internalMemory || 'N/A',
      batteryCapacity: product.batteryCapacity || 'N/A',
      countryOfOrigin: product.countryOfOrigin || 'N/A',
      supplierName: product.supplierName || 'N/A',
      location: product.location || 'N/A',
      gst: product.gst || 'N/A',
      verified: product.verified || 'N/A',
      memberSince: product.memberSince || 'N/A',
      rating: product.rating || 'N/A',
      phone: product.phone || 'N/A',
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

async function scrapeIndiaMart() {
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
  let hasMore = true;
  let loadCount = 0;
  const maxLoads = 10; // Limit to prevent infinite loop; adjust as needed

  // Handle pagination by clicking "Show more results"
  while (hasMore && loadCount < maxLoads) {
    console.log(`Loading more results (batch ${loadCount + 1})...`);

    // Wait for initial cards to load
    await page.waitForSelector('div.card.brs5', { timeout: 10000 }).catch(() => console.log('No cards found yet.'));

    // Click "Show more results" if present
    hasMore = await page.evaluate(() => {
      const moreButton = document.querySelector('div.showmoreresultsdiv button');
      if (moreButton && !moreButton.disabled) {
        moreButton.click();
        return true;
      }
      return false;
    });

    if (hasMore) {
      await new Promise(resolve => setTimeout(resolve, 3000)); // Wait for new content
      loadCount++;
    }
  }

  console.log('Extracting product data...');

  // Extract all cards after loading
  const productCards = await page.$$('div.card.brs5');
  console.log(`Found ${productCards.length} product cards.`);

  for (let i = 0; i < productCards.length; i++) {
    const card = productCards[i];
    let product = {
      name: 'N/A',
      price: 'N/A',
      modelNumber: 'N/A',
      color: 'N/A',
      internalMemory: 'N/A',
      batteryCapacity: 'N/A',
      countryOfOrigin: 'N/A',
      supplierName: 'N/A',
      location: 'N/A',
      gst: 'N/A',
      verified: 'N/A',
      memberSince: 'N/A',
      rating: 'N/A',
      phone: 'N/A',
      image: 'N/A',
      link: 'N/A'
    };

    try {
      // Extract basic card data
      product = await page.evaluate(card => {
        const nameElem = card.querySelector('span.elps.elps2 a.prd-name');
        const priceElem = card.querySelector('p.price');
        const supplierElem = card.querySelector('div.companyname a.cardlinks');
        const locationElem = card.querySelector('div.newLocationUi span.highlight');
        const gstElem = card.querySelector('div.pdinb > span.fs10');
        const verifiedElem = card.querySelector('div.fs10.dsfl.pdinb > span.lh11');
        const memberElem = card.querySelector('div.dsfl.pdinb > span.fs10.mt3');
        const ratingElem = card.querySelector('div.sRt');
        const imageElem = card.querySelector('img.productimg');

        return {
          name: nameElem ? nameElem.textContent.trim() : 'N/A',
          price: priceElem ? priceElem.textContent.trim() : 'N/A',
          supplierName: supplierElem ? supplierElem.textContent.trim() : 'N/A',
          link: supplierElem ? supplierElem.href : 'N/A',
          location: locationElem ? locationElem.textContent.trim() : 'N/A',
          gst: gstElem && gstElem.textContent.includes('GST') ? 'Yes' : 'N/A',
          verified: verifiedElem ? verifiedElem.textContent.trim() : 'N/A',
          memberSince: memberElem ? memberElem.textContent.trim() : 'N/A',
          rating: ratingElem ? ratingElem.textContent.trim() : 'N/A',
          image: imageElem ? imageElem.src : 'N/A'
        };
      }, card);

      // Click "View Mobile Number" to reveal phone
      try {
        const viewMobileButton = await card.$('span.mo.viewn.vmn.fs14.clr5.viewmoboverflow');
        if (viewMobileButton) {
          await viewMobileButton.click();
          await new Promise(resolve => setTimeout(resolve, 1000)); // Wait for number
          product.phone = await page.evaluate(card => {
            const phoneElem = card.querySelector('p.contactnumber span.pns_h.duet.fwb');
            return phoneElem ? phoneElem.textContent.trim() : 'N/A';
          }, card);
        }
      } catch (error) {
        console.error(`Failed to reveal phone for product ${i + 1}: ${error.message}`);
      }

      // Navigate to detail page for specs (Model Number, Color, etc.)
      if (product.link !== 'N/A') {
        const detailPage = await browser.newPage();
        try {
          await retryGoto(detailPage, product.link);
          await detailPage.waitForSelector('.proddetdesc', { timeout: 10000 }).catch(() => {});

          const specs = await detailPage.evaluate(() => {
            const specElements = document.querySelectorAll('.proddetdesc p, .proddetdesc span');
            const specs = {};
            specElements.forEach(elem => {
              const text = elem.textContent.trim();
              if (text.includes(':')) {
                const [key, value] = text.split(':').map(s => s.trim());
                specs[key] = value;
              }
            });
            return specs;
          });

          product.modelNumber = specs['Model Number'] || specs['model number'] || 'N/A';
          product.color = specs['Color'] || 'N/A';
          product.internalMemory = specs['Internal Memory(ROM)'] || specs['Memory Size(ROM)'] || 'N/A';
          product.batteryCapacity = specs['Battery Capacity'] || 'N/A';
          product.countryOfOrigin = specs['Country of Origin'] || 'N/A';
        } catch (error) {
          console.error(`Failed to scrape detail page for product ${i + 1}: ${error.message}`);
        } finally {
          await detailPage.close();
        }
      }

      console.log(`Extracted data for product ${i + 1}: ${product.name}`);
      allProducts.push(product);
    } catch (error) {
      console.error(`Error processing card ${i + 1}: ${error.message}`);
    }

    await new Promise(resolve => setTimeout(resolve, 2000)); // Polite delay
  }

  // Save all products to Excel
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
      console.error(`Attempt ${i + 1} failed for ${url}: ${error.message}`);
      if (i === retries - 1) throw error;
      await new Promise(resolve => setTimeout(resolve, 2000));
    }
  }
}

scrapeIndiaMart().catch(error => {
  console.error('Error in scraping:', error);
  process.exit(1);
});