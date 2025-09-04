const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');

const searchUrl = 'https://www.google.com/maps/search/supermarket/@12.7956,80.0272,8z/data=!4m2!2m1!6e6?entry=ttu&g_ep=EgoyMDI1MDgzMC4wIKXMDSoASAFQAw%3D%3D';

async function saveToExcel(products, filename = 'google_maps_supermarkets.xlsx', append = false) {
  const workbook = new ExcelJS.Workbook();
  
  if (append) {
    try {
      await workbook.xlsx.readFile(filename);
    } catch (error) {
      console.log(`No existing file found or error reading ${filename}. Creating new workbook.`);
    }
  }

  const worksheet = workbook.getWorksheet('Supermarkets') || workbook.addWorksheet('Supermarkets');

  if (!append) {
    worksheet.columns = [
      { header: 'Name', key: 'name', width: 40 },
      { header: 'Rating', key: 'rating', width: 10 },
      { header: 'Reviews', key: 'reviews', width: 10 },
      { header: 'Category', key: 'category', width: 20 },
      { header: 'Address', key: 'address', width: 50 },
      { header: 'Hours', key: 'hours', width: 20 },
      { header: 'Phone', key: 'phone', width: 20 },
      { header: 'Website', key: 'website', width: 50 },
      { header: 'Email', key: 'email', width: 30 },
      { header: 'Services', key: 'services', width: 30 },
      { header: 'Latitude', key: 'latitude', width: 15 },
      { header: 'Longitude', key: 'longitude', width: 15 },
      { header: 'Link', key: 'link', width: 60 }
    ];
  }

  products.forEach(product => {
    worksheet.addRow({
      name: product.name || 'N/A',
      rating: product.rating || 'N/A',
      reviews: product.reviews || 'N/A',
      category: product.category || 'N/A',
      address: product.address || 'N/A',
      hours: product.hours || 'N/A',
      phone: product.phone || 'N/A',
      website: product.website || 'N/A',
      email: product.email || 'N/A',
      services: product.services || 'N/A',
      latitude: product.latitude || 'N/A',
      longitude: product.longitude || 'N/A',
      link: product.link || 'N/A'
    });
  });

  try {
    await workbook.xlsx.writeFile(filename);
    console.log(`Data saved to ${filename} with ${products.length} supermarkets.`);
  } catch (error) {
    console.error(`Error writing to ${filename}: ${error.message}`);
    throw error;
  }
}

async function scrollPage(page) {
  let previousCardCount = 0;
  let scrollCount = 0;
  const maxScrolls = 20; // Adjust based on needs

  console.log('Starting infinite scroll...');
  while (scrollCount < maxScrolls) {
    // Wait for cards to load
    await page.waitForSelector('div.Nv2PK.THOPZb.CpccDe', { timeout: 10000 }).catch(() => console.log('No cards found yet.'));

    // Scroll to bottom of feed
    await page.evaluate(() => {
      const feed = document.querySelector('div[role="feed"]');
      if (feed) feed.scrollTop = feed.scrollHeight;
    });

    // Wait for new content to load
    await new Promise(resolve => setTimeout(resolve, 2000)); // 2-second delay

    // Check number of loaded cards
    const currentCardCount = await page.evaluate(() => document.querySelectorAll('div.Nv2PK.THOPZb.CpccDe').length);
    console.log(`Scroll ${scrollCount + 1}: Found ${currentCardCount} cards.`);

    if (currentCardCount === previousCardCount && currentCardCount > 0) {
      console.log('No new cards loaded. Stopping scroll.');
      break;
    }

    previousCardCount = currentCardCount;
    scrollCount++;
  }
}

async function scrapeGoogleMaps() {
  console.log('Launching browser...');
  const browser = await puppeteer.launch({
    headless: false, // Visible for debugging
    defaultViewport: { width: 1280, height: 800 },
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
    slowMo: 50 // Slight delay for reliability
  });

  const page = await browser.newPage();
  await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36');

  console.log('Navigating to search URL...');
  await retryGoto(page, searchUrl);

  // Handle cookie consent if present
  try {
    await page.waitForSelector('button[aria-label*="Accept"]', { timeout: 5000 });
    await page.click('button[aria-label*="Accept"]');
    await new Promise(resolve => setTimeout(resolve, 1000));
  } catch (error) {
    console.log('No cookie consent dialog found.');
  }

  await scrollPage(page);

  console.log('Extracting supermarket data...');
  const productCards = await page.$$('div.Nv2PK.THOPZb.CpccDe');
  console.log(`Found ${productCards.length} supermarkets.`);

  const supermarkets = [];

  for (let i = 0; i < productCards.length; i++) {
    const card = productCards[i];
    let product = {
      name: 'N/A',
      rating: 'N/A',
      reviews: 'N/A',
      category: 'N/A',
      address: 'N/A',
      hours: 'N/A',
      phone: 'N/A',
      website: 'N/A',
      email: 'N/A',
      services: 'N/A',
      latitude: 'N/A',
      longitude: 'N/A',
      link: 'N/A'
    };

    try {
      // Extract data from search result card
      product = await page.evaluate(card => {
        const nameElem = card.querySelector('.qBF1Pd.fontHeadlineSmall');
        const ratingElem = card.querySelector('.MW4etd');
        const reviewsElem = card.querySelector('.UY7F9');
        const infoElems = card.querySelectorAll('.W4Efsd .W4Efsd');
        const servicesElem = card.querySelector('.qty3Ue');
        const linkElem = card.querySelector('.hfpxzc');
        const imageElem = card.querySelector('.xwpmRb.qisNDe img');

        let category = 'N/A', address = 'N/A', hours = 'N/A', phone = 'N/A', services = 'N/A';
        infoElems.forEach((elem, index) => {
          const text = elem.textContent.trim();
          if (index === 0) {
            const parts = text.split('·').map(s => s.trim());
            category = parts[0] || 'N/A';
            address = parts.find(p => !p.includes('Open') && !p.includes('Closed') && !p.match(/\d{10,}/)) || 'N/A';
          } else if (index === 1) {
            hours = text.split('·')[0].trim() || 'N/A';
            phone = text.match(/\d{10,}/)?.[0] || 'N/A';
          }
        });

        if (servicesElem) {
          services = Array.from(servicesElem.querySelectorAll('.ah5Ghc span')).map(s => s.textContent.trim()).join(', ') || 'N/A';
        }

        const link = linkElem ? linkElem.getAttribute('href') : 'N/A';
        let latitude = 'N/A', longitude = 'N/A';
        if (link !== 'N/A') {
          const latMatch = link.match(/!3d([\d.-]+)/);
          const lonMatch = link.match(/!4d([\d.-]+)/);
          latitude = latMatch ? latMatch[1] : 'N/A';
          longitude = lonMatch ? lonMatch[1] : 'N/A';
        }

        return {
          name: nameElem ? nameElem.textContent.trim() : 'N/A',
          rating: ratingElem ? ratingElem.textContent.trim() : 'N/A',
          reviews: reviewsElem ? reviewsElem.textContent.replace(/[()]/g, '').trim() : 'N/A',
          category,
          address,
          hours,
          phone,
          services,
          latitude,
          longitude,
          link,
          image: imageElem ? imageElem.src : 'N/A'
        };
      }, card);

      // Navigate to detail page for accurate address, phone, hours, website, and email
      if (product.link !== 'N/A') {
        const detailPage = await browser.newPage();
        try {
          await retryGoto(detailPage, product.link);
          await detailPage.waitForSelector('.Io6YTe.fontBodyMedium.kR99db.fdkmkc', { timeout: 10000 }).catch(() => console.log(`Address/phone not found for ${product.name}`));

          // Extract address, phone, website, and email
          const details = await detailPage.evaluate(() => {
            const infoElems = document.querySelectorAll('.Io6YTe.fontBodyMedium.kR99db.fdkmkc');
            const websiteElem = document.querySelector('a[data-item-id="authority"]');
            const emailElem = document.querySelector('a[data-item-id*="email"]');
            return {
              address: infoElems[0] ? infoElems[0].textContent.trim() : 'N/A',
              phone: infoElems[1] ? infoElems[1].textContent.trim() : 'N/A',
              website: websiteElem ? websiteElem.getAttribute('href') : 'N/A',
              email: emailElem ? emailElem.getAttribute('href')?.replace('mailto:', '') : 'N/A'
            };
          });

          product.address = details.address !== 'N/A' ? details.address : product.address;
          product.phone = details.phone !== 'N/A' ? details.phone : product.phone;
          product.website = details.website;
          product.email = details.email;

          // Click to reveal weekly opening hours
          const hoursButton = await detailPage.$('.MkV9 .puWIL.hKrmvd');
          if (hoursButton) {
            await hoursButton.click();
            await new Promise(resolve => setTimeout(resolve, 1000)); // Wait for dropdown
            product.hours = await detailPage.evaluate(() => {
              const rows = document.querySelectorAll('table.eK4R0e tr.y0skZc');
              const hours = [];
              rows.forEach(row => {
                const day = row.querySelector('.ylH6lf')?.textContent.trim();
                const time = row.querySelector('.mxowUb li.G8aQO')?.textContent.trim();
                if (day && time) hours.push(`${day}: ${time}`);
              });
              return hours.join('; ') || 'N/A';
            });
          }
        } catch (error) {
          console.error(`Failed to scrape detail page for ${product.name}: ${error.message}`);
        } finally {
          await detailPage.close();
        }
      }

      console.log(`Extracted data for supermarket ${i + 1}: ${product.name}`);
      supermarkets.push(product);
    } catch (error) {
      console.error(`Error processing card ${i + 1}: ${error.message}`);
    }

    await new Promise(resolve => setTimeout(resolve, 3000)); // Polite delay between requests
  }

  // Save to Excel
  if (supermarkets.length === 0) {
    console.error('No supermarkets found.');
  } else {
    try {
      await saveToExcel(supermarkets);
      console.log(`Total supermarkets extracted: ${supermarkets.length}. Data saved to Excel.`);
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

scrapeGoogleMaps().catch(error => {
  console.error('Error in scraping:', error);
  process.exit(1);
});