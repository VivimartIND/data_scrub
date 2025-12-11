const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
const ExcelJS = require('exceljs');
puppeteer.use(StealthPlugin());

const CATEGORY_URL = "https://www.zeptonow.com/cn/masala-dry-fruits-more/masala-dry-fruits-more/cid/0c2ccf87-e32c-4438-9560-8d9488fc73e0/scid/8b44cef2-1bab-407e-aadd-29254e6778fa";

async function delay(min, max) {
  const ms = Math.floor(Math.random() * (max - min + 1)) + min;
  console.log(`Waiting ${ms}ms...`);
  await new Promise(r => setTimeout(r, ms));
}

async function safeGoto(page, url, label = "") {
  console.log(`Navigating to: ${url} ${label}`);
  for (let i = 0; i < 3; i++) {
    try {
      const response = await page.goto(url, { waitUntil: 'networkidle0', timeout: 50000 });
      const status = response.status();
      const finalUrl = page.url();

      console.log(`→ Response: ${status} | Final URL: ${finalUrl.substring(0, 80)}${finalUrl.length > 80 ? '...' : ''}`);

      if (status === 503 || finalUrl.includes('captcha') || finalUrl.includes('blocked') || finalUrl.includes('cf-')) {
        console.warn(`Blocked detected (status ${status}) – retrying...`);
        throw new Error('Blocked');
      }
      return true;
    } catch (e) {
      console.warn(`Retry ${i + 1}/3 failed: ${e.message}`);
      if (i === 2) return false;
      await delay(6000, 10000);
    }
  }
  return false;
}

async function scrape() {
  console.log("\nZepto Scraper 2025 – ULTRA DEBUG MODE ACTIVATED\n");

  const browser = await puppeteer.launch({
    headless: false,
    defaultViewport: null,
    userDataDir: './zepto_profile',
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--start-maximized']
  });

  const page = await browser.newPage();

  // Anti-detection
  await page.evaluateOnNewDocument(() => {
    Object.defineProperty(navigator, 'webdriver', { get: () => false });
    window.chrome = { runtime: {}, app: {}, LoadTimes: () => {} };
    Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3] });
  });

  // Step 1: Open homepage
  if (!await safeGoto(page, 'https://www.zeptonow.com', "(Homepage)")) {
    console.error("Cannot reach Zepto homepage → exiting");
    await browser.close();
    return;
  }

  // Step 2: Check if location needed
  const locationInput = await page.$('input[placeholder*="pincode" i], input[placeholder*="Pincode" i]');
  if (locationInput) {
    console.log("Location popup detected! Please select your area/pincode manually...");
    console.log("After you see products → press ENTER in this terminal");
    await new Promise(r => process.stdin.once('data', r));
  } else {
    console.log("Location already saved from previous run");
  }

  // Step 3: Go to category
  if (!await safeGoto(page, CATEGORY_URL, "(Category Page)")) {
    console.error("Failed to load category page");
    await browser.close();
    return;
  }

  // Step 4: Scroll & load all products
  console.log("Scrolling to load all products...");
  for (let i = 0; i < 25; i++) {
    await page.evaluate(() => window.scrollBy(0, 1200));
    await delay(1000, 2000);
  }

  // Step 5: Extract only real product cards
  const products = await page.evaluate(() => {
    const links = Array.from(document.querySelectorAll('a[href^="/pn/"]'));
    return links
      .filter(a => {
        const href = a.getAttribute('href') || '';
        const text = (a.textContent || '').toLowerCase();
        return href.includes('/pvid/') &&
               a.querySelector('img') &&
               !text.includes('load more') &&
               !text.includes('show more') &&
               !text.includes('view all');
      })
      .map(a => {
        const img = a.querySelector('img');
        return {
          name: img?.alt?.trim() || "No Alt Text",
          link: "https://www.zeptonow.com" + a.getAttribute('href'),
          image: img?.src || img?.dataset?.src || "No Image"
        };
      });
  });

  console.log(`Found ${products.length} valid product cards on listing page\n`);

  const results = [];
  const extraKeys = new Set();

  // Step 6: Scrape each product one by one
  for (let i = 0; i < products.length; i++) {
    const p = products[i];
    console.log(`\n[${i + 1}/${products.length}] Processing → ${p.name.substring(0, 60)}...`);

    const tab = await browser.newPage();

    let success = false;
    for (let attempt = 1; attempt <= 2; attempt++) {
      console.log(`   Attempt ${attempt}/2 → Opening product page...`);

      if (await safeGoto(tab, p.link, `(Product ${i + 1})`)) {

        // Wait until price appears or h1 loads
        const priceLoaded = await tab.waitForFunction(() => {
          return document.body.innerText.includes('₹') || document.querySelector('h1');
        }, { timeout: 18000 }).catch(() => false);

        if (!priceLoaded) {
          console.warn(`   Price or title not loaded in 18s – retrying...`);
          await tab.close();
          await delay(8000, 12000);
          continue;
        }

        await delay(4000, 7000); // Let React finish

        const data = await tab.evaluate(() => {
          // Title
          const title = document.querySelector('h1')?.innerText.trim() || "No Title";

          // All ₹ elements
          const rupees = Array.from(document.querySelectorAll('span, p, div'))
            .filter(el => /₹\d/.test(el.innerText))
            .map(el => ({
              text: el.innerText.trim(),
              size: parseFloat(getComputedStyle(el).fontSize) || 0,
              color: getComputedStyle(el).color
            }));

          // Selling price = largest font size + green/not gray
          const sellingPriceEl = rupees
            .filter(r => !r.text.includes('OFF'))
            .sort((a, b) => b.size - a.size)[0];
          const sellingPrice = sellingPriceEl ? sellingPriceEl.text.match(/₹[\d,.,]+/)?.[0] || "N/A" : "N/A";

          // MRP = either "₹xx OFF" or smaller gray text
          const discountText = document.body.innerText.match(/₹\d[\d.,]*\s*OFF/i);
          const mrpFromDiscount = discountText ? discountText[0].split('OFF')[0].trim() : null;

          const mrpEl = rupees.find(r => r.color.includes('rgb(88, 98, 116)') || r.text.includes('OFF'));
          const mrp = mrpFromDiscount || (mrpEl ? mrpEl.text.match(/₹[\d,.,]+/)?.[0] : "N/A");

          const discountEl = Array.from(document.querySelectorAll('span'))
            .find(el => /₹\d.*OFF/i.test(el.innerText));
          const discount = discountEl ? discountEl.innerText.trim() : "N/A";

          const desc = document.querySelector('meta[name="description"]')?.content || "No description";

          const images = Array.from(document.querySelectorAll('img'))
            .map(img => img.src || img.dataset.src || '')
            .filter(src => src && src.includes('product'))
            .slice(0, 12)
            .join('; ');

          const info = {};
          document.querySelectorAll('div, li, p').forEach(el => {
            const txt = el.innerText.trim();
            if (txt.includes(':') && txt.length < 180 && !txt.includes('₹') && txt.split(':').length === 2) {
              const [k, v] = txt.split(':');
              info[k.trim()] = v.trim();
            }
          });

          return { title, sellingPrice, mrp, discount, desc, images, info, debug_rupees: rupees.map(r => r.text) };
        });

        console.log(`   Title: ${data.title.substring(0, 60)}...`);
        console.log(`   Selling Price: ${data.sellingPrice}`);
        console.log(`   MRP Detected: ${data.mrp}`);
        console.log(`   Discount Text: ${data.discount}`);
        console.log(`   All ₹ found: ${data.debug_rupees.join(' | ')}`);

        if (data.title && !data.title.includes("This page isn’t working") && data.sellingPrice !== "N/A") {
          const priceNum = parseFloat(data.sellingPrice.replace(/[^\d.]/g, ''));
          const mrpNum = data.mrp !== "N/A" ? parseFloat(data.mrp.replace(/[^\d.]/g, '')) : priceNum;
          const calculatedOffer = mrpNum > priceNum ? (((mrpNum - priceNum) / mrpNum) * 100).toFixed(1) + "%" : "N/A";

          results.push({
            name: data.title,
            link: p.link,
            image: data.images.split('; ')[0] || p.image,
            price: data.sellingPrice,
            mrp: data.mrp,
            offer: data.discount !== "N/A" ? data.discount : calculatedOffer,
            description: data.desc,
            images: data.images,
            ...data.info
          });

          Object.keys(data.info).forEach(k => extraKeys.add(k));
          console.log(`SUCCESS → ${data.title} | ${data.sellingPrice} | MRP ${data.mrp} | ${data.discount || calculatedOffer}`);
          success = true;
          break;
        } else {
          console.warn(`   Bad data – retrying...`);
        }
      }
      await delay(8000, 12000);
    }

    if (!success) {
      console.log(`FAILED after 2 attempts → ${p.name}`);
      results.push({
        name: p.name + " (Blocked / Timeout)",
        link: p.link,
        price: "N/A", mrp: "N/A", offer: "N/A"
      });
    }

    await tab.close();
    await delay(4000, 8000); // Be respectful
  }

  // Excel export
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Zepto Products');
  const columns = [
    { header: "Name", key: "name", width: 55 },
    { header: "Link", key: "link", width: 85 },
    { header: "Image", key: "image", width: 60 },
    { header: "Price", key: "price", width: 15 },
    { header: "MRP", key: "mrp", width: 15 },
    { header: "Offer", key: "offer", width: 18 },
    { header: "Description", key: "description", width: 80 },
    { header: "Images", key: "images", width: 100 },
  ];
  Array.from(extraKeys).sort().forEach(k => columns.push({ header: k, key: k, width: 40 }));
  ws.columns = columns;
  results.forEach(r => ws.addRow(r));

  const filename = `ZEPTO_MASALA_DEBUG_${new Date().toISOString().slice(0,10)}.xlsx`;
  await wb.xlsx.writeFile(filename);
  console.log(`\nALL DONE! ${results.length} products saved → ${filename}\n`);
  await browser.close();
}

scrape().catch(err => {
  console.error("FATAL ERROR:", err);
});