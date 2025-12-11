const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
const ExcelJS = require('exceljs');
puppeteer.use(StealthPlugin());

// CHANGE THIS TO ANY ZEPTO CATEGORY URL
const CATEGORY_URL = "https://www.zeptonow.com/cn/masala-dry-fruits-more/masala-dry-fruits-more/cid/0c2ccf87-e32c-4438-9560-8d9488fc73e0/scid/8b44cef2-1bab-407e-aadd-29254e6778fa";

async function delay(min, max) {
  const ms = Math.floor(Math.random() * (max - min + 1)) + min;
  console.log(`Waiting ${ms}ms...`);
  await new Promise(r => setTimeout(r, ms));
}

async function safeGoto(page, url, label = "") {
  console.log(`Navigating → ${url} ${label}`);
  for (let i = 0; i < 3; i++) {
    try {
      const resp = await page.goto(url, { waitUntil: 'networkidle0', timeout: 60000 });
      const status = resp.status();
      const currentUrl = page.url();

      console.log(`→ Status: ${status} | URL: ${currentUrl.substring(0, 90)}${currentUrl.length > 90 ? '...' : ''}`);

      if (status >= 400 || currentUrl.includes('captcha') || currentUrl.includes('blocked') || currentUrl.includes('cf-')) {
        throw new Error('Blocked by Cloudflare');
      }
      return true;
    } catch (e) {
      console.warn(`Retry ${i + 1}/3 failed: ${e.message}`);
      if (i === 2) return false;
      await delay(8000, 12000);
    }
  }
  return false;
}

async function scrapeAllProducts() {
  console.log("\nZEPTO FULL CATEGORY SCRAPER 2025 – GETS ALL 400–500+ PRODUCTS\n");

  const browser = await puppeteer.launch({
    headless: false,                    // Must be visible for Zepto to trust you
    defaultViewport: null,
    userDataDir: './zepto_profile',     // Saves your location forever
    args: ['--start-maximized', '--no-sandbox', '--disable-setuid-sandbox']
  });

  const page = await browser.newPage();

  // Bypass bot detection
  await page.evaluateOnNewDocument(() => {
    Object.defineProperty(navigator, 'webdriver', { get: () => false });
    window.chrome = { runtime: {}, app: {}, LoadTimes: () => {} };
    Object.defineProperty(navigator, 'plugins', { get: () => [1,2,3,4,5] });
    Object.defineProperty(navigator, 'languages', { get: () => ['en-US', 'en'] });
  });

  // 1. Open homepage
  if (!await safeGoto(page, 'https://www.zeptonow.com', "(Homepage)")) {
    console.error("Cannot reach Zepto – check internet or try again later");
    await browser.close();
    return;
  }

  // 2. Handle location popup (only once)
  const pincodeInput = await page.$('input[placeholder*="pincode" i]');
  if (pincodeInput) {
    console.log("\nLocation not set! Please manually:");
    console.log("   → Enter your pincode");
    console.log("   → Select your area");
    console.log("   → Wait until products load");
    console.log("   → Then press ENTER here in terminal\n");
    await new Promise(r => process.stdin.once('data', r));
  } else {
    console.log("Location already saved – proceeding...\n");
  }

  // 3. Go to category
  if (!await safeGoto(page, CATEGORY_URL, "(Category Page)")) {
    console.error("Failed to load category page");
    await browser.close();
    return;
  }

  // 4. FORCE LOAD ALL PRODUCTS (448+)
  console.log("FORCING ZEPTO TO LOAD ALL PRODUCTS – THIS TAKES 3–6 MINUTES...\n");

  let lastCount = 0;
  let stableRounds = 0;

  for (let round = 1; round <= 80; round++) {
    // Scroll to bottom
    await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
    await delay(3000, 5000);

    // Click any "Load More" / "Show More" button
    const clicked = await page.evaluate(() => {
      const buttons = Array.from(document.querySelectorAll('button, div[role="button"], span[role="button"]'));
      const btn = buttons.find(el => {
        const text = (el.textContent || '').toLowerCase();
        return /load|show|view|more|products/i.test(text);
      });
      if (btn) {
        btn.scrollIntoView({ block: 'center' });
        btn.click();
        return true;
      }
      return false;
    });

    // Count current products
    const count = await page.evaluate(() => document.querySelectorAll('a[href^="/pn/"]').length);

    console.log(`Round ${round.toString().padStart(2)} | ${count.toString().padStart(4)} products ${clicked ? '→ CLICKED Load More' : ''}`);

    if (count === lastCount) {
      stableRounds++;
      if (stableRounds >= 12) {
        console.log(`\nNo new products for 12 rounds → ALL ${count} PRODUCTS LOADED!\n`);
        break;
      }
    } else {
      stableRounds = 0;
    }
    lastCount = count;
    await delay(2000, 4000);
  }

  // 5. Extract ALL product links
  const products = await page.evaluate(() => {
    return Array.from(document.querySelectorAll('a[href^="/pn/"]'))
      .filter(a => {
        const href = a.getAttribute('href') || '';
        return href.includes('/pvid/') && a.querySelector('img');
      })
      .map(a => {
        const img = a.querySelector('img');
        return {
          name: img?.alt?.trim() || "Unknown Product",
          link: "https://www.zeptonow.com" + a.getAttribute('href'),
          image: img?.src || img?.dataset?.src || "N/A"
        };
      });
  });

  console.log(`\nEXTRACTED ${products.length} PRODUCTS – STARTING DETAILED SCRAPING...\n`);

  const results = [];
  const extraKeys = new Set();

  // 6. Scrape each product page
  for (let i = 0; i < products.length; i++) {
    const p = products[i];
    console.log(`\n[${(i+1).toString().padStart(3)}/${products.length}] ${p.name.substring(0, 65)}...`);

    const tab = await browser.newPage();
    let success = false;

    for (let attempt = 1; attempt <= 2; attempt++) {
      if (await safeGoto(tab, p.link)) {
        await tab.waitForFunction(() => document.body.innerText.includes('₹') || document.querySelector('h1'), { timeout: 20000 }).catch(() => {});
        await delay(4000, 7000);

        const data = await tab.evaluate(() => {
          const title = document.querySelector('h1')?.innerText.trim() || "No Title";

          const rupees = Array.from(document.querySelectorAll('span, p, div'))
            .filter(el => /₹\d/.test(el.innerText))
            .map(el => ({
              text: el.innerText.trim(),
              size: parseFloat(getComputedStyle(el).fontSize) || 0,
              color: getComputedStyle(el).color
            }));

          const sellingPrice = rupees
            .filter(r => !r.text.includes('OFF'))
            .sort((a, b) => b.size - a.size)[0];
          const price = sellingPrice ? sellingPrice.text.match(/₹[\d.,]+/)?.[0] || "N/A" : "N/A";

          const discountMatch = document.body.innerText.match(/₹\d[\d.,]*\s*OFF/i);
          const mrpFromDiscount = discountMatch ? discountMatch[0].split('OFF')[0].trim() : null;
          const mrp = mrpFromDiscount || rupees.find(r => r.color === 'rgb(88, 98, 116)')?.text.match(/₹[\d.,]+/)?.[0] || "N/A";

          const discount = Array.from(document.querySelectorAll('span'))
            .find(el => /₹.*OFF/i.test(el.innerText))?.innerText.trim() || "N/A";

          const desc = document.querySelector('meta[name="description"]')?.content || "N/A";

          const images = Array.from(document.querySelectorAll('img'))
            .map(img => img.src || img.dataset.src)
            .filter(src => src && src.includes('product'))
            .slice(0, 12)
            .join('; ');

          const info = {};
          document.querySelectorAll('div, li, p').forEach(el => {
            const text = el.innerText.trim();
            if (text.includes(':') && text.length < 200 && !text.includes('₹') && text.split(':').length === 2) {
              const [k, v] = text.split(':');
              info[k.trim()] = v.trim();
            }
          });

          return { title, price, mrp, discount, desc, images, info };
        });

        console.log(`   Title: ${data.title.substring(0, 60)}...`);
        console.log(`   Price: ${data.price} | MRP: ${data.mrp} | Offer: ${data.discount}`);

        if (data.title && data.price !== "N/A" && !data.title.includes("page isn’t working")) {
          const priceNum = parseFloat(data.price.replace(/[^\d.]/g, ''));
          const mrpNum = data.mrp !== "N/A" ? parseFloat(data.mrp.replace(/[^\d.]/g, '')) : priceNum;
          const offer = mrpNum > priceNum ? (((mrpNum - priceNum) / mrpNum) * 100).toFixed(1) + "%" : data.discount;

          results.push({
            name: data.title,
            link: p.link,
            image: data.images.split('; ')[0] || p.image,
            price: data.price,
            mrp: data.mrp,
            offer: offer,
            description: data.desc,
            images: data.images,
            ...data.info
          });

          Object.keys(data.info).forEach(k => extraKeys.add(k));
          console.log(`SUCCESS – ${data.title.substring(0, 50)}...`);
          success = true;
          break;
        }
      }
      await delay(8000, 12000);
    }

    if (!success) {
      console.log(`FAILED – ${p.name}`);
      results.push({ name: p.name + " (Blocked)", link: p.link, price: "N/A", mrp: "N/A", offer: "N/A" });
    }

    await tab.close();
    await delay(3000, 7000);
  }

  // 7. Save to Excel
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Zepto Full Category');

  const columns = [
    { header: "Name", key: "name", width: 60 },
    { header: "Link", key: "link", width: 90 },
    { header: "Image", key: "image", width: 60 },
    { header: "Price", key: "price", width: 15 },
    { header: "MRP", key: "mrp", width: 15 },
    { header: "Offer", key: "offer", width: 18 },
    { header: "Description", key: "description", width: 80 },
    { header: "All Images", key: "images", width: 100 },
  ];

  Array.from(extraKeys).sort().forEach(key => {
    columns.push({ header: key, key: key, width: 40 });
  });

  ws.columns = columns;
  results.forEach(product => ws.addRow(product));

  const filename = `ZEPTO_FULL_${new Date().toISOString().slice(0,10)}_(${results.length}_products).xlsx`;
  await wb.xlsx.writeFile(filename);

  console.log(`\nMISSION COMPLETE!`);
  console.log(`${results.length} products saved to → ${filename}`);
  await browser.close();
}

scrapeAllProducts().catch(err => {
  console.error("CRASHED:", err);
});