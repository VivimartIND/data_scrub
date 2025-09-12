const puppeteer = require("puppeteer");
const ExcelJS = require("exceljs");

const url =
  "https://www.zeptonow.com/cn/masala-dry-fruits-more/masala-dry-fruits-more/cid/0c2ccf87-e32c-4438-9560-8d9488fc73e0/scid/8b44cef2-1bab-407e-aadd-29254e6778fa";

const TIMEOUTS = {
  navigation: 30000,
  selector: 5000,
};

async function scrapeProductData() {
  console.log("Running zap.js version 2025-09-11-01");
  let pLimitFn;
  try {
    const { default: pLimit } = await import("p-limit");
    pLimitFn = pLimit;
    console.log("p-limit loaded successfully for concurrent processing.");
  } catch (error) {
    console.warn("Warning: Failed to load p-limit. Falling back to sequential mode. Install p-limit with `npm install p-limit`. Error: " + error.message);
    pLimitFn = null;
  }

  console.log("Launching browser...");
  const browser = await puppeteer.launch({
    headless: true,
    defaultViewport: { width: 1280, height: 800 },
    executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || undefined,
    args: [
      "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    ],
  });

  const page = await browser.newPage();
  await page.setCookie({
    name: "session",
    value: "default",
    domain: "www.zeptonow.com",
    path: "/",
  });
  console.log("Navigating to URL...");
  await retryGoto(page, url);

  console.log("Scrolling to load all products...");
  let previousHeight = 0;
  let scrollAttempts = 0;
  const maxScrollAttempts = 10;
  while (scrollAttempts < maxScrollAttempts) {
    await page.evaluate(() => window.scrollBy(0, window.innerHeight * 2));
    await randomWait(500, 1500);
    const newHeight = await page.evaluate(() => document.body.scrollHeight);
    if (newHeight === previousHeight) break;
    previousHeight = newHeight;
    scrollAttempts++;
  }
  await randomWait(1000, 3000);

  try {
    const loadMore = await page.evaluateHandle(() => {
      const buttons = Array.from(document.querySelectorAll('button, [class*="load-more"], [class*="more"]'));
      return buttons.find((btn) => btn.textContent.toLowerCase().includes("load more") || btn.textContent.toLowerCase().includes("more"));
    });
    if (loadMore.asElement()) {
      console.log("Clicking Load More button...");
      await loadMore.click();
      await randomWait(2000, 4000);
    } else {
      console.log("No Load More button found.");
    }
  } catch (error) {
    console.warn("Failed to check for Load More button:", error.message);
  }

  console.log("Extracting product details from listing page...");
  const cardSelector = 'a[href^="/pn/"]';
  const cardHandles = await page.$$(cardSelector);
  const products = [];
  for (let i = 0; i < cardHandles.length; i++) {
    const name = await page.evaluate(
      (sel, idx) => {
        const cards = document.querySelectorAll(sel);
        if (idx >= cards.length) return "N/A";
        const imgEl = cards[idx].querySelector("img");
        return imgEl ? imgEl.alt.trim() : "N/A";
      },
      cardSelector,
      i
    );

    const link = await page.evaluate(
      (sel, idx) => {
        const cards = document.querySelectorAll(sel);
        if (idx >= cards.length) return "N/A";
        return "https://www.zeptonow.com" + cards[idx].getAttribute("href");
      },
      cardSelector,
      i
    );

    const image = await page.evaluate(
      (sel, idx) => {
        const cards = document.querySelectorAll(sel);
        if (idx >= cards.length) return "N/A";
        const imgEl = cards[idx].querySelector("img");
        return imgEl ? (imgEl.src || imgEl.getAttribute("data-src") || "N/A") : "N/A";
      },
      cardSelector,
      i
    );

    products.push({ name, link, image });
  }

  if (products.length === 0) {
    console.error("No products found. Logging page content for debugging...");
    const pageContent = await page.evaluate(() => document.body.innerHTML.substring(0, 1000)).catch(() => "N/A");
    console.log(`Main page content: ${pageContent}`);
  }
  console.log(`Found ${products.length} products. Now fetching details for each product...`);

  const allKeys = new Set();

  if (pLimitFn) {
    const limit = pLimitFn(3);
    const productBatches = [];
    for (let i = 0; i < products.length; i += 3) {
      productBatches.push(products.slice(i, i + 3));
    }
    for (const batch of productBatches) {
      await Promise.all(
        batch.map((product, idx) =>
          limit(async () => {
            const globalIdx = products.indexOf(product);
            console.log(`Processing product ${globalIdx + 1}/${products.length}: ${product.name}`);
            await processProductPage(browser, product, globalIdx, products, allKeys);
          })
        )
      );
      await randomWait(1000, 2000);
    }
  } else {
    for (let i = 0; i < products.length; i++) {
      console.log(`Processing product ${i + 1}/${products.length}: ${products[i].name}`);
      await processProductPage(browser, products[i], i, products, allKeys);
    }
  }

  console.log("All product details extracted. Saving to Excel...");

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Products");

  const baseColumns = [
    { header: "Name", key: "name", width: 40 },
    { header: "Link", key: "link", width: 60 },
    { header: "Image", key: "image", width: 50 },
    { header: "Price", key: "price", width: 15 },
    { header: "MRP", key: "mrp", width: 15 },
    { header: "Offer", key: "offer", width: 15 },
    { header: "Description", key: "description", width: 60 },
    { header: "Images", key: "images", width: 60 },
  ];

  const infoColumns = Array.from(allKeys).sort().map((key) => ({
    header: key,
    key: key,
    width: 40,
  }));

  worksheet.columns = [...baseColumns, ...infoColumns];

  products.forEach((product, index) => {
    console.log(`Saving product ${index + 1}/${products.length}: ${product.name}`);
    worksheet.addRow(product);
  });

  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  await workbook.xlsx.writeFile(`zepto_vegetables_${timestamp}.xlsx`);
  console.log(`Data saved to zepto_vegetables_${timestamp}.xlsx ✅`);

  await browser.close();
}

async function processProductPage(browser, product, index, products, allKeys) {
  const page = await browser.newPage();
  let attempts = 0;
  const maxAttempts = 2;
  let productDetails = null;

  while (attempts < maxAttempts && !productDetails) {
    try {
      await retryGoto(page, product.link, 2);
      await randomWait(2000, 4000);

      await page.waitForSelector("body", { timeout: TIMEOUTS.navigation }).catch(() =>
        console.warn(`Body selector not found for ${product.name} at ${product.link}`)
      );

      productDetails = await page.evaluate(() => {
        const name =
          document.querySelector("h1.font-semibold")?.textContent.trim() ||
          document.querySelector("h1, h2, h3, [class*='title'], [data-testid='pdp-product-name']")?.textContent.trim() ||
          document.querySelector('meta[itemprop="name"]')?.getAttribute("content")?.trim() ||
          document.querySelector("title")?.textContent.trim() ||
          "N/A";

        const priceElement = Array.from(document.querySelectorAll("p, span, div, [class*='price']")).find(
          (el) =>
            el.textContent.match(/₹\d+(\.\d+)?/) &&
            !el.textContent.includes("line-through") &&
            !el.textContent.includes("MRP") &&
            !el.textContent.includes("M.R.P")
        );
        const price = priceElement?.textContent.match(/₹\d+(\.\d+)?/)?.[0] || "N/A";

        let mrpElement = document.querySelector("span.line-through, [class*='mrp'], [class*='original-price']");
        let mrp = mrpElement ? mrpElement.textContent.replace("₹", "").trim() : null;
        if (!mrp) {
          const mrpAlt = Array.from(document.querySelectorAll("p, span, div")).find(
            (el) => el.textContent.includes("MRP ₹") || el.textContent.includes("M.R.P")
          );
          mrp = mrpAlt ? mrpAlt.textContent.match(/₹\d+/)?.[0].replace("₹", "").trim() : "N/A";
        } else {
          mrp = mrp || "N/A";
        }

        const descriptionElement =
          document.querySelector('meta[itemprop="description"]') ||
          document.querySelector("p[class*='description'], div[class*='description']");
        const description = descriptionElement
          ? descriptionElement.getAttribute("content") || descriptionElement.textContent.trim() || "N/A"
          : "N/A";

        const images =
          Array.from(
            document.querySelectorAll('button[aria-label^="image-preview-"] img, img[class*="product-image"], img[src*="product"]')
          )
            .map((img) => img.src || img.getAttribute("data-src"))
            .filter(Boolean)
            .join("; ") || "N/A";

        const info = {};
        document
          .querySelectorAll("div[class*='product-detail'], div[class*='info'], div[class*='detail'], div.flex, section[class*='info']")
          .forEach((el) => {
            const key = el.querySelector("h3, h4, strong, span[class*='key'], [class*='label']")?.textContent.trim() || null;
            const value = el.querySelector("p, span[class*='value'], [class*='detail']")?.textContent.trim() || null;
            if (key && value) info[key] = value;
          });

        if (Object.keys(info).length === 0) {
          document.querySelectorAll("ul li, div[class*='info'] > *, [class*='detail'] > *").forEach((el) => {
            const text = el.textContent.trim();
            if (text.includes(":")) {
              const [key, value] = text.split(":").map((s) => s.trim());
              if (key && value) info[key] = value;
            }
          });
        }

        return { name, price, mrp, description, images, info };
      });

      if (productDetails.name === "N/A" && productDetails.price === "N/A") {
        console.warn(`Incomplete data for ${product.name} at ${product.link}. Retrying...`);
        productDetails = null;
        attempts++;
        await randomWait(2000, 4000);
      }
    } catch (error) {
      console.error(`Error scraping product page for ${product.name} at ${product.link} (attempt ${attempts + 1}):`, error.message);
      attempts++;
      if (attempts < maxAttempts) {
        console.log(`Retrying ${product.name} at ${product.link}...`);
        await randomWait(2000, 4000);
      }
    }
  }

  if (!productDetails) {
    console.error(`Failed to scrape ${product.name} at ${product.link} after ${maxAttempts} attempts.`);
    const pageContent = await page.evaluate(() => document.body.innerHTML.substring(0, 1000)).catch(() => "N/A");
    console.log(`Partial page content for ${product.link}: ${pageContent}`);
    productDetails = {
      name: "N/A",
      price: "N/A",
      mrp: "N/A",
      description: "N/A",
      images: "N/A",
      info: {},
    };
  }

  const priceNum = productDetails.price !== "N/A" ? parseFloat(productDetails.price.replace("₹", "")) : null;
  const mrpNum = productDetails.mrp !== "N/A" ? parseFloat(productDetails.mrp) : null;
  const offer =
    priceNum && mrpNum && mrpNum > priceNum
      ? (((mrpNum - priceNum) / mrpNum) * 100).toFixed(2) + "%"
      : "N/A";

  products[index] = {
    name: productDetails.name,
    link: product.link,
    image: productDetails.images.split("; ")[0] || "N/A",
    price: productDetails.price,
    mrp: productDetails.mrp,
    offer,
    description: productDetails.description,
    images: productDetails.images,
    ...productDetails.info,
  };

  Object.keys(productDetails.info).forEach((k) => allKeys.add(k));
  console.log(`Extracted details for ${products[index].name}:`, products[index]);

  await page.close();
}

async function randomWait(min, max) {
  const waitTime = Math.floor(Math.random() * (max - min + 1)) + min;
  return new Promise((resolve) => setTimeout(resolve, waitTime));
}

async function retryGoto(page, url, retries = 3) {
  for (let i = 0; i < retries; i++) {
    try {
      await page.goto(url, { waitUntil: "networkidle2" });
      return;
    } catch (error) {
      console.error(`Attempt ${i + 1} failed for URL ${url}: ${error.message}`);
      if (i === retries - 1) throw error;
    }
  }
}

scrapeProductData().catch((error) => {
  console.error("Error in scrapeProductData:", error);
});
