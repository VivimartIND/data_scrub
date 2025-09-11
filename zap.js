const puppeteer = require("puppeteer");
const ExcelJS = require("exceljs");

const url =
  "https://www.zeptonow.com/cn/bath-body/shower-gels/cid/26e64367-19ad-4f80-a763-42599d4215ee/scid/e8801cfd-49d0-4388-bc29-f83258204d39";

async function scrapeProductData() {
  console.log("Launching browser...");
  const browser = await puppeteer.launch({
    headless: false,
    slowMo: 200,
    defaultViewport: { width: 1280, height: 800 },
    args: ['--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36']
  });

  const page = await browser.newPage();
  console.log("Navigating to URL...");
  await page.goto(url, { waitUntil: "networkidle2" });
  await randomWait(2000, 5000);

  console.log("Scrolling to load all products...");
  let previousHeight = 0;
  while (true) {
    await page.evaluate(() => window.scrollBy(0, window.innerHeight));
    await randomWait(1000, 3000);
    const newHeight = await page.evaluate(() => document.body.scrollHeight);
    if (newHeight === previousHeight) break;
    previousHeight = newHeight;
  }
  await randomWait(3000, 6000);

  console.log("Extracting product details from listing page...");
  const cardSelector = 'a[href^="/pn/"]';
  const cardHandles = await page.$$(cardSelector);
  const products = [];
  for (let i = 0; i < cardHandles.length; i++) {
    const name = await page.evaluate((sel, idx) => {
      const cards = document.querySelectorAll(sel);
      if (idx >= cards.length) return "N/A";
      const imgEl = cards[idx].querySelector("img");
      return imgEl ? imgEl.alt.trim() : "N/A";
    }, cardSelector, i);

    const link = await page.evaluate((sel, idx) => {
      const cards = document.querySelectorAll(sel);
      if (idx >= cards.length) return "N/A";
      return "https://www.zeptonow.com" + cards[idx].getAttribute("href");
    }, cardSelector, i);

    const image = await page.evaluate((sel, idx) => {
      const cards = document.querySelectorAll(sel);
      if (idx >= cards.length) return "N/A";
      const imgEl = cards[idx].querySelector("img");
      return imgEl ? (imgEl.src || imgEl.getAttribute("data-src") || "N/A") : "N/A";
    }, cardSelector, i);

    products.push({ name, link, image });
  }

  console.log(
    `Found ${products.length} products. Now fetching details for each product...`
  );

  const allKeys = new Set();

  for (let i = 0; i < products.length; i++) {
    await randomWait(3000, 8000);
    console.log(`Opening product page for: ${products[i].name}`);
    try {
      // Hover over the card
      await page.evaluate((sel, idx) => {
        const cards = document.querySelectorAll(sel);
        if (idx >= cards.length) return;
        cards[idx].dispatchEvent(new MouseEvent('mouseover', { bubbles: true }));
      }, cardSelector, i);
      await randomWait(500, 1500);

      // Click the card
      await page.evaluate((sel, idx) => {
        const cards = document.querySelectorAll(sel);
        if (idx >= cards.length) return;
        cards[idx].click();
      }, cardSelector, i);

      // Wait for PDP
      await page.waitForSelector('h1.font-semibold', { timeout: 60000 });
      await randomWait(2000, 5000);

      // Wait for info section if possible
      await page.waitForSelector('div.flex.items-start.gap-3', { timeout: 10000 }).catch(() => console.log("Info section not found or timed out"));

      const productDetails = await page.evaluate(() => {
        const name =
          document.querySelector("h1.font-semibold")?.textContent.trim() ||
          "N/A";

        // Price (current selected variant)
        const priceElement = Array.from(document.querySelectorAll("p, span")).find(
          (el) => el.textContent.includes("₹") && !el.textContent.includes("line-through") && !el.textContent.includes("MRP")
        );
        const price = priceElement?.textContent.match(/₹\d+/)?.[0] || "N/A";

        // MRP if discounted (line-through) or alternative
        let mrpElement = document.querySelector("span.line-through");
        let mrp = mrpElement ? mrpElement.textContent.replace("₹", "").trim() : null;
        if (!mrp) {
          const mrpAlt = Array.from(document.querySelectorAll("p, span")).find(
            (el) => el.textContent.includes("MRP ₹")
          );
          mrp = mrpAlt ? mrpAlt.textContent.match(/₹\d+/)?.[0].replace("₹", "").trim() : "N/A";
        } else {
          mrp = mrp || "N/A";
        }

        // Description
        const descriptionElement = document.querySelector(
          'meta[itemprop="description"]'
        );
        const description =
          descriptionElement?.getAttribute("content") || "N/A";

        // Images
        const images =
          Array.from(
            document.querySelectorAll('button[aria-label^="image-preview-"] img')
          )
            .map((img) => img.src)
            .join("; ") || "N/A";

        // Information section
        const info = {};
        document.querySelectorAll("div.flex.items-start.gap-3").forEach((el) => {
          const key = el.querySelector("h3")?.textContent.trim() || null;
          const value = el.querySelector("p")?.textContent.trim() || null;
          if (key && value) info[key] = value;
        });

        // If no info from flex, try ul li format
        if (Object.keys(info).length === 0) {
          document.querySelectorAll("ul li").forEach((li) => {
            const text = li.textContent.trim();
            if (text.includes(":")) {
              const [key, value] = text.split(":").map((s) => s.trim());
              if (key && value) info[key] = value;
            }
          });
        }

        return {
          name,
          price,
          mrp,
          description,
          images,
          info,
        };
      });

      // Calculate offer %
      const priceNum =
        productDetails.price !== "N/A"
          ? parseFloat(productDetails.price.replace("₹", ""))
          : null;
      const mrpNum =
        productDetails.mrp !== "N/A"
          ? parseFloat(productDetails.mrp)
          : null;
      const offer =
        priceNum && mrpNum && mrpNum > priceNum
          ? (((mrpNum - priceNum) / mrpNum) * 100).toFixed(2) + "%"
          : "N/A";

      // Merge everything
      products[i] = {
        name: productDetails.name,
        link: products[i].link,
        image: productDetails.images.split("; ")[0] || "N/A",
        price: productDetails.price,
        mrp: productDetails.mrp,
        offer: offer,
        description: productDetails.description,
        images: productDetails.images,
        ...productDetails.info,
      };

      // collect keys
      Object.keys(productDetails.info).forEach((k) => allKeys.add(k));
      console.log(`Extracted details for ${products[i].name}:`, products[i]);

      // Go back
      await page.goBack({ waitUntil: "networkidle2" });
      await page.waitForSelector(cardSelector, { timeout: 60000 });
      await randomWait(3000, 7000);
    } catch (error) {
      console.error(
        `Error while scraping product page for ${products[i].name}:`,
        error
      );
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

  const infoColumns = Array.from(allKeys).map((key) => ({
    header: key,
    key: key,
    width: 40,
  }));

  worksheet.columns = [...baseColumns, ...infoColumns];

  products.forEach((product) => {
    worksheet.addRow(product);
  });

  await workbook.xlsx.writeFile("zepto_products.xlsx");
  console.log("Data saved to zepto_products.xlsx ✅");

  await browser.close();
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
      console.error(`Attempt ${i + 1} failed: ${error.message}`);
      if (i === retries - 1) throw error;
    }
  }
}

scrapeProductData().catch((error) => {
  console.error("Error in scrapeProductData:", error);
});