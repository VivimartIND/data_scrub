const puppeteer = require("puppeteer");
const ExcelJS = require("exceljs");

const url = "https://blinkit.com/cn/vegetables-fruits/fresh-vegetables/cid/1487/1489";

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

  console.log("Scrolling product container to load all products...");
  let previousHeight = 0;
  while (true) {
    const newHeight = await page.evaluate(() => {
      const container = document.querySelector('#plpContainer');
      if (!container) return document.body.scrollHeight;
      container.scrollBy(0, container.scrollHeight);
      return container.scrollHeight;
    });
    await randomWait(3000, 6000);
    if (newHeight === previousHeight) break;
    previousHeight = newHeight;
  }
  await randomWait(3000, 6000);

  // Scroll main page as fallback
  console.log("Scrolling main page to ensure all products are loaded...");
  previousHeight = 0;
  while (true) {
    await page.evaluate(() => window.scrollBy(0, window.innerHeight));
    await randomWait(3000, 6000);
    const newHeight = await page.evaluate(() => document.body.scrollHeight);
    if (newHeight === previousHeight) break;
    previousHeight = newHeight;
  }
  await randomWait(3000, 6000);

  // Click "Show more" or similar buttons if present
  console.log("Clicking 'Show more' if present...");
  let loadMoreClicked = 0;
  while (true) {
    const loadMore = await page.evaluate(() => {
      const elements = Array.from(document.querySelectorAll("button, div, a"));
      const button = elements.find(el => {
        const text = el.textContent.toLowerCase();
        return text.includes("show more") || text.includes("load more") || text.includes("view more") || text.includes("see more");
      });
      if (button) {
        button.click();
        return true;
      }
      return false;
    });
    if (!loadMore) break;
    await randomWait(3000, 6000);
    loadMoreClicked++;
    if (loadMoreClicked > 50) break;
  }

  console.log("Extracting product details from listing page...");
  const cardSelector = 'div.tw-w-full.tw-px-3[data-pf="reset"]';
  const cardHandles = await page.$$(cardSelector);
  const products = [];
  for (let i = 0; i < cardHandles.length; i++) {
    const name = await page.evaluate((sel, idx) => {
      const cards = document.querySelectorAll(sel);
      if (idx >= cards.length) return "N/A";
      const nameEl = cards[idx].querySelector("div.tw-text-300.tw-font-semibold.tw-line-clamp-2");
      return nameEl ? nameEl.textContent.trim() : "N/A";
    }, cardSelector, i);

    const pack = await page.evaluate((sel, idx) => {
      const cards = document.querySelectorAll(sel);
      if (idx >= cards.length) return "N/A";
      const packEl = cards[idx].querySelector("div.tw-text-200.tw-font-medium.tw-line-clamp-1");
      return packEl ? packEl.textContent.trim() : "N/A";
    }, cardSelector, i);

    const price = await page.evaluate((sel, idx) => {
      const cards = document.querySelectorAll(sel);
      if (idx >= cards.length) return "N/A";
      const priceEl = cards[idx].querySelector("div.tw-text-200.tw-font-semibold");
      return priceEl ? priceEl.textContent.trim() : "N/A";
    }, cardSelector, i);

    products.push({ name, pack, price, image: "N/A", link: "" });
  }

  console.log(`Found ${products.length} products. Now fetching details for each product...`);

  const allKeys = new Set();

  for (let i = 0; i < products.length; i++) {
    await randomWait(3000, 8000);
    console.log(`Opening product page for: ${products[i].name}`);

    try {
      await page.evaluate((sel, idx) => {
        const cards = document.querySelectorAll(sel);
        if (idx >= cards.length) return;
        cards[idx].dispatchEvent(new MouseEvent('mouseover', { bubbles: true }));
      }, cardSelector, i);
      await randomWait(500, 1500);

      await page.evaluate((sel, idx) => {
        const cards = document.querySelectorAll(sel);
        if (idx >= cards.length) return;
        cards[idx].click();
      }, cardSelector, i);

      await page.waitForSelector('div.tw-text-600.tw-font-extrabold.tw-line-clamp-50', { timeout: 60000 });
      await randomWait(2000, 5000);

      products[i].link = page.url();

      await page.evaluate(() => {
        const buttons = Array.from(document.querySelectorAll("button"));
        const expandButton = buttons.find(b => b.textContent.includes("View more details"));
        if (expandButton) expandButton.click();
      });
      await randomWait(1000, 3000);

      const productDetails = await page.evaluate(() => {
        const name = document.querySelector("div.tw-text-600.tw-font-extrabold.tw-line-clamp-50")?.textContent.trim() || "N/A";

        // Price
        const priceElement = document.querySelector("div.tw-text-400.tw-font-bold");
        const price = priceElement ? priceElement.textContent.trim() : "N/A";

        // MRP (try line-through first, then fallback to other elements)
        let mrpElement = document.querySelector("div.tw-text-300.tw-font-medium.tw-line-through");
        let mrp = mrpElement ? mrpElement.textContent.trim() : null;
        if (!mrp) {
          const mrpAlt = Array.from(document.querySelectorAll("div")).find(
            el => el.textContent.includes("MRP") && el.textContent.includes("₹")
          );
          mrp = mrpAlt ? mrpAlt.textContent.match(/₹\d+/)?.[0] || "N/A" : "N/A";
        }
        if (mrp === "N/A" && price !== "N/A") {
          mrp = price; // Assume MRP = price if no discount
        }

        // Information section
        const info = {};
        document.querySelectorAll("div.tw-flex.tw-flex-col.tw-gap-1\\.5.tw-break-words").forEach((el) => {
          const key = el.querySelector("div.tw-text-300.tw-font-medium")?.textContent.trim() || null;
          const value = el.querySelector("div.tw-text-200.tw-font-regular.tw-whitespace-pre-wrap")?.textContent.trim() || null;
          if (key && value) info[key] = value;
        });

        // Highlights section
        document.querySelectorAll("div.tw-bg-grey-100.tw-w-fit.tw-flex-none.tw-rounded-2xl.tw-px-4.tw-pt-5").forEach((el) => {
          const key = el.querySelector("div.tw-text-100.tw-font-medium")?.textContent.trim() || null;
          const value = el.querySelector("div.tw-text-300.tw-font-semibold")?.textContent.trim() || null;
          if (key && value) info[key] = value;
        });

        // Description (prefer Description, fallback to Key Features)
        const description = info["Description"] || info["Key Features"] || "N/A";

        // Images
        const images = Array.from(
          document.querySelectorAll('img[src^="https://cdn.grofers.com/cdn-cgi/image/f=auto,fit=scale-down,q=85,metadata=none,w="]')
        ).map((img) => img.src).filter(src => src.includes("/da/cms-assets/")).join("; ") || "N/A";

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
          ? parseFloat(productDetails.mrp.replace("₹", ""))
          : null;
      const offer =
        priceNum && mrpNum && mrpNum > priceNum
          ? (((mrpNum - priceNum) / mrpNum) * 100).toFixed(2) + "%"
          : "N/A";

      // Merge everything
      products[i] = {
        name: productDetails.name,
        pack: products[i].pack,
        link: products[i].link,
        image: productDetails.images.split("; ")[0] || "N/A",
        price: productDetails.price,
        mrp: productDetails.mrp,
        offer: offer,
        description: productDetails.description,
        images: productDetails.images,
        ...productDetails.info,
      };

      Object.keys(productDetails.info).forEach((k) => allKeys.add(k));
      console.log(`Extracted details for ${products[i].name}:`, products[i]);

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
    { header: "Pack", key: "pack", width: 20 },
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

  await workbook.xlsx.writeFile("blinkit_products.xlsx");
  console.log("Data saved to blinkit_products.xlsx ✅");

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