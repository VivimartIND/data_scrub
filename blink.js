const puppeteer = require("puppeteer");
const ExcelJS = require("exceljs");

const url = "https://blinkit.com/cn/vegetables-fruits/fresh-vegetables/cid/1487/1489";

async function scrapeProductData() {
  console.log("Launching browser...");
  const browser = await puppeteer.launch({
    headless: false,
    slowMo: 200, // Increased slowMo for more human-like actions
    defaultViewport: { width: 1280, height: 800 },
    args: ['--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'] // Realistic user-agent
  });

  const page = await browser.newPage();
  console.log("Navigating to URL...");
  await page.goto(url, { waitUntil: "networkidle2" });
  await randomWait(2000, 5000); // Random initial wait after loading

  console.log("Scrolling to load all products...");
  let previousHeight = 0;
  while (true) {
    await page.evaluate(() => window.scrollBy(0, window.innerHeight)); // Scroll by viewport height for natural scrolling
    await randomWait(3000, 5000); // Increased random wait between scrolls
    const newHeight = await page.evaluate(() => document.body.scrollHeight);
    if (newHeight === previousHeight) break;
    previousHeight = newHeight;
  }
  await randomWait(3000, 6000); // Extra wait after full scroll

  // Attempt to click "Show more" or similar buttons if present
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
    if (loadMoreClicked > 50) break; // Safety limit
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

    // Image not present in provided listing snippet, will fetch from PDP
    products.push({ name, pack, price, image: "N/A", link: "" });
  }

  console.log(`Found ${products.length} products. Now fetching details for each product...`);

  const allKeys = new Set(); // collect all possible info fields dynamically

  for (let i = 0; i < products.length; i++) {
    await randomWait(3000, 8000); // Random wait before processing each product to mimic human browsing
    console.log(`Opening product page for: ${products[i].name}`);

    try {
      // Hover over the card before clicking to simulate human interaction
      await page.evaluate((sel, idx) => {
        const cards = document.querySelectorAll(sel);
        if (idx >= cards.length) return;
        cards[idx].dispatchEvent(new MouseEvent('mouseover', { bubbles: true }));
      }, cardSelector, i);
      await randomWait(500, 1500); // Short wait after hover

      // Click on the entire card using evaluate to avoid detached node issues
      await page.evaluate((sel, idx) => {
        const cards = document.querySelectorAll(sel);
        if (idx >= cards.length) return;
        cards[idx].click();
      }, cardSelector, i);

      // Wait for PDP content to load
      await page.waitForSelector('div.tw-text-600.tw-font-extrabold.tw-line-clamp-50', { timeout: 60000 });
      await randomWait(2000, 5000); // Random wait after PDP load

      products[i].link = page.url();

      // Expand "View more details" if present
      await page.evaluate(() => {
        const buttons = Array.from(document.querySelectorAll("button"));
        const expandButton = buttons.find(b => b.textContent.includes("View more details"));
        if (expandButton) expandButton.click();
      });
      await randomWait(1000, 3000); // Random wait after expanding details

      const productDetails = await page.evaluate(() => {
        const name = document.querySelector("div.tw-text-600.tw-font-extrabold.tw-line-clamp-50")?.textContent.trim() || "N/A";

        // Price (current selected variant)
        const priceElement = document.querySelector("div.tw-text-400.tw-font-bold");
        const price = priceElement ? priceElement.textContent.trim() : "N/A";

        // MRP if discounted (line-through)
        const mrpElement = document.querySelector("div.tw-text-300.tw-font-medium.tw-line-through");
        let mrp = mrpElement ? mrpElement.textContent.trim() : "N/A";

        // If no MRP, assume it's the same as price (for non-discounted items)
        if (mrp === "N/A") {
          mrp = price;
        }

        // Information section (Product Details)
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

        // Description from Description or Key Features if present
        const description = info["Description"] || info["Key Features"] || "N/A";

        // Images (main and carousel)
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
        ...productDetails.info, // spread all info fields dynamically
      };

      // collect keys for Excel headers
      Object.keys(productDetails.info).forEach((k) => allKeys.add(k));
      console.log(`Extracted details for ${products[i].name}:`, products[i]);

      // Go back to listing page
      await page.goBack({ waitUntil: "networkidle2" });
      await page.waitForSelector(cardSelector, { timeout: 60000 });
      await randomWait(3000, 7000); // Random wait after going back
    } catch (error) {
      console.error(
        `Error while scraping product page for ${products[i].name}:`,
        error
      );
    }
  }

  console.log("All product details extracted. Saving to Excel...");

  // Create Excel workbook
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Products");

  // Basic columns
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

  // Dynamic info columns
  const infoColumns = Array.from(allKeys).map((key) => ({
    header: key,
    key: key,
    width: 40,
  }));

  worksheet.columns = [...baseColumns, ...infoColumns];

  // Add rows
  products.forEach((product) => {
    worksheet.addRow(product);
  });

  await workbook.xlsx.writeFile("blinkit_products.xlsx");
  console.log("Data saved to blinkit_products.xlsx ✅");

  await browser.close();
}

// Helper for random waits (min, max in ms)
async function randomWait(min, max) {
  const waitTime = Math.floor(Math.random() * (max - min + 1)) + min;
  return new Promise((resolve) => setTimeout(resolve, waitTime));
}

// retry navigation helper
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