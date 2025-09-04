const puppeteer = require("puppeteer");
const ExcelJS = require("exceljs");

const url = "https://blinkit.com/cn/powdered-spices/cid/1557/50";

async function saveToExcel(products, filename = "blinkit_products.xlsx", append = false) {
  const workbook = new ExcelJS.Workbook();
  
  if (append) {
    try {
      await workbook.xlsx.readFile(filename);
    } catch (error) {
      console.log(`No existing file found or error reading ${filename}. Creating new workbook.`);
    }
  }

  const worksheet = workbook.getWorksheet("Products") || workbook.addWorksheet("Products");

  if (!append) {
    // Define columns based on all possible keys
    const allKeys = new Set();
    products.forEach((product) => {
      Object.keys(product).forEach((key) => allKeys.add(key));
    });

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

    const infoColumns = Array.from(allKeys)
      .filter((key) => !baseColumns.some((col) => col.key === key))
      .map((key) => ({ header: key, key, width: 40 }));

    worksheet.columns = [...baseColumns, ...infoColumns];
  }

  products.forEach((product) => {
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

async function scrapeProductData() {
  console.log("Launching browser...");
  const browser = await puppeteer.launch({
    headless: false,
    slowMo: 100,
    defaultViewport: { width: 1280, height: 800 },
  });

  const page = await browser.newPage();
  console.log("Navigating to URL...");
  await retryGoto(page, url);

  console.log("Scrolling to load all products...");
  let previousHeight = 0;
  let scrollAttempts = 0;
  const maxScrollAttempts = 50;
  while (scrollAttempts < maxScrollAttempts) {
    await page.evaluate(() => window.scrollBy(0, document.body.scrollHeight));
    await new Promise((resolve) => setTimeout(resolve, 2000));
    const newHeight = await page.evaluate(() => document.body.scrollHeight);
    if (newHeight === previousHeight) break;
    previousHeight = newHeight;
    scrollAttempts++;
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

    products.push({ name, pack, price, image: "N/A", link: "", description: "N/A", images: "N/A", mrp: "N/A", offer: "N/A" });
  }

  console.log(`Found ${products.length} products. Now fetching details for each product...`);

  const allProductDetails = [];
  let isFirstSave = true;

  // Handle Ctrl+C to save data and exit
  process.on("SIGINT", async () => {
    console.log("\nCaught interrupt signal. Saving collected data...");
    try {
      if (allProductDetails.length > 0) {
        await saveToExcel(allProductDetails, "blinkit_products.xlsx", !isFirstSave);
        console.log(`Saved ${allProductDetails.length} products to blinkit_products.xlsx`);
      } else {
        console.log("No products to save.");
      }
    } catch (error) {
      console.error(`Error saving on interrupt: ${error.message}`);
    }
    await browser.close();
    process.exit(0);
  });

  for (let i = 0; i < products.length; i++) {
    console.log(`Opening product page for: ${products[i].name}`);
    let retries = 3;
    let success = false;

    while (retries > 0 && !success) {
      try {
        // Click on the entire card using evaluate to avoid detached node issues
        await page.evaluate((sel, idx) => {
          const cards = document.querySelectorAll(sel);
          if (idx >= cards.length) return;
          cards[idx].click();
        }, cardSelector, i);

        // Wait for PDP content to load
        await page.waitForSelector("div.tw-text-600.tw-font-extrabold.tw-line-clamp-50", { timeout: 90000 });

        products[i].link = page.url();

        // Expand "View more details" if present
        await page.evaluate(() => {
          const buttons = Array.from(document.querySelectorAll("button"));
          const expandButton = buttons.find((b) => b.textContent.includes("View more details"));
          if (expandButton) expandButton.click();
        });
        await new Promise((resolve) => setTimeout(resolve, 1500));

        const productDetails = await page.evaluate(() => {
          const name = document.querySelector("div.tw-text-600.tw-font-extrabold.tw-line-clamp-50")?.textContent.trim() || "N/A";

          // Price (current selected variant)
          const priceElement = document.querySelector("div.tw-text-400.tw-font-bold");
          const price = priceElement ? priceElement.textContent.trim() : "N/A";

          // MRP if discounted (line-through)
          const mrpElement = document.querySelector("div.tw-text-300.tw-font-medium.tw-line-through");
          const mrp = mrpElement ? mrpElement.textContent.trim() : "N/A";

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

          // Description from Key Features if present
          const description = info["Key Features"] || "N/A";

          // Images (main and carousel)
          const images = Array.from(
            document.querySelectorAll('img[src^="https://cdn.grofers.com/cdn-cgi/image/f=auto,fit=scale-down,q=85,metadata=none,w="]')
          )
            .map((img) => img.src)
            .filter((src) => src.includes("/da/cms-assets/"))
            .join("; ") || "N/A";

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
        const priceNum = productDetails.price !== "N/A" ? parseFloat(productDetails.price.replace("₹", "")) : null;
        const mrpNum = productDetails.mrp !== "N/A" ? parseFloat(productDetails.mrp.replace("₹", "")) : null;
        const offer = priceNum && mrpNum ? (((mrpNum - priceNum) / mrpNum) * 100).toFixed(2) + "%" : "N/A";

        // Merge everything
        products[i] = {
          name: productDetails.name,
          pack: products[i].pack,
          link: products[i].link,
          image: productDetails.images.split("; ")[0] || "N/A",
          price: productDetails.price,
          mrp: productDetails.mrp,
          offer,
          description: productDetails.description,
          images: productDetails.images,
          ...productDetails.info, // spread all info fields dynamically
        };

        console.log(`Extracted details for ${products[i].name}:`, products[i]);
        allProductDetails.push(products[i]);
        success = true;

        // Save immediately after each product
        try {
          await saveToExcel([products[i]], "blinkit_products.xlsx", !isFirstSave);
          isFirstSave = false;
        } catch (error) {
          console.error(`Failed to save product ${products[i].name}: ${error.message}`);
        }

        // Go back to listing page
        await page.goBack({ waitUntil: "networkidle2" });
        await page.waitForSelector(cardSelector, { timeout: 90000 });
      } catch (error) {
        console.error(`Attempt ${4 - retries} failed for ${products[i].name}: ${error.message}`);
        retries--;
        if (retries === 0) {
          console.error(`All retries failed for ${products[i].name}. Saving partial data.`);
          allProductDetails.push(products[i]);
          try {
            await saveToExcel([products[i]], "blinkit_products.xlsx", !isFirstSave);
            isFirstSave = false;
          } catch (saveError) {
            console.error(`Failed to save partial data for ${products[i].name}: ${saveError.message}`);
          }
          // Try to go back to listing page
          try {
            await page.goBack({ waitUntil: "networkidle2" });
            await page.waitForSelector(cardSelector, { timeout: 90000 });
          } catch (goBackError) {
            console.error(`Failed to go back for ${products[i].name}: ${goBackError.message}`);
            // Reload listing page if goBack fails
            await retryGoto(page, url);
            await page.waitForSelector(cardSelector, { timeout: 90000 });
          }
        } else {
          // Wait before retrying
          await new Promise((resolve) => setTimeout(resolve, 2000));
        }
      }
    }
  }

  console.log("All product details extracted. Performing final save...");
  try {
    await saveToExcel(allProductDetails, "blinkit_products.xlsx", false);
    console.log(`Final save: Total products extracted: ${allProductDetails.length}. Data saved to Excel.`);
  } catch (error) {
    console.error(`Error in final save: ${error.message}`);
  }

  await browser.close();
}

async function retryGoto(page, url, retries = 3) {
  for (let i = 0; i < retries; i++) {
    try {
      await page.goto(url, { waitUntil: "networkidle2", timeout: 60000 });
      return;
    } catch (error) {
      console.error(`Attempt ${i + 1} failed: ${error.message}`);
      if (i === retries - 1) throw error;
      await new Promise((resolve) => setTimeout(resolve, 2000));
    }
  }
}

scrapeProductData().catch(async (error) => {
  console.error("Error in scrapeProductData:", error);
  process.exit(1);
});