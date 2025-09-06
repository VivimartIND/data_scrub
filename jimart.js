const puppeteer = require("puppeteer");
const ExcelJS = require("exceljs");

const url = "https://www.jiomart.com/c/groceries/biscuits-drinks-packaged-foods/chips-namkeens/29000?prod_mart_master_vertical_products_popularity%5Bpage%5D=3";

// Custom delay function
const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

async function saveToExcel(products, filename = "jiomart_products.xlsx", append = false) {
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
    const allKeys = new Set();
    products.forEach((product) => {
      Object.keys(product).forEach((key) => allKeys.add(key));
    });

    const baseColumns = [
      { header: "Name", key: "name", width: 40 },
      { header: "Brand", key: "brand", width: 20 },
      { header: "Price", key: "price", width: 15 },
      { header: "MRP", key: "mrp", width: 15 },
      { header: "Offer", key: "offer", width: 15 },
      { header: "Seller", key: "seller", width: 30 },
      { header: "Description", key: "description", width: 60 },
      { header: "Images", key: "images", width: 60 },
      { header: "Link", key: "link", width: 60 },
      { header: "Image", key: "image", width: 50 },
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

async function scrapeJioMart() {
  console.log("Launching browser...");
  const browser = await puppeteer.launch({
    headless: false,
    slowMo: 100,
    defaultViewport: { width: 1280, height: 800 },
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  });

  const page = await browser.newPage();
  await page.setUserAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");
  console.log("Navigating to URL...");
  await retryGoto(page, url);

  console.log("Scrolling to load all products...");
  let previousHeight = 0;
  let scrollAttempts = 0;
  const maxScrollAttempts = 50;
  while (scrollAttempts < maxScrollAttempts) {
    const newHeight = await page.evaluate(() => {
      window.scrollBy(0, document.body.scrollHeight);
      return document.body.scrollHeight;
    });
    await delay(2000);
    if (newHeight === previousHeight) {
      console.log("No more products to load.");
      break;
    }
    previousHeight = newHeight;
    scrollAttempts++;
    console.log(`Scroll attempt ${scrollAttempts}, page height: ${newHeight}`);
  }

  console.log("Extracting product cards...");
  const cardSelector = "li.ais-InfiniteHits-item";
  await page.waitForSelector(cardSelector, { timeout: 15000 }).catch(() => console.log("No cards found or timeout."));
  const cardHandles = await page.$$(cardSelector);
  console.log(`Found ${cardHandles.length} product cards.`);

  const products = [];
  for (let i = 0; i < cardHandles.length; i++) {
    const product = await page.evaluate((sel, idx) => {
      const cards = document.querySelectorAll(sel);
      if (idx >= cards.length) return null;

      const card = cards[idx];
      const nameEl = card.querySelector("div.plp-card-details-name");
      const priceEl = card.querySelector("span.jm-heading-xxs");
      const mrpEl = card.querySelector("span.jm-body-xxs.jm-fc-primary-grey-60.line-through");
      const offerEl = card.querySelector("span.jm-badge");
      const imageEl = card.querySelector("img.lazyautosizes");
      const linkEl = card.querySelector("a.plp-card-wrapper");

      console.log(`Card ${idx + 1} - Name: ${nameEl ? nameEl.textContent.trim() : "Not found"}`);
      console.log(`Card ${idx + 1} - Price: ${priceEl ? priceEl.textContent.trim() : "Not found"}`);
      console.log(`Card ${idx + 1} - MRP: ${mrpEl ? mrpEl.textContent.trim() : "Not found"}`);
      console.log(`Card ${idx + 1} - Offer: ${offerEl ? offerEl.textContent.trim() : "Not found"}`);
      console.log(`Card ${idx + 1} - Image: ${imageEl ? imageEl.src : "Not found"}`);
      console.log(`Card ${idx + 1} - Link: ${linkEl ? linkEl.href : "Not found"}`);

      return {
        name: nameEl ? nameEl.textContent.trim() : "N/A",
        price: priceEl ? priceEl.textContent.trim() : "N/A",
        mrp: mrpEl ? mrpEl.textContent.trim() : "N/A",
        offer: offerEl ? offerEl.textContent.trim() : "N/A",
        image: imageEl ? imageEl.src : "N/A",
        link: linkEl ? linkEl.href : "N/A",
        brand: "N/A",
        seller: "N/A",
        description: "N/A",
        images: "N/A",
      };
    }, cardSelector, i);

    if (product) products.push(product);
  }

  console.log(`Found ${products.length} products. Fetching details...`);

  const allProductDetails = [];
  let isFirstSave = true;

  // Handle Ctrl+C
  process.on("SIGINT", async () => {
    console.log("\nCaught interrupt signal. Saving collected data...");
    try {
      if (allProductDetails.length > 0) {
        await saveToExcel(allProductDetails, "jiomart_products.xlsx", !isFirstSave);
        console.log(`Saved ${allProductDetails.length} products to jiomart_products.xlsx`);
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
    if (products[i].link === "N/A") {
      console.log(`Skipping product ${i + 1}: No valid link.`);
      allProductDetails.push(products[i]);
      continue;
    }

    console.log(`Opening product page for: ${products[i].name}`);
    let retries = 3;
    let success = false;
    let detailPage;

    while (retries > 0 && !success) {
      try {
        detailPage = await browser.newPage();
        await retryGoto(detailPage, products[i].link);

        // Wait for product name
        await detailPage.waitForSelector("div#pdp_product_name", { timeout: 15000 }).catch(() => console.log(`Name element not found for ${products[i].link}`));

        // Expand "More Details" for description and specifications
        await detailPage.evaluate(() => {
          const buttons = document.querySelectorAll("button.pdp-more-details");
          buttons.forEach((button) => button.click());
        });
        await delay(1500);

        const productDetails = await detailPage.evaluate(() => {
          const nameEl = document.querySelector("div#pdp_product_name");
          const brandEl = document.querySelector("a#top_brand_name");
          const priceEl = document.querySelector("div#price_section span.jm-heading-xs");
          const mrpEl = document.querySelector("div#price_section span.line-through");
          const offerEl = document.querySelector("div#price_section span.jm-badge");
          const sellerEl = document.querySelector("section#buybox_soldby_container h2.jm-body-m-bold.jm-fc-primary-60");
          const descriptionEl = document.querySelector("div#pdp_description");
          const imagesEls = document.querySelectorAll("img.swiper-thumb-slides-img, img.largeimage.swiper-slide-img");

          // Extract specifications
          const info = {};
          document.querySelectorAll("table.product-specifications-table tbody tr").forEach((row) => {
            const key = row.querySelector("th")?.textContent.trim();
            const value = row.querySelector("td")?.textContent.trim();
            if (key && value) info[key] = value;
          });

          // Extract images
          const images = Array.from(imagesEls)
            .map((img) => img.src)
            .filter((src) => src.includes("jiomart.com/images/product/original"))
            .join("; ") || "N/A";

          console.log(`PDP - Name: ${nameEl ? nameEl.textContent.trim() : "Not found"}`);
          console.log(`PDP - Brand: ${brandEl ? brandEl.textContent.trim() : "Not found"}`);
          console.log(`PDP - Price: ${priceEl ? priceEl.textContent.trim() : "Not found"}`);
          console.log(`PDP - MRP: ${mrpEl ? mrpEl.textContent.trim() : "Not found"}`);
          console.log(`PDP - Offer: ${offerEl ? offerEl.textContent.trim() : "Not found"}`);
          console.log(`PDP - Seller: ${sellerEl ? sellerEl.textContent.trim() : "Not found"}`);
          console.log(`PDP - Description: ${descriptionEl ? descriptionEl.textContent.trim() : "Not found"}`);
          console.log(`PDP - Images: ${images}`);

          return {
            name: nameEl ? nameEl.textContent.trim() : "N/A",
            brand: brandEl ? brandEl.textContent.trim() : "N/A",
            price: priceEl ? priceEl.textContent.trim() : "N/A",
            mrp: mrpEl ? mrpEl.textContent.trim() : "N/A",
            offer: offerEl ? offerEl.textContent.trim() : "N/A",
            seller: sellerEl ? sellerEl.textContent.trim() : "N/A",
            description: descriptionEl ? descriptionEl.textContent.trim() : "N/A",
            images,
            info,
          };
        });

        products[i] = {
          name: productDetails.name,
          brand: productDetails.brand,
          price: productDetails.price,
          mrp: productDetails.mrp,
          offer: productDetails.offer,
          seller: productDetails.seller,
          description: productDetails.description,
          images: productDetails.images,
          image: products[i].image,
          link: products[i].link,
          ...productDetails.info,
        };

        console.log(`Extracted details for ${products[i].name}:`, products[i]);
        allProductDetails.push(products[i]);
        success = true;

        // Save immediately
        try {
          await saveToExcel([products[i]], "jiomart_products.xlsx", !isFirstSave);
          isFirstSave = false;
        } catch (error) {
          console.error(`Failed to save product ${products[i].name}: ${error.message}`);
        }

        await detailPage.close();
      } catch (error) {
        console.error(`Attempt ${4 - retries} failed for ${products[i].name}: ${error.message}`);
        retries--;
        if (retries === 0) {
          console.error(`All retries failed for ${products[i].name}. Saving partial data.`);
          allProductDetails.push(products[i]);
          try {
            await saveToExcel([products[i]], "jiomart_products.xlsx", !isFirstSave);
            isFirstSave = false;
          } catch (saveError) {
            console.error(`Failed to save partial data for ${products[i].name}: ${saveError.message}`);
          }
        }
        if (detailPage) await detailPage.close();
        await delay(2000);
      }
    }
  }

  console.log("All product details extracted. Performing final save...");
  try {
    await saveToExcel(allProductDetails, "jiomart_products.xlsx", false);
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
      await delay(2000);
    }
  }
}

scrapeJioMart().catch(async (error) => {
  console.error("Error in scrapeJioMart:", error);
  process.exit(1);
});