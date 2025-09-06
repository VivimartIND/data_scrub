const puppeteer = require("puppeteer");
const ExcelJS = require("exceljs");

const url =
  "https://www.zeptonow.com/cn/bath-body/shower-gels/cid/26e64367-19ad-4f80-a763-42599d4215ee/scid/e8801cfd-49d0-4388-bc29-f83258204d39";

async function scrapeProductData() {
  console.log("Launching browser...");
  const browser = await puppeteer.launch({
    headless: false,
    slowMo: 100,
    defaultViewport: { width: 1280, height: 800 },
  });

  const page = await browser.newPage();
  console.log("Navigating to URL...");
  await page.goto(url, { waitUntil: "networkidle2" });

  console.log("Scrolling to load all products...");
  let previousHeight = 0;
  while (true) {
    await page.evaluate(() => window.scrollBy(0, document.body.scrollHeight));
    await new Promise((resolve) => setTimeout(resolve, 2000));
    const newHeight = await page.evaluate(() => document.body.scrollHeight);
    if (newHeight === previousHeight) break;
    previousHeight = newHeight;
  }

  console.log("Extracting product details from listing page...");
  const products = await page.evaluate(() => {
    return Array.from(document.querySelectorAll('a[href^="/pn/"]')).map((a) => {
      const href = a.getAttribute("href");
      return {
        name: a.querySelector("img")?.alt || "N/A",
        link: "https://www.zeptonow.com" + href,
        image: a.querySelector("img")?.src || "N/A",
      };
    });
  });

  console.log(
    `Found ${products.length} products. Now fetching details for each product...`
  );

  const allKeys = new Set(); // collect all possible info fields dynamically

  for (let i = 0; i < products.length; i++) {
    console.log(`Opening product page for: ${products[i].name}`);
    let productPage;
    try {
      productPage = await browser.newPage();
      await retryGoto(productPage, products[i].link);

      const productDetails = await productPage.evaluate(() => {
        const name =
          document.querySelector("h1.font-semibold")?.textContent.trim() ||
          "N/A";

        // Price & MRP
        const priceElement = Array.from(document.querySelectorAll("p")).find(
          (p) => p.textContent.includes("₹") && !p.textContent.includes("line-through")
        );
        const price = priceElement?.textContent.match(/₹\d+/)?.[0] || "N/A";
        const mrpElement = document.querySelector("span.line-through");
        const mrp = mrpElement?.textContent.replace("₹", "").trim() || "N/A";

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
        priceNum && mrpNum
          ? (((mrpNum - priceNum) / mrpNum) * 100).toFixed(2) + "%"
          : "N/A";

      // Merge everything
      products[i] = {
        name: productDetails.name,
        link: products[i].link,
        image: products[i].image,
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
    } catch (error) {
      console.error(
        `Error while scraping product page for ${products[i].name}:`,
        error
      );
    } finally {
      if (productPage) await productPage.close();
    }
  }

  console.log("All product details extracted. Saving to Excel...");

  // Create Excel workbook
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Products");

  // Basic columns
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

  await workbook.xlsx.writeFile("zepto_products.xlsx");
  console.log("Data saved to zepto_products.xlsx ✅");

  await browser.close();
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