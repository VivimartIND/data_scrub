const { chromium } = require('playwright');

async function main() {
  // Launch browser in non-headless mode for debugging (set to true for production)
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({
    userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
    viewport: { width: 1280, height: 720 },
  });
  const page = await context.newPage();

  try {
    // Step 1: Scrape listing page
    const listingUrl = 'https://blinkit.com/cn/vegetables/cid/294/296'; // Adjust to actual listing page URL
    const products = await scrapeListingPage(page, listingUrl);
    console.log('Products from listing page:', products);

    // Step 2: Scrape individual product pages (optional)
    for (const product of products.slice(0, 2)) { // Limit to 2 products for testing
      const productData = await scrapeProductPage(page, product.name, product.productId, product.slug);
      if (productData) {
        console.log(`Final data for ${product.name}:`, { ...product, ...productData });
      }
    }
  } catch (error) {
    console.error('Main error:', error);
    await page.screenshot({ path: `error-main-${Date.now()}.png` });
  } finally {
    await browser.close();
  }
}

async function scrapeListingPage(page, url) {
  console.log(`Navigating to listing page: ${url}`);
  await page.goto(url, { waitUntil: 'networkidle', timeout: 120000 });

  const products = await page.evaluate(() => {
    const cards = document.querySelectorAll('.tw-relative.tw-flex.tw-h-full.tw-flex-col');
    return Array.from(cards).map(card => {
      const name = card.querySelector('.tw-text-300.tw-font-semibold.tw-line-clamp-2')?.innerText || 'N/A';
      const price = card.querySelector('.tw-text-200.tw-font-semibold')?.innerText || 'N/A';
      const pack = card.querySelector('.tw-text-200.tw-font-medium.tw-line-clamp-1')?.innerText || 'N/A';
      const image = card.querySelector('img')?.src || 'N/A';
      const productId = card.id || 'N/A';
      const deliveryTime = card.querySelector('.tw-text-050.tw-font-bold.tw-uppercase')?.innerText || 'N/A';
      const slug = card.querySelector('.tw-text-300.tw-font-semibold.tw-line-clamp-2')?.innerText
        ?.toLowerCase()
        .replace(/[^a-z0-9]+/g, '-') || 'N/A';
      return { name, price, pack, image, productId, deliveryTime, slug };
    });
  });

  return products;
}

async function scrapeProductPage(page, productName, productId, productSlug) {
  const url = `https://blinkit.com/prn/${productSlug}/prid/${productId}`;
  const maxRetries = 3;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      console.log(`Attempt ${attempt} for ${productName} at ${url}`);
      await page.goto(url, { waitUntil: 'networkidle', timeout: 120000 });

      // Wait for a key element to ensure page is loaded
      await page.waitForSelector('.product-name', { timeout: 30000 }); // Adjust selector as needed

      const productData = await page.evaluate(() => {
        const name = document.querySelector('.product-name')?.innerText || 'N/A';
        const price = document.querySelector('.product-price')?.innerText || 'N/A';
        const description = document.querySelector('.product-description')?.innerText || 'N/A';
        const images = Array.from(document.querySelectorAll('.product-images img')).map(img => img.src) || ['N/A'];
        const unit = document.querySelector('.product-unit')?.innerText || 'N/A';
        const healthBenefits = document.querySelector('.health-benefits')?.innerText || 'N/A';
        // Add other fields as needed
        return { name, price, description, images, unit, healthBenefits };
      });

      console.log(`Extracted details for ${productName}:`, productData);
      return productData;
    } catch (error) {
      console.log(`Attempt ${attempt} failed for ${productName}: ${error}`);
      if (attempt === maxRetries) {
        console.log(`All retries failed for ${productName}. Skipping.`);
        await page.screenshot({ path: `error-${productName}-${Date.now()}.png` });
        return null;
      }
      await page.waitForTimeout(2000); // Wait 2 seconds before retrying
      await page.context().clearCookies(); // Clear cookies to reset session
    }
  }
}

main().catch(console.error);