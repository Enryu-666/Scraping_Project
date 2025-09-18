// main scraping script
const { Builder, By, Key, until } = require("selenium-webdriver");
// Import the saveToExcel function from your new file
const { saveToExcel } = require("./excelWriter.js");

async function WebScrapper() {
    let driver = await new Builder().forBrowser("chrome").build();
    let reviews = [];

    try {
        await driver.get("https://www.tokopedia.com/samsung-official-store/");
        const ulasan = await driver.wait(until.elementLocated(By.css('button[data-testid="Ulasan"]')));
        await ulasan.click();

        let page = 1;
        const max = 2;

        while (page <= max) {
            await driver.wait(until.elementLocated(By.css('article.css-1pr2lii')));
            const ulasanContainer = await driver.findElements(By.css('article.css-1pr2lii'));

            for (let elem of ulasanContainer) {
                let username = '';
                let rating = '';
                let comment = '';

                try {
                    username = await (await elem.findElement(By.css('div.css-k4rf3m'))).getText();
                } catch { }

                try {
                    rating = await (await elem.findElement(By.css('div[data-testid="icnStarRating"]'))).getAttribute('aria-label');
                } catch { }

                try {
                    comment = await (await elem.findElement(By.css('span[data-testid="lblItemUlasan"]'))).getText();
                } catch { }

                reviews.push({ username, rating, comment, page });
            }
            
            console.log(`Scraped page ${page}...`);

            let nextBtn;
            try {
                nextBtn = await driver.findElement(By.css('button[aria-label="Laman berikutnya"]'));
                const isDisabled = await nextBtn.getAttribute('disabled');
                if (isDisabled !== null) break;
                await nextBtn.click();
                page++;
                await driver.sleep(5000);
            } catch {
                break;
            }
        }
        
        if (reviews.length > 0) {
            // Call the imported function with the data, filename, and an optional directory
            const outputDirectory = 'scraped_data_output'; // Define your output folder
            saveToExcel(reviews, 'ulasan_tokopedia.xlsx', outputDirectory);
        } else {
            console.log('No reviews were scraped.');
        }

    } catch (error) {
        console.error('Error occurred:', error);
    } finally {
        await driver.quit();
    }
} 

WebScrapper();