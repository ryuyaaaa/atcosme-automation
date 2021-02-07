const fs = require('fs');
const puppeteer = require('puppeteer');
const officegen = require('officegen');
const TARGET_URL = 'https://www.cosme.net/product/product_id/10037217/top';
const EXCEL_SHEET_NAME = 'ReFa CARAT';

// const start_time = performance.now();

var xlsx = officegen('xlsx');

xlsx.on('finalize', function (written) {
    console.log('エクセルファイルへの書き込みが終了しました。');
});

xlsx.on('error', function (err) {
    console.log(err);
});

let sheet = xlsx.makeNewSheet();
sheet.name = EXCEL_SHEET_NAME;

sheet.data[0] = [];
sheet.data[0][0] = '投稿者名';
sheet.data[0][1] = 'クチコミ';

let col = 0;

(async () => {
    try {
        const browser = await puppeteer.launch({
            args: [
                '--disable-gpu',
                '--disable-dev-shm-usage',
                '--disable-setuid-sandbox',
                '--no-first-run',
                '--no-sandbox',
                '--no-zygote',
                '--single-process'
            ]
        });

        const page = await browser.newPage();
        await page.setRequestInterception(true);

        page.on('request', request => {
            if (request.url().includes('cosme.net/')) {
                request.continue().catch(err => console.error(err));
            } else {
                request.abort().catch(err => console.error(err));
            }
        })
        page.on('console', msg => {
            for (let i = 0; i < msg._args.length; ++i) console.log(`${i}: ${msg._args[i]}`);
        });
        await page.goto(TARGET_URL);

        let sum_review_count = -1;
        let review_count = 0;

        setInterval(function () {
            process.stdout.write(`取得中... ${review_count}/${sum_review_count}\r`);
            if (review_count === sum_review_count) process.exit();
        }, 500);

        // クチコミの数
        let review_count_selector = 'ul.rev-cnt li a span.count.cnt';
        await page.waitForSelector(review_count_selector);
        sum_review_count = parseInt(await page.$$eval(review_count_selector, selector => selector[0].innerText), 10);
        console.log(sum_review_count + '件のクチコミの取得を開始します。');

        // クチコミX件をクリック
        let review_button_selector = 'ul.rev-cnt li a';
        await page.waitForSelector(review_button_selector);
        await page.$$eval(review_button_selector, selector => {
            selector[0].click();
        });

        let review_list_item_read_more_selector = '#product-review-list .review-sec div.inner div.body p.read span a';
        await page.waitForSelector(review_list_item_read_more_selector);
        let read_mores = await page.$$('#product-review-list .review-sec div.inner div.body p.read span a');
        await read_mores[0].click();

        let review_indivi_selector = 'div.inner div.body p.read';
        await page.waitForSelector(review_indivi_selector);

        let pre_reviewer_name = null;
        while (true) {
            let retry_count = 0;
            let reviewer_name_selector = 'span.reviewer-name';
            await page.waitForSelector(reviewer_name_selector);
            while (true) {
                let current_reviewer_name = await page.$$eval(reviewer_name_selector, selector => selector[0].innerText);
                if (!!current_reviewer_name && current_reviewer_name != pre_reviewer_name) {
                    col++;
                    pre_reviewer_name = current_reviewer_name;
                    sheet.data[col] = [];
                    sheet.data[col][0] = current_reviewer_name;
                    break;
                } else {
                    await page.waitFor(300);
                    retry_count++;
                    if (retry_count >= 10) {
                        console.log('エラー発生');
                        await browser.close();
                        return;
                    }
                }
            }

            let review_content = await page.$$eval(review_indivi_selector, selector => selector[0].innerText);

            review_count++;
            sheet.data[col][1] = review_content;
            // console.log('.');

            try {
                let next_button_selector = 'ul li.next a';
                await page.waitForSelector(next_button_selector, { timeout: 5000 });
                let next_buttons = await page.$$(next_button_selector);
                next_buttons[0].click();
                await page.waitForNavigation();
            } catch (err) {
                break;
            }
        }

        console.log(review_count + '個のクチコミを取得しました。')
        let out = fs.createWriteStream('ReFa_CARAT.xlsx');

        out.on('error', function (err) {
            console.log(err);
        });

        xlsx.generate(out);
        await browser.close();

    } catch (err) {
        console.log(err);
    }
})();

// const end_time = performance.now();
// console.log(`実行時間：${Math.floor((end_time - start_time / 1000))}秒`);