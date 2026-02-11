const { Builder, By, until } = require('selenium-webdriver');
const xlsx = require('sheetjs');
const fs = require('fs');
const path = require('path');
require('dotenv').config();

const config = {
    env: {
        URL_LOGIN: process.env.URL_LOGIN,
        USER_EMAIL: process.env.USER_EMAIL,
        USER_PASSWORD: process.env.USER_PASSWORD,
        URL_ABSENSI: process.env.URL_ABSENSI
    }
};

async function runAutomation() {
    const workbook = xlsx.readFile('Data Santri Lulu Januari 2026.ods');
    const dataSantri = xlsx.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
    
    let driver = await new Builder().forBrowser('firefox').build();
    await driver.manage().window().setRect({ width: 1366, height: 768 });

    try {
        await driver.get(config.env.URL_LOGIN);
        await driver.findElement(By.id('email')).sendKeys(config.env.USER_EMAIL);
        await driver.findElement(By.id('password')).sendKeys(config.env.USER_PASSWORD);
        await driver.findElement(By.xpath("/html/body/div[1]/div/div/div/div/form/div[4]/button")).click();
        await driver.wait(until.urlContains('dashboard'), 15000);
        await driver.sleep(5000);
        await driver.findElement(By.xpath("//div[contains(text(), 'Transaksi')]"));
        await driver.findElement(By.xpath("/html/body/div[1]/div[1]/aside/ul/li[4]/a")).click();
        await driver.wait(until.elementLocated(By.xpath("//div[@data-i18n='Kelola Absensi Santri']")), 5000);
        await driver.findElement(By.xpath("//div[@data-i18n='Kelola Absensi Santri']")).click();
        await driver.wait(until.urlContains('absen_santri'), 10000);

        for (let santri of dataSantri) {
            const namaSiswa = santri["Santri"] ? santri["Santri"].trim() : null;
            
            if (!namaSiswa) continue;

            const namaFile = (santri["Santri"] || "unknown").replace(/[^a-zA-Z0-9]/g, "_");
            console.log(`\nMemproses Santri: ${namaSiswa}`);

            // LOGIKA PENYARINGAN NAMA
            const namaLow = namaSiswa.toLowerCase();
            if (
                namaLow.includes("fabio") || 
                namaLow.includes("ali akbar alfatih") || 
                namaLow.includes("aisyah fazilla")
            ) {
                await driver.get(config.env.URL_ABSENSI);
                await driver.wait(until.urlContains('absen_santri'), 10000);
                const xpathBaris = `//tr[contains(normalize-space(.), "${namaSiswa}")]//a[contains(@class, 'btn-warning')]`;
                let btnPresensi = await driver.wait(until.elementLocated(By.xpath(xpathBaris)), 5000);
                await btnPresensi.click();
                await driver.sleep(5000);
                await driver.findElement(By.xpath('/html/body/div/div[1]/div/div/div[1]/div/div[2]/div/div/form/div[4]/a[2]')).click();
                await driver.sleep(5000);
                await takeScreenshot(driver, `Santri_${namaFile}.png`);
            }
            else {
                // Jalur Normal untuk Santri Lain
                await driver.get(config.env.URL_ABSENSI);
                
                // Cari baris kelas di tabel utama
                const xpathBaris = `//tr[contains(., "${namaSiswa}")]//a[contains(@class, 'btn-warning')]`;
                let btnPresensi = await driver.wait(until.elementLocated(By.xpath(xpathBaris)), 10000);
                await btnPresensi.click();

                const xpathCapaianBulan01 = `//h5[contains(., "${namaSiswa}")]/parent::div/following-sibling::div//a[contains(., "01")]`;
                let btnCapaian = await driver.wait(until.elementLocated(By.xpath(xpathCapaianBulan01)), 10000);
                await driver.executeScript("arguments[0].click();", btnCapaian);
                await driver.sleep(5000);
                await takeScreenshot(driver, `Santri_${namaFile}.png`);
            }

            // 3. ISI FORM & SIMPAN
            // try {
            //     await isiFormCapaian(driver, santri);
            //     console.log(`Berhasil input data: ${namaSiswa}`);
            // } catch (err) {
            //     console.error(`Gagal simpan data ${namaSiswa}: ${err.message}`);
            // }
        }
    } finally {
        console.log("\nSemua proses selesai.");
        await driver.quit();
    }
}

// Fungsi helper dipindahkan ke luar agar lebih rapi dan efisien
async function takeScreenshot(driver, filename) {
    // Menggunakan __dirname agar folder dibuat relatif terhadap lokasi file script ini
    const screenshotDir = path.join(__dirname, 'Screenshot');
    
    // Memastikan direktori ada, jika tidak, maka akan dibuat
    if (!fs.existsSync(screenshotDir)) {
        fs.mkdirSync(screenshotDir, { recursive: true });
    }
    let image = await driver.takeScreenshot();
    const filePath = path.join(screenshotDir, filename);
    fs.writeFileSync(filePath, image, 'base64');
    console.log(`Screenshot disimpan: ${filePath}`);
}

runAutomation();