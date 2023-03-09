const Excel = require("exceljs");
const puppeteer = require("puppeteer");
const cheerio = require("cheerio");
const workbook = new Excel.Workbook();
const _progress = require("cli-progress");

const urls = [];

const keywords = [
  "Gri Kırlent Kılıfı",
  "Yeşil Kırlent Kılıfı",
  // "Sarı Kırlent Kılıfı",
  // "Siyah Beyaz Kırlent Kılıfı",
  // "Mavi Kırlent Kılıfı",
  // "Siyah Kırlent Kılıfı ",
  // "Beyaz Kırlent Kılıfı",
  // "Turuncu Kırlent Kılıfı",
  // "Lacivert Kırlent Kılıfı",
  // "Pembe Kırlent Kılıfı",
  // "Dekoratif Kırlent Kılıf",
  // "Peluş Kırlent Kılıfı",
  // "Örgü Kırlent Kılıfı",
  // "Desenli Kırlent Kılıfı",
  // "Düz Renk Kırlent Kılıfı",
  // "Kadife Kırlent Kılıfı",
  // "Dikdörtgen Kırlent Kılıfı",
  // "Peluş Koltuk Şalı",
  // "Sarı Koltuk Şalı",
  // "Beyaz Koltuk Şalı",
  // "Turuncu Koltuk Şalı",
  // "Hardal Koltuk Şalı",
  // "Mavi Koltuk Şalı",
  // "Siyah Masa Örtüsü",
  // "Beyaz Masa Örtüsü",
  // "Kırmızı Masa Örtüsü",
  // "Gri Masa Örtüsü",
  // "Kahverengi Masa Örtüsü",
  // "Mavi Masa Örtüsü",
  // "Pembe Masa Örtüsü",
  // "Yeşil Masa Örtüsü",
  // "Lacivert Masa Örtüsü",
  // "Simli Masa Örtüsü",
  // "Pullu Masa Örtüsü",
  // "Mor Masa Örtüsü",
  // "Gümüş Masa Örtüsü",
  // "Gold Masa Örtüsü",
  // "Pötikare Masa Örtüsü",
  // "Ekose Masa Örtüsü",
  // "Yemek Masa Örtüsü",
  // "Siyah peçete",
  // "Yeşil Peçete",
  // "Beyaz Peçete",
  // "Gri Peçete",
  // "Keten Peçete",
  // "Gold Runner masa bandı",
  // "Gri Runner masa bandı",
  // "runner masa bandı peteçe seti",
  // "Siyah Sandalye Minderi",
  // "Gri Sandalye Minderi",
  // "Arkalıklı Sandalye Minderi",
  // "Gri Yatak Örtüsü",
  // "Siyah Yatak Örtüsü",
  // "Beyaz Yatak Örtüsü",
  // "Pudra Yatak Örtüsü",
  // "Tek Kişilik Yatak Örtüsü",
  // "Çift Kişilik Yatak Örtüsü",
  // "Yatak Örtüsü Seti",
  // "Çeyizlik Yatak Örtüsü",
  // "Kapitone Yatak Örtüsü",
  // "Çocuk Yatak Örtüsü",
  // "Pike Yatak Örtüsü",
  // "Gelin Yatak Örtüsü",
  // "Günlük Yatak Örtüsü",
  // "Örgü Yatak Örtüsü",
  // "Saten Yatak Örtüsü",
];

const tableDIR = "tablolar/trendyol/trendyolv2.xlsx";
const productPerKeyword = 2

keywords.map((k) => {
  const slugifyKeyword = k.replaceAll(" ", "+");
  urls.push({
    name: k,
    keywordSearch: `https://www.trendyol.com/sr?q=${slugifyKeyword}&sst=MOST_RATED`,
    links: [],
  });
  console.log(urls)
});
console.log("Ürünler Taranıyor... \n");
const b1 = new _progress.Bar({}, _progress.Presets.shades_grey);
b1.start(keywords.length, 0);
// the bar value - will be linear incremented
let value = 0;

(async () => {
  const browser = await puppeteer.launch({ headless: true });
  const page = await browser.newPage();

  let currentUrlIndex = 0;
  const openNextUrl = async () => {
    console.log('Link elements: ' + linkElements[i])
    if (currentUrlIndex >= urls.length) {
      console.log(urls)
      return;
    }
    const url = urls[currentUrlIndex].keywordSearch;
    await page.goto(url, {
      waitUntil: "domcontentloaded",
      timeout: 0,
    });
    await page.goto(url);
    const linkElements = await page.$$(".voltran-product-list a");
    for (let i = 0; i < productPerKeyword; i++) {
      if (linkElements[i]) {
        urls[currentUrlIndex].links.push(
          await (await linkElements[i].getProperty("href")).jsonValue()
        );
      }
    }

    value++;
    b1.update(value);

    currentUrlIndex++;
    await openNextUrl();
  };
  await openNextUrl();

  let currentVal;

  //Ürünler burada başlıyor
  if (currentUrlIndex === urls.length) {
    for (const cat in urls) {
      const tableHeaders = [
        { header: "Ürün Linki", key: "url" },
        { header: "Adı", key: "ad" },
        { header: "Satış Fiyatı", key: "price" },
        { header: "Değerlendirilme Sayısı", key: "reviews" },
        { header: "Değerlendirme Puanı", key: "score" },
        { header: "Renk", key: "color" },
        { header: "Ebat", key: "ebat" },
        { header: "Ürün Görsel Linki", key: "image" },
      ];

      const sheet = workbook.addWorksheet(urls[cat].name);

      sheet.columns = tableHeaders;

      let currentKeywordIndex = 0;

      if(!currentVal){
        currentVal = keywords.length + urls[cat].links.length
      } else {
        currentVal = currentVal + urls[cat].links.length
      }

      b1.setTotal(currentVal);

      const openKeywordUrl = async () => {
        if (cat >= urls.length - 1 && currentKeywordIndex >= urls[cat].links.length) {
          console.log('\n\nTarama işlemi tamamlandı.')
          console.log(`Taranan ürünler ${tableDIR} excel tablosuna kaydedildi.`)
          console.log(urls)
          await browser.close();
          process.exit();
        }
        if (currentKeywordIndex >= urls[cat].links.length) {
          return;
        }
        const url = urls[cat].links[currentKeywordIndex];
        await page.goto(url, {
          waitUntil: "domcontentloaded",
          timeout: 0,
        });
        await page.goto(url);
        const content = await page.content();
        const $ = cheerio.load(content);
        const ad = $("#product-name").text();
        const price = $("#originalPrice").text();
        const reviews = $("#comments-container > a > span").text();
        const score = $("#productReviews > span.rating-star").text();
        const color = $('label[data-propertyname="Renk"].checked .variant-name').text();
        const ebat = $('label[data-propertyname="Ebatlar"].checked').text();
        const image = $("#productDetailsCarousel > div.owl-stage-outer > div > div.owl-item.active > a > picture > img").attr("src");
        sheet.addRow({ url: url || "-", ad: ad.trim() || "-", price: price || "-", reviews: reviews || "-", score: score || "-", color: color || "-", ebat: ebat || "-", image: image || "-" });
        workbook.xlsx
          .writeFile(`${tableDIR}`)
          .then(() => {})
          .catch((error) => {
            console.log(error.message);
          });
        currentKeywordIndex++;

        value++;
        // update the bar value
        b1.update(value);
        await openKeywordUrl();
      };
      await openKeywordUrl();
    }
  }
})();
