const puppeteer = require("puppeteer");
var fs = require("fs");
var writeStream = fs.createWriteStream("natega.xls");

var header =
  "Name" +
  "\t" +
  "Setting Number" +
  "\t" +
  "Status" +
  "\t" +
  "Total" +
  "\t" +
  "Percent" +
  "\n";

writeStream.write(header);

const totalGrades = 410;

(async () => {

  if (
    process.argv[2] == null ||
    process.argv[3] == null ||
    isNaN(process.argv[2]) ||
    isNaN(process.argv[3])
  ) {

    console.log("Please enter arguments correctly.");
    process.exit();
    
  }

  //min nearly 101000
  //max nearly 530100
  const setting_number_begin = process.argv[2];
  const setting_number_end = process.argv[3];

  const browser = await puppeteer.launch({ headless: true });
  const page = await browser.newPage();
  await page.goto("http://natega.thanwya.emis.gov.eg/");

  for (var i = setting_number_begin; i < setting_number_end; i++) {
    await page.focus("#TextBox1");
    await page.keyboard.down("Control");
    await page.keyboard.press("A");
    await page.keyboard.up("Control");
    await page.keyboard.press("Backspace");
    await page.type("#TextBox1", i.toString());
    await page.click("#Button3", { waitUntil: "domcontentloaded" });
    await page.waitForSelector("#Table2");

    let data = await page.evaluate(() => {
      if (document.querySelector("span[id=std_name]") != null) {
        let name = document.querySelector("span[id=std_name]").innerHTML;
        let setting_num = document.querySelector("span[id=seating_no]")
          .innerHTML;
        let status = document.querySelector("span[id=Label7]").innerHTML;
        let total = document.querySelector("span[id=total]").innerHTML;
        return {
          name,
          setting_num,
          status,
          total
        };
      } else {
        return null;
      }
    });

    if (data != null) {
      var row =
        data.name +
        "\t" +
        data.setting_num +
        "\t" +
        data.status +
        "\t" +
        isNaN(data.total)
          ? "\n"
          : data.total + "\t" + (data.total / totalGrades) * 100 + "%" + "\n";

      await writeStream.write(row);
    }
  }

  await browser.close();
})();
