const puppeteer = require("puppeteer");
const fs = require("fs/promises");
const Excel = require("excel4node");

(async () => {
  let browser;

  try {
    browser = await puppeteer.launch({ headless: false,  protocolTimeout: 60000, });
    const page = await browser.newPage();

    const codigos = await fs.readFile("expedientes.txt", "utf-8");
    const codigosArray = codigos
      .split("\n")
      .filter((linea) => linea.trim() !== "");
    const codigosLimpios = codigosArray.map((codigo) => codigo.trim());

    let objeto = {};
    try {
      const jsonData = await fs.readFile("resultados.json", "utf-8");
      objeto = JSON.parse(jsonData);
    } catch (error) {
      console.log(
        "No se encontraron datos previos en el archivo JSON. Se crearán nuevos datos."
      );
    }

    for (const codigo of codigosLimpios) {
      await page.goto("https://www.anses.gob.ar/consultas/fecha-de-cobro");

      try {
        await page.waitForSelector("#edit-nro-cuil", { timeout: 5000 });
      } catch (error) {
        console.error(
          `El selector '#edit-nro-cuil' no se encontró en la página. Pasando al siguiente expediente.`
        );
        continue;
      }

      await page.type("#edit-nro-cuil", codigo);
      await page.click("#edit-submit");

      await new Promise((resolve) => setTimeout(resolve, 5000));

      try {
        await page.waitForSelector("#content", { timeout: 5000 });

        const nombreCompleto = await page.$eval("div.person h2", (element) =>
          element.textContent.trim()
        );
        const beneficio = await page.$eval("div.benefit h3", (element) =>
          element.textContent.trim()
        );

        const fechaCobro = await page.$eval("div.date", (element) =>
          element.textContent
            .trim()
            .replace(/\s+/g, " ")
            .replace(/\n+/g, "")
            .replace(/(.+)\s+(\w+)\s+(\d+)/, "$2 $3 $1")
        );

        if (objeto[codigo]) {
          objeto[codigo] = {
            ...objeto[codigo],
            nombre: nombreCompleto,
            Cuil: codigo,
            Beneficio: beneficio,
            Fecha_Cobro: fechaCobro,
          };
        } else {
          objeto[codigo] = {
            Nombre: nombreCompleto,
            Cuil: codigo,
            Beneficio: beneficio,
            Fecha_Cobro: fechaCobro,
          };
        }
      } catch (error) {
        console.error(` este Expediente no tiene tramite iniciado ${codigo} `);
      }
    }
    await fs.writeFile("resultados.json", JSON.stringify(objeto, null, 2));
    const jsonData = await fs.readFile("resultados.json", "utf-8");
    const data = JSON.parse(jsonData);
    //crear libro excel
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("Datos");

    const headers = Object.keys(objeto[Object.keys(objeto)[0]]);
    headers.forEach((header, index) => {
      worksheet.cell(1, index + 1).string(header);
    });

    Object.values(objeto).forEach((row, rowIndex) => {
      headers.forEach((header, colIndex) => {
        worksheet.cell(rowIndex + 2, colIndex + 1).string(String(row[header]));
      });
    });

    await workbook.write("datos.xlsx");
    console.log("Archivo Excel guardado correctamente.");
  } catch (error) {
    console.error("Error:", error);
  } finally {
    if (browser) {
      await browser.close();
    }
  }
})();
