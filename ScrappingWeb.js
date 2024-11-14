const axios = require('axios');
const cheerio = require('cheerio');
const ExcelJS = require('exceljs');

//1- Función Main.
async function main() {
    const paginas = await leerArchivoExcel("entrada.xlsx"); //Lectura del archivo de entrada.
    const datosScrappeados = []; //Constante donde se guardarán todos los datos que se vayan a escribir en el archivo de salida.

    for (const pagina of paginas) { //Se recorre cada página cargada en el array de paginas.
        try {
            const datos = await scrapearPagina(pagina.url); //Se llama a la función "scrapearPagina" y se guardan los resultados del scrapping en la constante "datos".
            if (datos.length > 0) {
                datosScrappeados.push({ //Se cargan los datos scrapeados en "DatosScrappeados".
                    nombreDePagina: pagina.nombreDePagina,
                    url: pagina.url,
                    datos
                });
            }
        } catch (error) {
            console.log(`Error al procesar la página ${pagina.url}`); //Control de errores try-catch.
        }
    }

    await escribirArchivoExcel("salida.xlsx", datosScrappeados); //Se llama a la función "escribirArchivoExcel" y se le otorga los datos a escribir.
}

//2- Función de lectura del archivo de entrada.
async function leerArchivoExcel(ruta) {
    const workbook = new ExcelJS.Workbook(); //Creacion del "libro de trabajo"
    await workbook.xlsx.readFile(ruta); //Lectura del archivo de entrada
    const sheet = workbook.getWorksheet(1); //Obtención de datos de la primera hoja

    const paginas = [];

    sheet.eachRow((filas, numeroDeFila) => {
        if (numeroDeFila > 1) {  //Salteamos el encabezado
            const nombreDePagina = filas.getCell(1).text; //Se obtiene el dato de la celda 1
            const url = filas.getCell(2).text; //Se obtiene el dato de la celda 2
            paginas.push({ nombreDePagina, url }); //Se guardan los datos en el array
        }
    });
    return paginas;
}

//3- Función de Scrapping.
async function scrapearPagina(url) {
    const { data } = await axios.get(url); //Se obtienen los datos DOM de la página utilizando Axios.
    const $ = cheerio.load(data); //Se carga el DOM utilizando Cheerio.
    let resultados = [];

    //Extraemos los datos del DOM según pagina web.
    if (url.includes("stackoverflow")) { //Scrapeo para la página StackOverflow.
        $('text[aria-label="Response"]').each((_, element) => {

            const lenguaje = $(element).text().trim(); //Extracción del nombre del lenguaje.
            const porcentaje = $(element).next('text[aria-label="Unit"]').text().trim(); //Extracción del porcentaje del lenguaje.

            if (lenguaje && porcentaje)  //Mientras que ninguna de las dos constantes esté vacia...
                resultados.push(lenguaje + '|' + porcentaje); //Se guardan los resultados de la extracción.
        });
    } else if (url.includes("browserstack")) { //Scrapeo para la página BrowserStack.
        $('h4').each((_, element) => {

            const datos = $(element).text().trim(); //Extracción del nombre y posicion del lenguaje.
            const [posicion, lenguaje] = datos.split("."); //División del nombre y posicion del lenguaje.

            if (lenguaje && posicion)
                resultados.push(lenguaje + '|' + posicion);
        });
    } else if (url.includes("crossover")) {  //Scrapeo de la página crossover
        $('li.list-item').each((_, element) => {

            const datos = $(element).text().trim();
            const [posicion, lenguaje] = datos.split(".");

            if (lenguaje && posicion)
                resultados.push(lenguaje + '|' + posicion);
        });
    } else {
        $('h3').each((_, element) => { //Scrapeo de las páginas intelivita, eluminoustechnologies y hackr
            let datos;
            if (url.includes("hackr"))
                datos = $(element).find('strong').text().trim(); //Extracción de datos en Hackr.
            else
                datos = $(element).text().trim(); //Extracción de datos en intelivita y eluminoustechnologies.

            const [posicion, lenguaje] = datos.split(".");

            if (lenguaje && posicion) {
                resultados.push(lenguaje + '|' + posicion);
            }
        });
    }
    return resultados.slice(0, 10); //Solo devuelve los primeros 10 elementos.
}

//4- Función de escritura de los datos obtenidos.
async function escribirArchivoExcel(ruta, bloquesDeDatos) {
    const libroDeTrabajo = new ExcelJS.Workbook(); //Creacion del "libro de trabajo"
    const hoja = libroDeTrabajo.addWorksheet('Datos Scrappeados'); //Se crea una hoja llamada "Datos Scrappeados"
    hoja.addRow(['Nombre de Página', 'URL', 'Datos Scrappeados']); //Creacion del encabezado

    bloquesDeDatos.forEach(bloque => {
        bloque.datos.forEach(datosScrappeados => {
            hoja.addRow([bloque.nombreDePagina, bloque.url, datosScrappeados]); //Se cargan los datos en la hoja.
        });
    });

    await libroDeTrabajo.xlsx.writeFile(ruta); //Se escribe el archivo en la ruta especificada.
    console.log('Archivo Excel guardado en', ruta);
}

//Ejecución del Main.
main();