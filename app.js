const express = require('express');
const bodyParser = require('body-parser');
const XLSX = require('xlsx');
const fs = require('fs');
const jsdom = require('jsdom');  // Necesitamos jsdom para procesar el HTML y extraer los datos
const { JSDOM } = jsdom;
const app = express();

app.use(bodyParser.json());

// Ruta para generar el archivo XLSX
app.post('/generateXLSX', (req, res) => {
    const htmlContent = req.body.htmlContent;  // Recibe el contenido HTML enviado desde Salesforce

    // Usar jsdom para procesar el HTML
    const dom = new JSDOM(htmlContent);
    const rows = [...dom.window.document.querySelectorAll('tr')];  // Obtener todas las filas (tr) de la tabla

    // Convertir las filas a una matriz de arrays (cada fila será un array de celdas)
    const data = rows.map(row => {
        const cells = [...row.querySelectorAll('td, th')];
        return cells.map(cell => cell.textContent.trim());  // Obtener el texto de cada celda
    });

    // Crear el archivo XLSX
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(data);  // Convierte la matriz de arrays en una hoja de Excel
    XLSX.utils.book_append_sheet(wb, ws, 'Informe');

    // Guardar el archivo temporalmente
    const filePath = './informe_ejecutivo.xlsx';
    XLSX.writeFile(wb, filePath);

    // Enviar el archivo XLSX de vuelta como respuesta
    res.download(filePath, 'Informe_Ejecutivo.xlsx', (err) => {
        if (err) {
            console.error('Error al descargar el archivo:', err);
            res.status(500).send('Error generando el archivo');
        }
        // Eliminar el archivo temporal después de la descarga
        fs.unlinkSync(filePath);
    });
});

// Escuchar en el puerto 3000
app.listen(3000, () => {
    console.log('Servicio de generación de XLSX escuchando en el puerto 3000');
});
