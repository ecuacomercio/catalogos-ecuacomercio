const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const fetch = require('node-fetch');

// --- CONFIGURACIÓN ---
const ONEDRIVE_EXCEL_URL = 'https://drive.google.com/uc?export=download&id=1IrwtjQjgkMIv8egyadHo1_Bu93UuYkxK';
const HOME_HTML_PATH = path.join(__dirname, '..', 'index.html');

async function actualizarHome() {
    try {
        // 1. Leer el HTML actual para extraer las imágenes (t) y datos existentes
        console.log('Leyendo index.html actual...');
        if (!fs.existsSync(HOME_HTML_PATH)) {
            throw new Error(`No se encontró el archivo: ${HOME_HTML_PATH}`);
        }
        let html = fs.readFileSync(HOME_HTML_PATH, 'utf8');

        const regex = /(<script id="search-index" type="application\/json">)(.*?)(<\/script>)/s;
        const match = html.match(regex);

        if (!match) throw new Error('No se encontró el script search-index en index.html');

        const oldSearchIndex = JSON.parse(match[2].trim());
        console.log(`Productos actuales en el HTML: ${oldSearchIndex.length}`);

        // Diccionario por código para búsqueda rápida
        const oldDataByCode = {};
        oldSearchIndex.forEach(p => {
            oldDataByCode[p.c] = p;
        });

        // 2. Descargar el Excel de OneDrive
        console.log('Descargando Excel desde OneDrive...');
        const response = await fetch(ONEDRIVE_EXCEL_URL);
        if (!response.ok) throw new Error(`Error al descargar: ${response.statusText}`);
        
        const arrayBuffer = await response.arrayBuffer();
        const buffer = Buffer.from(arrayBuffer);

        const workbook = XLSX.read(buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        console.log(`Productos en el Excel: ${data.length}`);

        // 3. Fusionar: construir el nuevo search-index
        let nuevos = 0;
        let actualizados = 0;
        let sinCambios = 0;

        const newSearchIndex = data.map(prod => {
            const oldProd = oldDataByCode[prod.c];
            const marcaLower = String(prod.marca).toLowerCase().trim();
            
            if (oldProd) {
                // Producto existente: conservamos t, b, href originales
                const actualizado = {
                    ...oldProd,               // Mantiene t, b, href del HTML
                    n: prod.n,                // Actualiza nombre por si cambió
                    p: prod.p                 // ⭐ Actualiza precio
                };
                
                if (oldProd.p !== prod.p) actualizados++;
                else sinCambios++;
                
                return actualizado;
            } else {
                // Producto nuevo: creamos la estructura completa
                nuevos++;
                return {
                    n: prod.n,
                    c: prod.c,
                    p: prod.p,
                    b: String(prod.marca).toUpperCase(), // MAYÚSCULAS como en el HTML
                    href: `${marcaLower}/`,
                    t: "" // Sin imagen
                };
            }
        });

        // 4. Escribir el nuevo search-index en el HTML
        const jsonSearch = JSON.stringify(newSearchIndex);
        html = html.replace(regex, `$1${jsonSearch}$3`);
        fs.writeFileSync(HOME_HTML_PATH, html);

        console.log(`\n🎉 ¡Éxito!`);
        console.log(`   - Precios actualizados: ${actualizados}`);
        console.log(`   - Productos sin cambios: ${sinCambios}`);
        console.log(`   - Productos nuevos (sin imagen): ${nuevos}`);
        console.log(`   - Total en el search-index: ${newSearchIndex.length}`);

    } catch (error) {
        console.error('❌ Error crítico:', error);
        process.exit(1);
    }
}

actualizarHome();