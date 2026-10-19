const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const fetch = require('node-fetch');

// --- CONFIGURACIÓN ---
const EXCEL_URL = 'https://drive.google.com/uc?export=download&id=1IrwtjQjgkMIv8egyadHo1_Bu93UuYkxK';

// Carpetas que NO son marcas
const CARPETAS_IGNORADAS = ['scripts', 'node_modules', '.git', '.github', 'pedido', 'tutorial', 'original'];

async function sincronizarMarcas() {
    try {
        // 1. Descargar el Excel
        console.log('Descargando Excel desde Google Drive...');
        const response = await fetch(EXCEL_URL);
        if (!response.ok) throw new Error(`Error al descargar: ${response.statusText}`);
        
        const arrayBuffer = await response.arrayBuffer();
        const buffer = Buffer.from(arrayBuffer);

        const workbook = XLSX.read(buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        console.log(`Productos en el Excel: ${data.length}`);

        // 2. Agrupar productos por marca
        const productosPorMarca = {};
        data.forEach(prod => {
            const marca = String(prod.marca).toLowerCase().trim();
            if (!productosPorMarca[marca]) productosPorMarca[marca] = [];
            productosPorMarca[marca].push(prod);
        });

        console.log(`Marcas encontradas en el Excel: ${Object.keys(productosPorMarca).join(', ')}\n`);

        // 3. Procesar cada carpeta de marca
        let totalActualizados = 0;
        let totalNuevos = 0;

        for (const [marca, productosExcel] of Object.entries(productosPorMarca)) {
            const marcaPath = path.join(__dirname, '..', marca, 'index.html');
            
            if (!fs.existsSync(marcaPath)) {
                console.warn(`⚠️ No existe la carpeta: ${marca}/index.html. Saltando...`);
                continue;
            }

            console.log(`Procesando: ${marca} (${productosExcel.length} productos en Excel)`);
            
            let html = fs.readFileSync(marcaPath, 'utf8');
            
            // Buscar el script product-data
            const regex = /(<script id="product-data" type="application\/json">)(.*?)(<\/script>)/s;
            const match = html.match(regex);

            if (!match) {
                console.warn(`   ⚠️ No se encontró 'product-data' en ${marca}/index.html. Saltando...`);
                continue;
            }

            // Extraer el array actual del HTML
            const productosHTML = JSON.parse(match[2].trim());
            const htmlDataByCode = {};
            productosHTML.forEach(p => {
                htmlDataByCode[p.c] = p;
            });

            // Crear el nuevo array fusionado
            let actualizados = 0;
            let nuevos = 0;

            const productosFusionados = productosExcel.map(prodExcel => {
                const prodHTML = htmlDataByCode[prodExcel.c];
                
                if (prodHTML) {
                    // Producto existe: actualizar precio y nombre, mantener imagen
                    if (prodHTML.p !== prodExcel.p) actualizados++;
                    
                    return {
                        ...prodHTML,           // Mantiene i (imagen), cod, g, etc.
                        n: prodExcel.n,        // Actualiza nombre
                        p: prodExcel.p         // ⭐ Actualiza precio
                    };
                } else {
                    // Producto nuevo: sin imagen
                    nuevos++;
                    return {
                        c: prodExcel.c,
                        cod: prodExcel.cod || prodExcel.c,
                        n: prodExcel.n,
                        g: prodExcel.g,
                        p: prodExcel.p,
                        i: "" // Sin imagen
                    };
                }
            });

            // Escribir el nuevo array en el HTML
            const jsonProductos = JSON.stringify(productosFusionados);
            html = html.replace(regex, `$1${jsonProductos}$3`);
            fs.writeFileSync(marcaPath, html);

            console.log(`   ✅ Precios actualizados: ${actualizados}`);
            console.log(`   ✅ Productos nuevos (sin imagen): ${nuevos}`);
            console.log(`   ✅ Total en ${marca}: ${productosFusionados.length}\n`);

            totalActualizados += actualizados;
            totalNuevos += nuevos;
        }

        console.log(`\n🎉 ¡Proceso terminado!`);
        console.log(`   - Total precios actualizados en todas las marcas: ${totalActualizados}`);
        console.log(`   - Total productos nuevos agregados: ${totalNuevos}`);

    } catch (error) {
        console.error('❌ Error crítico:', error);
        process.exit(1);
    }
}

sincronizarMarcas();