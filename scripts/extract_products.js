const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const RAIZ = path.join(__dirname, '..');
const ARCHIVO_SALIDA = path.join(RAIZ, 'productos_extraidos.xlsx');

const CARPETAS_IGNORADAS = [
    'scripts', 'node_modules', '.git', '.github', 'pedido', 'assets', 'imagenes'
];

function extraerProductos() {
    console.log('Iniciando extracción de productos...');
    let todosLosProductos = [];

    const carpetas = fs.readdirSync(RAIZ).filter(item => {
        const rutaCompleta = path.join(RAIZ, item);
        return fs.statSync(rutaCompleta).isDirectory() && !CARPETAS_IGNORADAS.includes(item);
    });

    carpetas.forEach(marca => {
        const archivoHTML = path.join(RAIZ, marca, 'index.html');
        
        if (fs.existsSync(archivoHTML)) {
            const html = fs.readFileSync(archivoHTML, 'utf8');
            const regex = /<script id="product-data" type="application\/json">(.*?)<\/script>/s;
            const match = html.match(regex);

            if (match && match[1]) {
                try {
                    const productos = JSON.parse(match[1].trim());
                    
                    // --- FILTRO PARA EVITAR EL ERROR DE EXCEL ---
                    // Creamos una copia del producto sin la imagen (o cualquier campo > 30000 caracteres)
                    const productosLimpios = productos.map(p => {
                        const productoFiltrado = { marca: marca };
                        
                        for (const [key, value] of Object.entries(p)) {
                            // Si el valor es texto y es muy largo, lo ignoramos (es la imagen base64)
                            if (typeof value === 'string' && value.length > 30000) {
                                continue; 
                            }
                            productoFiltrado[key] = value;
                        }
                        return productoFiltrado;
                    });

                    todosLosProductos = todosLosProductos.concat(productosLimpios);
                    console.log(`✅ ${marca}: ${productos.length} productos extraídos.`);
                } catch (e) {
                    console.error(`❌ Error al parsear JSON en ${marca}/index.html:`, e.message);
                }
            }
        }
    });

    if (todosLosProductos.length === 0) {
        console.log('No se encontraron productos.');
        return;
    }

    // Crear el Excel
    const hoja = XLSX.utils.json_to_sheet(todosLosProductos);
    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hoja, 'Productos');

    XLSX.writeFile(libro, ARCHIVO_SALIDA);

    console.log(`\n🎉 ¡Éxito! Se extrajeron ${todosLosProductos.length} productos.`);
    console.log(`📁 Archivo creado en: ${ARCHIVO_SALIDA}`);
}

extraerProductos();