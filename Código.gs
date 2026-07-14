// @ts-nocheck
// 1. LA PUERTA DE ENTRADA
function doGet() {
  return HtmlService.createTemplateFromFile('Interfaz_Pro')
    .evaluate()
    .setTitle('IMPROEX - Sistema de Gestión')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
 
// 2. FUNCIÓN INCLUDE (para dividir HTML)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
 
// 2. CARGA DE CONFIGURACIÓN
function obtenerClientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName("CONFIG");
  return config.getRange("A2:A" + config.getLastRow()).getValues().flat().filter(String);
}
 
// 3. REGISTRO DE NUEVAS ORDENES (BASE_DATOS)
function registrarOC_DesdeApp(obj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bd = ss.getSheetByName("BASE_DATOS");
  bd.appendRow([
    2026, "Marzo", obj.cliente, obj.oc, obj.referencia, 
    Number(obj.cantidad) || 0, 0, obj.fecha, 0, "PROGRAMADO", obj.fecha_ingreso
  ]);
  return "✅ ¡Orden " + obj.oc + " registrada exitosamente!";
}
 
/**
 * 4. REGISTRO DE INGRESOS DE MATERIA PRIMA
 * Guarda en la hoja proporcionada
 */
function registrarNuevoInsumoInventario(obj) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const nombreHoja = "INGRESOS_INVENTARIO"; // Asegúrate de que este sea el nombre en tu Sheet
    let hoja = ss.getSheetByName(nombreHoja);
 
    if (!hoja) {
      hoja = ss.insertSheet(nombreHoja);
      hoja.appendRow(["FECHA_HORA", "TIPO", "CÓDIGO", "LOTE", "FACTURA", "PROVEEDOR", "CANT_KG", "CANT_UND"]);
      hoja.getRange("A1:H1").setFontWeight("bold").setBackground("#f1c40f");
    }
 
    hoja.appendRow([
      new Date(),
      obj.tipo,
      obj.codigo,
      obj.lote,
      obj.factura,
      obj.proveedor,
      obj.cantidad_kg || 0,
      obj.cantidad_und || 0
    ]);
 
    return "✅ Ingreso de " + obj.tipo + " registrado correctamente.";
  } catch (e) {
    return "❌ Error: " + e.toString();
  }
}
 
/**
 * 5. ACTUALIZADO: REPORTE DE PROVISIÓN (CÁLCULO ESTRICTO ISTRAW 50K)
 * Garantiza que 40 cajas de ISTRAW sumen 2,000,000 unidades
 */
function obtenerReporteProvision() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bd = ss.getSheetByName("BASE_DATOS");
 
  // CONEXIÓN AL ARCHIVO EXTERNO DE TRAZABILIDAD
  const ssTrazabilidad = SpreadsheetApp.openById("1QVpDvsyC0VS_xcrMiHp0Cp2Eu9uCkpnyHix6C2pokHY");
  const baseCajas = ssTrazabilidad.getSheetByName("BASE_CAJAS"); 
 
  const dataBD = bd.getDataRange().getValues();
  const dataCajas = baseCajas.getDataRange().getValues();
 
  // FACTORES DE CONVERSIÓN
  const UNIDADES_POR_CAJA = { 
    "TBA 417 PLA": 32000, 
    "TBA 417": 28000, 
    "TBA 418": 26000,
    "ISTRAW": 50000 // REGLA MAESTRA: Todo lo que diga ISTRAW suma 50k
  };
 
  let prodRealUnidades = {};
 
  if (baseCajas) {
    dataCajas.slice(1).forEach(f => {
      // 1. Limpieza de OC: Eliminamos cualquier carácter no numérico
      let ocRaw = f[9]; // Columna J
      let oc = ocRaw ? String(ocRaw).replace(/\D/g, '') : ""; 
 
      // 2. Identificación de Referencia
      let refCaja = String(f[7] || "").toUpperCase(); // Columna H
 
      let factor = 28000; // Valor por defecto
 
      // Aplicación de Prioridad: ISTRAW siempre 50k
      if (refCaja.includes("ISTRAW")) {
        factor = 50000;
      } else {
        for (let clave in UNIDADES_POR_CAJA) {
          if (refCaja.includes(clave)) {
            factor = UNIDADES_POR_CAJA[clave];
            break;
          }
        }
      }
 
      if (oc !== "") {
        prodRealUnidades[oc] = (prodRealUnidades[oc] || 0) + factor;
      }
    });
  }
 
  const hoy = new Date();
 
  return dataBD.slice(1).filter(f => f[3] != "").map(fila => {
    // Limpieza de OC en BASE_DATOS para match perfecto
    const ocBD_Original = String(fila[3]).trim();
    const ocLimpia = ocBD_Original.replace(/\D/g, '');
 
    const cantidadTotalSolicitada = Number(fila[5]) || 0;
    const producidoTotal = prodRealUnidades[ocLimpia] || 0;
 
    const avanceReal = cantidadTotalSolicitada > 0 ? (producidoTotal / cantidadTotalSolicitada) : 0;
    const fechaCompromiso = fila[7] ? new Date(fila[7]) : null;
 
    let diasRestantes = null;
    let semaforo = "🟢 NORMAL";
 
    if (fechaCompromiso && !isNaN(fechaCompromiso.getTime())) {
      const diffTime = fechaCompromiso - hoy;
      diasRestantes = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      if (avanceReal < 0.80 && diasRestantes <= 2 && diasRestantes >= 0) semaforo = "🔴 RIESGO CRÍTICO";
      else if (diasRestantes < 0 && avanceReal < 1) semaforo = "⚠️ ATRASADA";
    }
 
    return {
      cliente: fila[2] ? String(fila[2]).trim() : "SIN CLIENTE",
      oc: ocBD_Original,
      referencia: fila[4],
      cantidad: cantidadTotalSolicitada,
      producido: producidoTotal,
      avance: avanceReal, 
      estado: avanceReal >= 1 ? "FINALIZADO" : (avanceReal > 0 ? "EN PRODUCCIÓN" : "PROGRAMADO"),
      fecha: (fechaCompromiso && !isNaN(fechaCompromiso.getTime())) ? Utilities.formatDate(fechaCompromiso, "GMT-5", "yyyy-MM-dd") : "Sin fecha",
      diasRestantes: diasRestantes,
      semaforo: semaforo
    };
  }).reverse();
}
 
/**
 * 6. CÁLCULO DE MATERIAS PRIMAS (Sin descuentos de mermas)
 */
function obtenerConsumoMateriales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bd = ss.getSheetByName("BASE_DATOS");
  const ingresos = ss.getSheetByName("INGRESOS_INVENTARIO");
  const desperdicio = ss.getSheetByName("DESPERDICIO"); 
  const data = bd.getDataRange().getValues();
 
  let poliTotalKg = 0, filmTotalKg = 0, cartonTotalUnd = 0, plaPuroTotalKg = 0, totalUnidades = 0;
  let detallePoli = {}, detalleFilm = {}, detalleCarton = {}, detalleSoloPLA = {}; 
 
  let totalMermasKg = 0;
  if (desperdicio) {
    const dataD = desperdicio.getDataRange().getValues();
    dataD.slice(1).forEach(f => {
      // Se mantiene la lectura por si se usa en KPIs, pero NO se descuenta del stock real
      totalMermasKg += (Number(f[3]) || 0) + (Number(f[4]) || 0);
    });
  }
 
  let stockReal = { "POLIPROPILENO": 0, "FILM": 0, "CARTON": 0, "PLA": 0 };
  if (ingresos) {
    const dataIngresos = ingresos.getDataRange().getValues();
    dataIngresos.slice(1).forEach(fila => {
      let tipo = String(fila[1] || "").toUpperCase();
      let kg = Number(fila[6]) || 0;
      let und = Number(fila[7]) || 0;
      if (tipo.includes("POLIPROPILENO")) stockReal["POLIPROPILENO"] += kg;
      else if (tipo.includes("FILM")) stockReal["FILM"] += kg;
      else if (tipo.includes("CARTON")) stockReal["CARTON"] += und;
      else if (tipo.includes("PLA") || tipo.includes("BIO")) stockReal["PLA"] += kg;
    });
  }
 
  // DESCUENTO ELIMINADO POR SOLICITUD DEL USUARIO
  // stockReal["POLIPROPILENO"] -= totalMermasKg;
 
  const PESOS_POLI = { "TBA 417": 0.40, "TBA 418": 0.42, "ISTRAW": 0.25, "SORBETON": 2.0, "TIPO H": 1.8, "8 MM": 0.83, "5 MM": 0.42 };
  const FACTORES_FILM = { "TBA 417": 3.2 / 20000, "TBA 418": 3.2 / 20000, "ISTRAW": 3.5 / 50000 };
  const EMBALAJE = { "TBA 417 PLA": 32000, "TBA 417": 28000, "TBA 418": 26000, "ISTRAW 9.8": 50000, "ISTRAW 11.8": 50000 };
 
  data.slice(1).forEach(fila => {
    let cant = parseFloat(fila[5]) || 0; 
    if (cant <= 0) return;
    totalUnidades += cant;
    let ref = String(fila[4] || "").toUpperCase();
    let keyLog = ref;
 
    if (ref.includes("PLA")) {
      let pesoTotalMezcla = (cant * 0.40) / 1000; 
      let pPoli = pesoTotalMezcla * 0.50;  
      let pPLA = pesoTotalMezcla * 0.25;    
      let pFilm = cant * (3.2 / 20000);
      let pCarton = cant / 32000;
      poliTotalKg += pPoli; 
      plaPuroTotalKg += pPLA; 
      filmTotalKg += pFilm; 
      cartonTotalUnd += pCarton;
      detallePoli[keyLog] = (detallePoli[keyLog] || 0) + pPoli;
      detalleSoloPLA[keyLog] = (detalleSoloPLA[keyLog] || 0) + pPLA;
      detalleFilm[keyLog] = (detalleFilm[keyLog] || 0) + pFilm;
      detalleCarton[keyLog] = (detalleCarton[keyLog] || 0) + pCarton;
    } else {
      for (let c in PESOS_POLI) {
        if (ref.includes(c)) {
          let v = (cant * PESOS_POLI[c]) / 1000;
          poliTotalKg += v;
          detallePoli[keyLog] = (detallePoli[keyLog] || 0) + v;
          break;
        }
      }
      for (let c in FACTORES_FILM) {
        if (ref.includes(c)) {
          let v = (cant * FACTORES_FILM[c]);
          filmTotalKg += v;
          detalleFilm[keyLog] = (detalleFilm[keyLog] || 0) + v;
          break;
        }
      }
      let cE = false;
      for (let c in EMBALAJE) {
        if (ref.includes(c)) {
          let v = cant / EMBALAJE[c];
          cartonTotalUnd += v;
          detalleCarton[keyLog] = (detalleCarton[keyLog] || 0) + v;
          cE = true;
          break;
        }
      }
      if (!cE) {
        let v = cant / 28000;
        cartonTotalUnd += v;
        detalleCarton[keyLog] = (detalleCarton[keyLog] || 0) + v;
      }
    }
  });
 
  return {
    poli: poliTotalKg.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}),
    film: filmTotalKg.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}),
    carton: Math.ceil(cartonTotalUnd).toLocaleString('en-US'),
    plaPuro: plaPuroTotalKg.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}),
    totalCant: totalUnidades,
    stockCargado: stockReal,
    mermasTotales: totalMermasKg.toFixed(2),
    detallePoli: detallePoli, 
    detalleFilm: detalleFilm, 
    detalleCarton: detalleCarton, 
    detalleMezclaPLA: detalleSoloPLA
  };
}
/**
 * 7. SUMA DE DESPERDICIO - FILTRO POR TEXTO (MARCO - IMPROEX)
 */
function obtenerMermaMensual() {
  try {
    const idArchivo = "1QVpDvsyC0VS_xcrMiHp0Cp2Eu9uCkpnyHix6C2pokHY"; 
    const ssExterna = SpreadsheetApp.openById(idArchivo);
    let hoja = ssExterna.getSheetByName("DESPERDICIO");
 
    if (!hoja) return "Hoja NO encontrada";
 
    const data = hoja.getDataRange().getValues();
    let sumaMarzo = 0;
 
    // DEFINIMOS LO QUE BUSCAMOS COMO TEXTO (Marzo 2026)
    // Usamos "03" para marzo y "2026"
    const MES_BUSCADO = "03"; 
    const ANIO_BUSCADO = "2026";
 
    for (let i = 1; i < data.length; i++) {
      let fila = data[i];
      let celdaA = fila[0]; // Timestamp
      let celdaE = fila[4]; // Merma
 
      if (!celdaA || celdaA === "") continue;
 
      // --- PASO 1: FORZAR FORMATO DE FECHA A TEXTO "dd/mm/yyyy" ---
      // Esto estandariza cualquier valor de la columna A
      let fechaTexto = Utilities.formatDate(new Date(celdaA), "GMT-5", "dd/MM/yyyy");
      let partes = fechaTexto.split("/"); // [dd, MM, yyyy]
 
      let mesCelda = partes[1];
      let anioCelda = partes[2];
 
      // --- PASO 2: COMPARACIÓN DE TEXTO (MÁS SEGURA) ---
      if (mesCelda === MES_BUSCADO && anioCelda === ANIO_BUSCADO) {
 
        // --- PASO 3: SUMA MATEMÁTICA PURA ---
        let valor = 0;
        if (typeof celdaE === 'number') {
          valor = celdaE;
        } else {
          // Limpiamos el texto por si tiene comas de miles o basura
          let limpio = String(celdaE).replace(/[^0-9,.]/g, '').replace(',', '.');
          valor = parseFloat(limpio);
        }
 
        if (!isNaN(valor)) {
          sumaMarzo += valor;
        }
      }
    }
 
    // Retorno final
    return sumaMarzo.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}) + " kg";
 
  } catch (e) {
    return "Error en suma: " + e.message;
  }
}
/**
 * 8. CONSULTA DINÁMICA DE MERMAS (BUSCADOR MENSUAL)
 * Filtra la hoja "DESPERDICIO" por mes, año y línea de producción.
 * @param {number} mesBuscado - Mes en número (Ej: 3 para Marzo)
 * @param {number} anioBuscado - Año en número (Ej: 2026)
 * @param {string} filtroLinea - "1", "2", "3", "4" o "GLOBAL"
 */
function ejecutarConsultaMermas(mesBuscado, anioBuscado, filtroLinea) {
  try {
    const idArchivo = "1QVpDvsyC0VS_xcrMiHp0Cp2Eu9uCkpnyHix6C2pokHY";
    const ss = SpreadsheetApp.openById(idArchivo);
    const hoja = ss.getSheetByName("DESPERDICIO");
 
    if (!hoja) return { exito: false, mensaje: "Error: Hoja DESPERDICIO no encontrada." };
 
    // Usamos getDisplayValues para leer el texto tal cual aparece en pantalla
    const data = hoja.getDataRange().getDisplayValues();
    let sumaTotal = 0;
    let registrosEncontrados = 0;
 
    for (let i = 1; i < data.length; i++) {
      let fila = data[i];
      let textoFecha = fila[0].trim(); // Columna A
      let lineaCelda = fila[1].trim();  // Columna B
      let valorPeso = fila[4];         // Columna E
 
      if (!textoFecha) continue;
 
      // --- EXTRACCIÓN MANUAL DE FECHA (dd/mm/yyyy) ---
      let partesFecha = textoFecha.split(" ")[0].split("/");
      if (partesFecha.length >= 3) {
        let dia = parseInt(partesFecha[0], 10);
        let mes = parseInt(partesFecha[1], 10);
        let anio = parseInt(partesFecha[2], 10);
 
        // 1. Filtro de Fecha (Mes y Año)
        if (mes === parseInt(mesBuscado) && anio === parseInt(anioBuscado)) {
 
          // 2. Filtro de Línea (Si es "GLOBAL" no filtra, sino busca coincidencia)
          if (filtroLinea === "GLOBAL" || lineaCelda === String(filtroLinea)) {
 
            // 3. Limpieza de valor numérico (evita concatenación de texto)
            let valorLimpio = String(valorPeso).replace(',', '.').replace(/[^\d.]/g, '');
            let num = parseFloat(valorLimpio);
 
            if (!isNaN(num) && num < 5000) { // Filtro de seguridad: máximo 5000kg por registro
              sumaTotal += num;
              registrosEncontrados++;
            }
          }
        }
      }
    }
 
    return {
      exito: true,
      resultado: sumaTotal.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}) + " kg",
      conteo: registrosEncontrados,
      periodo: mesBuscado + "/" + anioBuscado,
      linea: filtroLinea
    };
 
  } catch (e) {
    return { exito: false, mensaje: "Error en Función 8: " + e.message };
  }
}
/**
 * 9. MERMA POR LÍNEA (PARA GRÁFICO)
 * Devuelve la suma de merma (kg) por línea 1-4 para un mes/año.
 * Usa la hoja externa "DESPERDICIO" del archivo de trazabilidad.
 */
function obtenerMermaPorLinea(mesBuscado, anioBuscado) {
  try {
    const idArchivo = "1QVpDvsyC0VS_xcrMiHp0Cp2Eu9uCkpnyHix6C2pokHY";
    const ss = SpreadsheetApp.openById(idArchivo);
    const hoja = ss.getSheetByName("DESPERDICIO");
    if (!hoja) return { exito: false, mensaje: "Hoja DESPERDICIO no encontrada." };
 
    // Leemos texto tal como se ve (evita problemas de formato)
    const data = hoja.getDataRange().getDisplayValues();
 
    // Totales por línea
    const totales = { "1": 0, "2": 0, "3": 0, "4": 0 };
 
    for (let i = 1; i < data.length; i++) {
      const fila = data[i];
 
      const textoFecha = (fila[0] || "").trim(); // Col A: Fecha/Hora
      const linea = (fila[1] || "").trim();      // Col B: Línea
      const valorPeso = fila[4];                 // Col E: Merma (kg)
 
      if (!textoFecha) continue;
      if (!totales.hasOwnProperty(linea)) continue;
 
      // Extrae dd/mm/yyyy (ignora hora)
      const partes = textoFecha.split(" ")[0].split("/");
      if (partes.length < 3) continue;
 
      const mes = parseInt(partes[1], 10);
      const anio = parseInt(partes[2], 10);
 
      if (mes !== parseInt(mesBuscado, 10)) continue;
      if (anio !== parseInt(anioBuscado, 10)) continue;
 
      // Limpieza numérica segura (coma/puntos)
      const limpio = String(valorPeso).replace(',', '.').replace(/[^\d.]/g, '');
      const num = parseFloat(limpio);
 
      if (!isNaN(num) && num < 5000) totales[linea] += num;
    }
 
    return {
      exito: true,
      labels: ["Línea 1", "Línea 2", "Línea 3", "Línea 4"],
      values: [totales["1"], totales["2"], totales["3"], totales["4"]],
      mes: mesBuscado,
      anio: anioBuscado
    };
  } catch (e) {
    return { exito: false, mensaje: "Error en Función 9: " + e.message };
  }
}
 
/**
 * 10. Función Maestra de Producción por Línea
 * Automatizada para mostrar datos del mes actual por defecto.
 */
function obtenerProduccionPorLinea(params) {
  try {
    const idTrazabilidad = "1QVpDvsyC0VS_xcrMiHp0Cp2Eu9uCkpnyHix6C2pokHY";
    const ssTrazabilidad = SpreadsheetApp.openById(idTrazabilidad);
    const hojaProduccion = ssTrazabilidad.getSheetByName("BASE_CAJAS");

    if (!hojaProduccion) return { exito: false, mensaje: 'No se encontró la pestaña "BASE_CAJAS"' };

    const datos = hojaProduccion.getDataRange().getDisplayValues();
    if (datos.length <= 1) return { exito: true, datos: [], resumen: [] };

    // --- CONFIGURACIÓN DE FILTROS AUTOMÁTICOS ---
    const p = params || {};
    let fInicio, fFin;

    if (p.fechaInicio && p.fechaFin) {
      // Si el usuario eligió fechas en el sidebar
      fInicio = new Date(p.fechaInicio + "T00:00:00");
      fFin = new Date(p.fechaFin + "T23:59:59");
    } else {
      // Carga automática: Desde el 1ero de este mes hasta hoy
      let hoy = new Date();
      fInicio = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
      fFin = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate(), 23, 59, 59);
    }

    let resumen = {};

    // --- PROCESAMIENTO DE FILAS ---
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const timestamp = fila[0]; // Columna A
      const linea = String(fila[4] || '').trim(); // Columna E
      const cliente = String(fila[5] || '').trim(); // Columna F
      const referencia = String(fila[7] || '').trim(); // Columna H
      const ocNum = String(fila[9] || '').trim(); // Columna J

      if (!linea || !timestamp) continue;

      // 1. Validar Rango de Fecha
      const partes = timestamp.split(" ")[0].split("/");
      if (partes.length === 3) {
        const fechaFila = new Date(partes[2], partes[1] - 1, partes[0]);
        if (fechaFila < fInicio || fechaFila > fFin) continue;
      }

      // 2. Aplicar Filtros de Sidebar (si existen)
      if (p.linea && p.linea !== 'TODAS' && linea !== p.linea) continue;
      if (p.cliente && cliente.toUpperCase() !== p.cliente.toUpperCase()) continue;
      if (p.referencia && !referencia.toUpperCase().includes(p.referencia.toUpperCase())) continue;

      // 3. Inicializar objeto de línea
      if (!resumen[linea]) {
        resumen[linea] = { 
          linea: linea, 
          unidades: 0, 
          ocsUnicas: new Set(), 
          cajas: 0,
          clientes: new Set()
        };
      }

      // 4. Mapeo de Unidades (1 fila = 1 caja)
      let u = 28000; // Valor base
      const refUP = referencia.toUpperCase();
      const cliUP = cliente.toUpperCase();

      if (refUP.includes("9.8") || cliUP.includes("NESTLE")) u = 50000;
      else if (refUP.includes("418")) u = 26000;
      else if (cliUP.includes("VITA") && refUP.includes("PLA")) u = 32000;
      else if (refUP.includes("SORBETON")) u = 10000; // Ejemplo adicional

      resumen[linea].unidades += u;
      resumen[linea].ocsUnicas.add(ocNum);
      resumen[linea].cajas++;
      resumen[linea].clientes.add(cliente);
    }

    // --- CONSTRUCCIÓN DE RESULTADO FINAL ---
    const datosFinales = Object.values(resumen).map(l => ({
      linea: l.linea,
      ocs: l.ocsUnicas.size, // Conteo real de OCs únicas
      unidades: l.unidades,
      eficiencia: (l.linea === "1" ? 88 : l.linea === "2" ? 92 : l.linea === "3" ? 82 : 85),
      cliente: Array.from(l.clientes).slice(0, 2).join(", "),
      referencia: "Cajas: " + l.cajas
    }));

    // Ordenar por número de línea
    datosFinales.sort((a, b) => a.linea.localeCompare(b.linea));

    return {
      exito: true,
      datos: datosFinales,
      resumen: datosFinales.map(d => ({ numero: d.linea, eficiencia: d.eficiencia, unidades: d.unidades })),
      mensaje: "Datos cargados correctamente."
    };

  } catch (e) {
    console.error("Error en Función 10: " + e.toString());
    return { exito: false, mensaje: "Error de servidor: " + e.toString() };
  }
}