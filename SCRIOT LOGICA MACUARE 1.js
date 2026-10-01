/**
 * CÓDIGO CONSOLIDADO, UNIFICADO Y OPTIMIZADO - SISTEMA MACUARE HÍPICO
 * Pestañas: Usuarios, Carreras, Tickets, Resultados, Caja_Vendedores, Historial_Metricas, Ajustes_Banca
 * Actualizado con Desglose Dinámico y Cálculo Real de Premios por Sub-Jugada
 */

const COMISION_QUEDA = 0.30;
const MULTIPLICADOR_PREMIO = 10; // Relación fija de pago ajustada a 10 a 1
const CUOTA_SEGURIDAD_COMBINACION = 1000; // Límite máximo por combinación exacta para bloqueo o fraccionamiento

// Función auxiliar optimizada con CacheService para leer el Fondo de Banca
function obtenerFondoBanca() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("fondo_banca");
  if (cached !== null) {
    return parseFloat(cached);
  }
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Ajustes_Banca");
    if (sheet) {
      var val = sheet.getRange("F2").getValue();
      if (val !== "" && !isNaN(val)) {
        cache.put("fondo_banca", val, 300); // Guardar por 5 minutos
        return parseFloat(val);
      }
    }
  } catch (e) {
    // Valor por defecto ante error
  }
  return 100000;
}

// Función auxiliar ultra rápida: Lee directamente F2 y G2 de Carreras (Sin lecturas masivas)
function validarHoraCierreApuestas() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaCarreras = ss.getSheetByName("Carreras");
    if (!hojaCarreras) return { permitido: true };
    
    var fechaValida1Raw = hojaCarreras.getRange("F2").getValue();
    var horaValida1Str = String(hojaCarreras.getRange("G2").getValue()).trim();
    
    if (!horaValida1Str) return { permitido: true };
    
    var fechaActual = new Date();
    var anio = fechaActual.getFullYear();
    var mes = fechaActual.getMonth();
    var dia = fechaActual.getDate();
    
    if (fechaValida1Raw) {
      if (fechaValida1Raw instanceof Date) {
        anio = fechaValida1Raw.getFullYear();
        mes = fechaValida1Raw.getMonth();
        dia = fechaValida1Raw.getDate();
      } else {
        var fechaValida1Str = String(fechaValida1Raw).trim();
        if (fechaValida1Str.indexOf('T') !== -1) fechaValida1Str = fechaValida1Str.split('T')[0];
        var partesFecha = fechaValida1Str.split("-");
        if (partesFecha.length === 3) {
          anio = parseInt(partesFecha[0], 10);
          mes = parseInt(partesFecha[1], 10) - 1;
          dia = parseInt(partesFecha[2], 10);
        }
      }
    }
    
    var horaCarrera1 = new Date(anio, mes, dia);
    var horaLimpia = horaValida1Str.toLowerCase().replace(/\s+/g, '');
    var esPM = horaLimpia.indexOf('pm') !== -1;
    var esAM = horaLimpia.indexOf('am') !== -1;
    
    horaLimpia = horaLimpia.replace('pm', '').replace('am', '');
    var separador = horaLimpia.indexOf(':') !== -1 ? ':' : '.';
    var partesHora = horaLimpia.split(separador);
    
    if (partesHora.length >= 2) {
      var h = parseInt(partesHora[0], 10) || 0;
      var m = parseInt(partesHora[1], 10) || 0;
      
      if (esPM && h < 12) h += 12;
      if (esAM && h === 12) h = 0;
      
      horaCarrera1.setHours(h, m, 0, 0);
    } else {
      horaCarrera1.setHours(23, 59, 59, 0);
    }
    
    var horaCierreLimite = new Date(horaCarrera1.getTime() - (10 * 60 * 1000)); // 10 min antes
    
    if (fechaActual.getTime() >= horaCierreLimite.getTime()) {
      return { 
        permitido: false, 
        mensaje: `Las apuestas están cerradas. La hora límite de la jornada ya pasó (${Utilities.formatDate(horaCarrera1, Session.getScriptTimeZone(), "dd/MM/yyyy hh:mm a")}).` 
      };
    }

  } catch (e) {
    // Continuidad por defecto ante errores menores
  }
  return { permitido: true };
}

// SISTEMA RÁPIDO DE FRACCIONAMIENTO Y CONTROL DE DUPLICADOS CON DATOS EN MEMORIA
function verificarFraccionamientoYDuplicadosConDatos(dataTickets, combinacionStr, montoIngresado) {
  var limitesPorCombinacion = {};
  var totalApuestasAcumuladas = 0;
  
  for (var i = 1; i < dataTickets.length; i++) {
    var combRegistrada = String(dataTickets[i][3] || "").trim();
    var estatusExistente = String(dataTickets[i][6]).trim();
    var montoRegistrado = parseFloat(dataTickets[i][4]) || 0;
    
    totalApuestasAcumuladas += montoRegistrado;
    if (estatusExistente.toLowerCase() === "anulado") continue;
    
    if (combRegistrada.includes(" | ")) {
      var partesAnt = combRegistrada.split(" | ");
      partesAnt.forEach(function(p) {
        var match = p.match(/(.*?)\s*\(Bs\s*([\d\.]+)\)/);
        if (match) {
          var cExacta = match[1].trim();
          var mExacto = parseFloat(match[2]) || 0;
          limitesPorCombinacion[cExacta] = (limitesPorCombinacion[cExacta] || 0) + mExacto;
        }
      });
    } else if (combRegistrada.includes("(Bs")) {
      var matchSimple = combRegistrada.match(/(.*?)\s*\(Bs\s*([\d\.]+)\)/);
      if (matchSimple) {
        var cExacta = matchSimple[1].trim();
        limitesPorCombinacion[cExacta] = (limitesPorCombinacion[cExacta] || 0) + montoRegistrado;
      }
    } else {
      limitesPorCombinacion[combRegistrada] = (limitesPorCombinacion[combRegistrada] || 0) + montoRegistrado;
    }
  }
  
  var subJugadas = combinacionStr.includes(" | ") ? combinacionStr.split(" | ") : [combinacionStr];
  var detallesAfectados = [];
  
  for (var j = 0; j < subJugadas.length; j++) {
    var item = subJugadas[j];
    var combActual = item;
    var montoActual = montoIngresado;
    
    var matchItem = item.match(/(.*?)\s*\(Bs\s*([\d\.]+)\)/);
    if (matchItem) {
      combActual = matchItem[1].trim();
      montoActual = parseFloat(matchItem[2]) || 0;
    }
    
    var vendidoAcumulado = limitesPorCombinacion[combActual] || 0;
    var totalConNuevaJugada = vendidoAcumulado + montoActual;
    
    if (totalConNuevaJugada > CUOTA_SEGURIDAD_COMBINACION) {
      var cupoRestante = CUOTA_SEGURIDAD_COMBINACION - vendidoAcumulado;
      if (cupoRestante < 0) cupoRestante = 0;
      var numeroJugadaRef = j + 1;
      
      detallesAfectados.push({
        indiceItem: j,
        numeroJugada: numeroJugadaRef,
        cupoRestante: cupoRestante,
        combinacionAfectada: combActual,
        montoIntentado: montoActual,
        mensaje: `La combinación [${combActual}] supera el cupo de seguridad. Monto disponible restante: Bs ${cupoRestante.toFixed(2)}.`
      });
    }
  }
  
  if (detallesAfectados.length > 0) {
    var mensajesUnicos = [];
    var vistoMensaje = {};
    detallesAfectados.forEach(function(d) {
      if (!vistoMensaje[d.mensaje]) {
        vistoMensaje[d.mensaje] = true;
        mensajesUnicos.push(d.mensaje);
      }
    });

    var mensajePrincipal = mensajesUnicos.length === 1 
      ? mensajesUnicos[0] 
      : "Se han detectado las siguientes restricciones de cupo:\n• " + mensajesUnicos.join("\n• ");

    return {
      requiereFraccionamiento: true,
      detallesAfectados: detallesAfectados,
      mensaje: mensajePrincipal,
      totalApuestasAcumuladas: totalApuestasAcumuladas
    };
  }
  
  return { requiereFraccionamiento: false, totalApuestasAcumuladas: totalApuestasAcumuladas };
}

// GENERADOR DINÁMICO DE ENLACES DE ACCESO PARA EL JUGADOR
function generarTokenAccesoJugador(ticketId, referenciaPago) {
  var tokenUnico = "ACCESO-" + Utilities.getUuid().substring(0, 10).toUpperCase();
  var urlBaseWebApp = ScriptApp.getService().getUrl();
  var enlaceUnico = urlBaseWebApp + "?action=verTicket&token=" + tokenUnico + "&ticket=" + ticketId + "&ref=" + encodeURIComponent(referenciaPago);
  
  return {
    token: tokenUnico,
    enlace: enlaceUnico
  };
}

// Obtener nombre del usuario activo
function obtenerUsuarioLogueadoBackend(nombreUsuario) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetUsr = ss.getSheetByName("Usuarios");
  var dataUsr = sheetUsr.getDataRange().getValues();
  
  for (var u = 1; u < dataUsr.length; u++) {
    if (String(dataUsr[u][1]).trim() === String(nombreUsuario).trim()) {
      return { success: true, usuario: dataUsr[u][1], rol: dataUsr[u][3] };
    }
  }
  return { success: true, usuario: nombreUsuario || "Taquilla Principal", rol: "Vendedor" };
}

// REGISTRAR MÉTRICAS Y REPORTAR PAGO
function registrarMetrica(params) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    return { success: false, mensaje: "El sistema está ocupado procesando otra solicitud. Intente de nuevo." };
  }
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Caja_Vendedores");
    if (!sheet) {
      return { success: false, mensaje: "No se encontró la hoja Caja_Vendedores" };
    }
    
    var taquilla = params.taquilla || "";
    var ventas = params.ventas || 0;
    var premios = params.premios || 0;
    var laQueda = params.laQueda || 0;
    var comision = params.comision || 0;
    var neto = params.neto || 0;
    var referencia = params.referencia || "";
    var estatus = params.estatus || "PENDIENTE";
    
    var data = sheet.getDataRange().getValues();
    var filaEncontrada = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === String(taquilla).trim().toLowerCase()) {
        filaEncontrada = i + 1;
        break;
      }
    }
    
    if (filaEncontrada > -1) {
      sheet.getRange(filaEncontrada, 2, 1, 7).setValues([[
        ventas, premios, laQueda, comision, neto, referencia, estatus
      ]]);
    } else {
      sheet.appendRow([taquilla, ventas, premios, laQueda, comision, neto, referencia, estatus]);
    }
    
    return { success: true, mensaje: "Métrica y pago reportados correctamente" };
  } catch (error) {
    return { success: false, mensaje: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

// APROBAR PAGO DESDE EL ADMIN
function aprobarPagoMetrica(params) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    return { success: false, mensaje: "Sistema ocupado. Intente de nuevo." };
  }
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Caja_Vendedores");
    if (!sheet) {
      return { success: false, mensaje: "No se encontró la hoja Caja_Vendedores" };
    }
    
    var taquilla = String(params.taquilla || "").trim().toLowerCase();
    var data = sheet.getDataRange().getValues();
    var cabeceras = data[0];
    
    var colVendedor = -1;
    var colEstatus = -1;
    
    for (var j = 0; j < cabeceras.length; j++) {
      var head = String(cabeceras[j]).trim().toLowerCase();
      if (head === "vendedor" || head === "taquilla") colVendedor = j;
      if (head === "estatus" || head === "status") colEstatus = j;
    }
    
    if (colVendedor === -1) colVendedor = 0;
    if (colEstatus === -1) colEstatus = 7;
    
    var filaEncontrada = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][colVendedor]).trim().toLowerCase() === taquilla) {
        filaEncontrada = i + 1;
        break;
      }
    }
    
    if (filaEncontrada > -1) {
      sheet.getRange(filaEncontrada, colEstatus + 1).setValue("VERIFICADO");
      if (params.referencia) {
        sheet.getRange(filaEncontrada, 7).setValue(String(params.referencia).trim());
      }
      return { success: true, mensaje: "¡Pago verificado con éxito!" };
    } else {
      return { success: false, mensaje: "No se encontró la taquilla '" + params.taquilla + "'." };
    }
  } catch (error) {
    return { success: false, mensaje: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

// CONTROLADOR PRINCIPAL doGet OPTIMIZADO PARA MÁXIMA VELOCIDAD
function doGet(e) {
  var params = e && e.parameter ? e.parameter : {};
  var action = params.action;
  var callback = params.callback; 
  var resultado = { success: false };

  if (action === "login") {
    resultado = verificarLogin(params.usuario, params.password);
  } else if (action === "obtenerUsuario") {
    resultado = obtenerUsuarioLogueadoBackend(params.usuario);
  } else if (action === "registrarMetrica") {
    resultado = registrarMetrica(params);
  } else if (action === "aprobarPagoMetrica" || action === "aprobarBalance") {
    resultado = aprobarPagoMetrica(params);
  } else if (action === "verTicket") {
    var tokenAcceso = params.token;
    var idBoleto = params.ticket;
    var refPago = params.ref;
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetTickets = ss.getSheetByName("Tickets");
    var dataTickets = sheetTickets.getDataRange().getValues();
    var ticketInfo = null;
    
    for (var i = 1; i < dataTickets.length; i++) {
      if (String(dataTickets[i][0]).trim() === String(idBoleto).trim()) {
        ticketInfo = {
          ID_Ticket: dataTickets[i][0],
          Vendedor: dataTickets[i][1],
          Jugador: dataTickets[i][2],
          Seleccion6: dataTickets[i][3],
          Monto: dataTickets[i][4],
          PremioPotencial: dataTickets[i][5],
          Estatus: dataTickets[i][6],
          Referencia: dataTickets[i][10] || refPago
        };
        break;
      }
    }
    
    if (ticketInfo) {
      resultado = { success: true, mensaje: "Token verificado correctamente.", token: tokenAcceso, data: ticketInfo };
    } else {
      resultado = { success: false, mensaje: "El ticket asociado no fue encontrado." };
    }
  } else if (action === "generarEnlaceToken") {
    var datosGeneracion = generarTokenAccesoJugador(params.ticketId, params.referencia || "");
    resultado = { success: true, token: datosGeneracion.token, enlace: datosGeneracion.enlace };
  } else if (action === "getTopeRiesgo") {
    resultado = { success: true, topeRiesgo: obtenerFondoBanca() };
  } else if (action === "guardarTopeRiesgo") {
    try {
      var nuevoTope = parseFloat(params.topeRiesgo);
      if (isNaN(nuevoTope)) {
        resultado = { success: false, message: "Valor de riesgo inválido." };
      } else {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheetAjustes = ss.getSheetByName("Ajustes_Banca") || ss.insertSheet("Ajustes_Banca");
        sheetAjustes.getRange("F2").setValue(nuevoTope);
        CacheService.getScriptCache().remove("fondo_banca");
        resultado = { success: true, message: "¡Éxito! actualizado: " + nuevoTope };
      }
    } catch (err) {
      resultado = { success: false, message: "Error al guardar la combinación: " + err.toString() };
    }
  } else if (action === "guardarResultados") {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetRes = ss.getSheetByName("Resultados");
    var sheetCar = ss.getSheetByName("Carreras");
    
    var dataCar = sheetCar.getDataRange().getValues();
    var carrerasMap = {};
    for (var r = 1; r < dataCar.length; r++) {
      var val = String(dataCar[r][0]).trim();
      var numCab = dataCar[r][1];
      var estCab = String(dataCar[r][3]).trim().toLowerCase();
      
      if (val && numCab !== "" && estCab !== "retirado") {
        if (!carrerasMap[val]) carrerasMap[val] = [];
        carrerasMap[val].push(Number(numCab));
      }
    }
    
    var ajustesSheet = ss.getSheetByName("Ajustes_Banca");
    var ajustesData = ajustesSheet ? ajustesSheet.getDataRange().getValues() : [];
    var reglasPuestos = [];
    
    for (var a = 1; a < ajustesData.length; a++) {
      var minInc = parseInt(ajustesData[a][1]);
      var maxInc = parseInt(ajustesData[a][2]);
      var puestosOt = parseInt(ajustesData[a][3]);
      if (!isNaN(minInc) && !isNaN(maxInc) && !isNaN(puestosOt)) {
        reglasPuestos.push({ min: minInc, max: maxInc, puestos: puestosOt });
      }
    }
    
    var puestosValidaCalc = {};
    for (var vKey in carrerasMap) {
      var totalInscritos = carrerasMap[vKey].length;
      var puestosAsignados = 1;
      
      if (reglasPuestos.length > 0) {
        for (var rp = 0; rp < reglasPuestos.length; rp++) {
          if (totalInscritos >= reglasPuestos[rp].min && totalInscritos <= reglasPuestos[rp].max) {
            puestosAsignados = reglasPuestos[rp].puestos;
            break;
          }
        }
      } else {
        puestosAsignados = totalInscritos >= 12 ? 3 : (totalInscritos >= 10 ? 2 : 1);
      }
      puestosValidaCalc[vKey] = puestosAsignados;
    }
    
    var v1Data = params.v1 ? params.v1.split(",") : [];
    var v2Data = params.v2 ? params.v2.split(",") : [];
    var v3Data = params.v3 ? params.v3.split(",") : [];
    var v4Data = params.v4 ? params.v4.split(",") : [];
    var v5Data = params.v5 ? params.v5.split(",") : [];
    var v6Data = params.v6 ? params.v6.split(",") : [];
    
    var matrizResultados = [
      ["Valida 1", v1Data[0] || "", v1Data[1] || "", v1Data[2] || "", v1Data[0] ? (params.p1 || puestosValidaCalc["Valida 1"] || 1) : ""],
      ["Valida 2", v2Data[0] || "", v2Data[1] || "", v2Data[2] || "", v2Data[0] ? (params.p2 || puestosValidaCalc["Valida 2"] || 1) : ""],
      ["Valida 3", v3Data[0] || "", v3Data[1] || "", v3Data[2] || "", v3Data[0] ? (params.p3 || puestosValidaCalc["Valida 3"] || 1) : ""],
      ["Valida 4", v4Data[0] || "", v4Data[1] || "", v4Data[2] || "", v4Data[0] ? (params.p4 || puestosValidaCalc["Valida 4"] || 1) : ""],
      ["Valida 5", v5Data[0] || "", v5Data[1] || "", v5Data[2] || "", v5Data[0] ? (params.p5 || puestosValidaCalc["Valida 5"] || 1) : ""],
      ["Valida 6", v6Data[0] || "", v6Data[1] || "", v6Data[2] || "", v6Data[0] ? (params.p6 || puestosValidaCalc["Valida 6"] || 1) : ""]
    ];
    
    sheetRes.getRange(2, 1, 6, 5).setValues(matrizResultados);
    procesarResultadosYCalcularGanadores();
    resultado = { success: true, message: "¡Resultados guardados y estados actualizados!" };
  } else if (action === "guardarCarrera") {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetCar = ss.getSheetByName("Carreras");
    var dataCar = sheetCar.getDataRange().getValues();
    var filaDestino = -1;
    
    for (var i = 1; i < dataCar.length; i++) {
      if (String(dataCar[i][0]).trim() === String(params.valida).trim() && (!dataCar[i][1] || dataCar[i][1] === "")) {
        filaDestino = i + 1;
        break;
      }
    }
    if (filaDestino === -1) filaDestino = dataCar.length + 1;
    
    sheetCar.getRange(filaDestino, 1, 1, 9).setValues([[
      params.valida, params.numero, params.ejemplar, params.estado || "CORRIO",
      params.hipodromo || "", params.fecha || "", params.hora || "", params.distancia || "", params.reunion || ""
    ]]);
    resultado = { success: true, message: "¡Ejemplar registrado con éxito!" };
  } else if (action === "retirarEjemplar") {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetCar = ss.getSheetByName("Carreras");
    var dataCar = sheetCar.getDataRange().getValues();
    var encontrado = false;
    
    for (var j = 1; j < dataCar.length; j++) {
      if (String(dataCar[j][0]).trim() === String(params.carrera).trim() && String(dataCar[j][1]).trim() === String(params.numero).trim()) {
        sheetCar.getRange(j + 1, 4).setValue("Retirado");
        encontrado = true;
        break;
      }
    }
    resultado = encontrado ? { success: true, message: "¡Ejemplar retirado!" } : { success: false, message: "Ejemplar no encontrado." };
  } else if (action === "pagarTicket" || action === "marcarPagado") {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetTickets = ss.getSheetByName("Tickets");
    var dataTickets = sheetTickets.getDataRange().getValues();
    var codigoTicket = String(params.idTicket || params.codigo || "").trim();
    var referencia = String(params.referencia || "").trim();
    var encontrado = false;
    
    for (var t = 1; t < dataTickets.length; t++) {
      if (String(dataTickets[t][0]).trim() === codigoTicket) {
        var estTicketActual = String(dataTickets[t][6]).trim().toLowerCase();
        if (estTicketActual === "ganador" || estTicketActual === "ganadores") {
          sheetTickets.getRange(t + 1, 7).setValue("Pagado");
          sheetTickets.getRange(t + 1, 11).setValue(referencia);
          encontrado = true;
          break;
        } else {
          return ContentService.createTextOutput(JSON.stringify({ success: false, message: "El ticket no está en estatus Ganador o ya fue pagado." })).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }
    if (encontrado) {
      recalcularCajasVendedores();
      resultado = { success: true, message: "¡Ticket pagado y caja actualizada!" };
    } else {
      resultado = { success: false, message: "No se encontró el ticket." };
    }
  } else if (action === "getTickets") {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Tickets");
    var sheetRes = ss.getSheetByName("Resultados");
    var sheetCar = ss.getSheetByName("Carreras");
    var lastRow = sheet.getLastRow();
    var listTickets = [];
    
    // Precargar datos oficiales y carreras para calcular el premio real dinámico por sub-jugada
    var resultadosData = sheetRes ? sheetRes.getDataRange().getValues() : [];
    var carrerasData = sheetCar ? sheetCar.getDataRange().getValues() : [];
    
    var resultadosOficiales = {};
    var carrerasOficialesCorridas = {};
    for (var r = 1; r < resultadosData.length; r++) {
      if (resultadosData[r][0]) {
        var valName = resultadosData[r][0];
        var ganadoresVal = [];
        var cuantosPuestos = parseInt(resultadosData[r][4]) || 3;
        var tieneGanadorReal = (resultadosData[r][1] !== "" && resultadosData[r][1] !== null && resultadosData[r][1] !== undefined);
        if (tieneGanadorReal) {
          for (var col = 1; col <= cuantosPuestos; col++) {
            if (resultadosData[r][col] !== "" && resultadosData[r][col] !== null && resultadosData[r][col] !== undefined) {
              ganadoresVal.push(parseInt(resultadosData[r][col]));
            }
          }
          resultadosOficiales[valName] = ganadoresVal;
          carrerasOficialesCorridas[valName] = true;
        }
      }
    }

    var caballosRetirados = {};
    for (var c = 1; c < carrerasData.length; c++) {
      var valCar = String(carrerasData[c][0]).trim();
      var numCar = parseInt(carrerasData[c][1]);
      var estCar = String(carrerasData[c][3]).trim().toLowerCase();
      if (valCar && !isNaN(numCar) && estCar === "retirado") {
        if (!caballosRetirados[valCar]) caballosRetirados[valCar] = {};
        caballosRetirados[valCar][numCar] = true;
      }
    }

    var caballosActivosPorValida = {};
    for (var c = 1; c < carrerasData.length; c++) {
      var valCar = String(carrerasData[c][0]).trim();
      var numCar = parseInt(carrerasData[c][1]);
      var estCar = String(carrerasData[c][3]).trim().toLowerCase();
      if (valCar && !isNaN(numCar) && estCar !== "retirado") {
        if (!caballosActivosPorValida[valCar]) caballosActivosPorValida[valCar] = [];
        caballosActivosPorValida[valCar].push(numCar);
      }
    }

    var favoritosPorValida = {};
    for (var valKey in caballosActivosPorValida) {
      var activos = caballosActivosPorValida[valKey];
      if (activos.length > 0) favoritosPorValida[valKey] = activos[0];
    }
    
    if (lastRow > 1) {
      var startRow = Math.max(2, lastRow - 300); 
      var numRows = lastRow - startRow + 1;
      var dataTickets = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();
      
      for (var t = 0; t < dataTickets.length; t++) {
        var row = dataTickets[t];
        if (row[0] || row[1]) {
          var seleccionStr = String(row[3]).trim();
          var montoTotal = parseFloat(row[4]) || 0;
          var estatusTicket = String(row[6]).trim();
          
          var premioPotencialGuardado = parseFloat(row[5]) || (montoTotal * MULTIPLICADOR_PREMIO);
          var premioRealCalculado = 0;
          var desglosesPremio = [];
          var premioPotencialFormateado = "";
          
          if (seleccionStr.includes(" | ")) {
            var subJugadas = seleccionStr.split(" | ");
            var montoPorSub = montoTotal / subJugadas.length;
            var todasCorrieron = true;
            
            subJugadas.forEach(function(item) {
              var combPart = item;
              var mSub = montoPorSub;
              var matchSub = item.match(/(.*?)\s*\(Bs\s*([\d\.]+)\)/);
              if (matchSub) {
                combPart = matchSub[1].trim();
                mSub = parseFloat(matchSub[2]) || 0;
              }
              
              var caballosElegidos = combPart.split(" - ");
              var subGanadora = true;
              var subCorrioTodas = true;
              
              for (var v = 0; v < 6; v++) {
                var validaKey = "Valida " + (v + 1);
                var seleccionCaballo = parseInt(caballosElegidos[v]);
                
                if (caballosRetirados[validaKey] && caballosRetirados[validaKey][seleccionCaballo]) {
                  if (favoritosPorValida[validaKey] !== undefined) seleccionCaballo = favoritosPorValida[validaKey];
                }
                
                if (carrerasOficialesCorridas[validaKey]) {
                  var ganadoresValida = resultadosOficiales[validaKey] || [];
                  if (ganadoresValida.indexOf(seleccionCaballo) === -1) {
                    subGanadora = false;
                    break;
                  }
                } else {
                  subCorrioTodas = false;
                  todasCorrieron = false;
                }
              }
              
              if (subGanadora && subCorrioTodas) {
                var pSub = mSub * MULTIPLICADOR_PREMIO;
                premioRealCalculado += pSub;
                desglosesPremio.push("Bs " + pSub.toLocaleString('es-VE', {minimumFractionDigits: 2, maximumFractionDigits: 2}) + " (Ganada)");
              } else if (!subCorrioTodas) {
                desglosesPremio.push("Pendiente");
              } else {
                desglosesPremio.push("Perdida");
              }
            });
            
            if (premioRealCalculado > 0) {
              premioPotencialFormateado = "Bs " + premioRealCalculado.toLocaleString('es-VE', {minimumFractionDigits: 2, maximumFractionDigits: 2}) + " (Real / Pot. Bs " + premioPotencialGuardado.toLocaleString('es-VE', {minimumFractionDigits: 2, maximumFractionDigits: 2}) + ")";
            } else if (!todasCorrieron) {
              premioPotencialFormateado = "Bs 0,00 (Pendiente / Pot. Bs " + premioPotencialGuardado.toLocaleString('es-VE', {minimumFractionDigits: 2, maximumFractionDigits: 2}) + ")";
            } else {
              premioPotencialFormateado = "Bs 0,00 (Pot. Bs " + premioPotencialGuardado.toLocaleString('es-VE', {minimumFractionDigits: 2, maximumFractionDigits: 2}) + ")";
            }
          } else {
            // Ticket simple
            var subGanadoraSimple = true;
            var subCorrioSimple = true;
            var combPart = seleccionStr;
            var matchSimple = seleccionStr.match(/(.*?)\s*\(Bs\s*([\d\.]+)\)/);
            if (matchSimple) combPart = matchSimple[1].trim();
            var caballosElegidos = combPart.split(" - ");
            
            for (var v = 0; v < 6; v++) {
              var validaKey = "Valida " + (v + 1);
              var selC = parseInt(caballosElegidos[v]);
              if (carrerasOficialesCorridas[validaKey]) {
                var ganadoresValida = resultadosOficiales[validaKey] || [];
                if (ganadoresValida.indexOf(selC) === -1) {
                  subGanadoraSimple = false;
                  break;
                }
              } else {
                subCorrioSimple = false;
              }
            }
            
            if (subGanadoraSimple && subCorrioSimple) {
              premioRealCalculado = montoTotal * MULTIPLICADOR_PREMIO;
              premioPotencialFormateado = "Bs " + premioRealCalculado.toLocaleString('es-VE', {minimumFractionDigits: 2, maximumFractionDigits: 2});
            } else if (!subCorrioSimple) {
              premioPotencialFormateado = "Bs " + premioPotencialGuardado.toLocaleString('es-VE', {minimumFractionDigits: 2, maximumFractionDigits: 2}) + " (Potencial)";
            } else {
              premioPotencialFormateado = "Bs 0,00";
            }
          }

          listTickets.push({
            ID_Ticket: row[0], 
            Vendedor: row[1], 
            Jugador: row[2], 
            Seleccion6: row[3],
            Monto: row[4], 
            PremioPotencial: premioPotencialFormateado, 
            Estatus: row[6], 
            Banco_PagoMovil: row[7],
            Telefono_PagoMovil: row[8], 
            Cedula_PagoMovil: row[9], 
            Referencia_PagoMovil: row[10] || ""
          });
        }
      }
    }
    resultado = { success: true, data: listTickets.reverse() };
  } else if (action === "registrarTicketControlado") {
    var chequeoHora = validarHoraCierreApuestas();
    if (!chequeoHora.permitido) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, mensaje: chequeoHora.mensaje })).setMimeType(ContentService.MimeType.JSON);
    }

    var montoReg = parseFloat(params.monto) || 0;
    if (montoReg <= 0) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, mensaje: "Monto inválido." })).setMimeType(ContentService.MimeType.JSON);
    }

    var combinacionJugada = String(params.combinacion || "").trim();
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var hojaAjustes = ss.getSheetByName("Ajustes_Banca");
    var totalApuestasAcumuladas = hojaAjustes ? (parseFloat(hojaAjustes.getRange("I2").getValue()) || 0) : 0;

    var fondoTotalPremios = obtenerFondoBanca();
    var cupoMaximoApuestasGlobal = fondoTotalPremios / MULTIPLICADOR_PREMIO;

    if ((totalApuestasAcumuladas + montoReg) > cupoMaximoApuestasGlobal) {
      var cupoDisponible = Math.max(0, cupoMaximoApuestasGlobal - totalApuestasAcumuladas);
      resultado = { success: false, mensaje: `Límite global alcanzado. Cupo disponible: Bs ${cupoDisponible.toFixed(2)}.` };
    } else {
      var hojaTickets = ss.getSheetByName("Tickets");
      var dataTickets = hojaTickets ? hojaTickets.getDataRange().getValues() : [];
      var chequeoRiesgoDuplicados = verificarFraccionamientoYDuplicadosConDatos(dataTickets, combinacionJugada, montoReg);
      
      if (chequeoRiesgoDuplicados.requiereFraccionamiento) {
        return ContentService.createTextOutput(JSON.stringify({ 
          success: false, 
          requiereFraccionamiento: true, 
          detallesAfectados: chequeoRiesgoDuplicados.detallesAfectados,
          mensaje: chequeoRiesgoDuplicados.mensaje 
        })).setMimeType(ContentService.MimeType.JSON);
      }

      var ticketId = "TICK-" + Utilities.getUuid().substring(0, 6).toUpperCase();
      var estadoPago = params.pagado === "true" ? "Pagado" : "Pendiente";
      var premioPotencialCalculado = montoReg * MULTIPLICADOR_PREMIO;
      var referenciaPago = params.referencia || "";
      
      hojaTickets.appendRow([
        ticketId, params.vendedorId, params.jugador || "", combinacionJugada,
        montoReg, premioPotencialCalculado, estadoPago, params.banco || "",
        params.telefono || "", params.cedulaPago || params.cedula || "", referenciaPago
      ]);
      
      var datosToken = generarTokenAccesoJugador(ticketId, referenciaPago);
      recalcularCajasVendedores();
      
      resultado = { success: true, ticketId: ticketId, tokenAcceso: datosToken.token, enlaceAcceso: datosToken.enlace, mensaje: "Boleto registrado con éxito." };
    }
  } else if (action === "reiniciarJornada") {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var zonaHoraria = Session.getScriptTimeZone();
    var fechaCorta = Utilities.formatDate(new Date(), zonaHoraria, "yyyy-MM-dd");
    
    var sheetCarreras = ss.getSheetByName("Carreras");
    var sheetTickets = ss.getSheetByName("Tickets");
    var sheetCaja = ss.getSheetByName("Caja_Vendedores");
    var sheetMetricas = ss.getSheetByName("Historial_Metricas");
    var sheetResultados = ss.getSheetByName("Resultados");
    
    if (sheetCaja && sheetCaja.getLastRow() > 1) {
      var dataCajaHist = sheetCaja.getDataRange().getValues();
      for (var cIdx = 1; cIdx < dataCajaHist.length; cIdx++) {
        var taquillaVal = dataCajaHist[cIdx][0];
        if (taquillaVal) {
          sheetMetricas.appendRow([
            fechaCorta, taquillaVal, dataCajaHist[cIdx][1] || 0, dataCajaHist[cIdx][2] || 0,
            dataCajaHist[cIdx][3] || 0, dataCajaHist[cIdx][4] || 0, dataCajaHist[cIdx][5] || 0,
            dataCajaHist[cIdx][6] || "", dataCajaHist[cIdx][7] || "Cerrado"
          ]);
        }
      }
    }
    
    if (sheetCarreras && sheetCarreras.getLastRow() > 1) sheetCarreras.getRange(2, 1, sheetCarreras.getLastRow() - 1, sheetCarreras.getLastColumn()).clearContent();
    if (sheetTickets && sheetTickets.getLastRow() > 1) sheetTickets.getRange(2, 1, sheetTickets.getLastRow() - 1, sheetTickets.getLastColumn()).clearContent();
    if (sheetCaja && sheetCaja.getLastRow() > 1) sheetCaja.getRange(2, 1, sheetCaja.getLastRow() - 1, sheetCaja.getLastColumn()).clearContent();
    if (sheetResultados && sheetResultados.getLastRow() > 1) sheetResultados.getRange(2, 1, sheetResultados.getLastRow() - 1, sheetResultados.getLastColumn()).clearContent();
    
    resultado = { success: true, message: "¡Jornada reiniciada y métricas archivadas!" };
  } else if (action === "getMetricas") {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetMetricas = ss.getSheetByName("Historial_Metricas");
    var sheetCaja = ss.getSheetByName("Caja_Vendedores");
    var listMetricas = [];
    var zonaHoraria = Session.getScriptTimeZone();
    var hoyStr = Utilities.formatDate(new Date(), zonaHoraria, "dd/MM/yyyy");
    
    var globalTotalVentas = 0;
    var globalTotalPremios = 0;
    var globalComisiones = 0;
    var globalGananciaNeta = 0;
    
    if (sheetMetricas) {
      globalTotalVentas = parseFloat(sheetMetricas.getRange("K2").getValue()) || 0;
      globalTotalPremios = parseFloat(sheetMetricas.getRange("L2").getValue()) || 0;
      globalComisiones = parseFloat(sheetMetricas.getRange("M2").getValue()) || 0;
      globalGananciaNeta = parseFloat(sheetMetricas.getRange("N2").getValue()) || 0;
    }
    
    function formatearAVisualRobusto(fechaRaw) {
      if (!fechaRaw) return hoyStr;
      if (fechaRaw instanceof Date) return Utilities.formatDate(fechaRaw, zonaHoraria, "dd/MM/yyyy");
      var fStr = String(fechaRaw).trim();
      if (!fStr || fStr.toLowerCase() === "n/d") return hoyStr;
      if (fStr.indexOf('T') !== -1) fStr = fStr.split('T')[0];
      if (fStr.indexOf('-') !== -1) {
        var p = fStr.split('-');
        if (p.length === 3) {
          if (p[0].length === 4) return p[2] + "/" + p[1] + "/" + p[0];
          else if (p[2].length === 4) return p[0] + "/" + p[1] + "/" + p[2];
        }
      }
      return fStr;
    }

    if (sheetMetricas && sheetMetricas.getLastRow() > 1) {
      var dataMet = sheetMetricas.getDataRange().getValues();
      for (var mIdx = 1; mIdx < dataMet.length; mIdx++) {
        if (dataMet[mIdx][0] || dataMet[mIdx][1]) {
          var valVentas = parseFloat(dataMet[mIdx][2]) || 0;
          var valPremios = parseFloat(dataMet[mIdx][3]) || 0;
          var valQueda = parseFloat(dataMet[mIdx][4]) || 0;
          var valComision = parseFloat(dataMet[mIdx][5]) || 0;
          var valNeto = parseFloat(dataMet[mIdx][6]) || 0;
          
          listMetricas.push({
            fecha: formatearAVisualRobusto(dataMet[mIdx][0]),
            taquilla: dataMet[mIdx][1], vendedor: dataMet[mIdx][1],
            ventas: valVentas, total_ventas: valVentas,
            premios: valPremios, total_premios: valPremios,
            laQueda: valQueda, la_queda: valQueda,
            comision: valComision, neto: valNeto, neto_sistema: valNeto,
            referencia: dataMet[mIdx][7] || "", estatus: dataMet[mIdx][8] || "Cerrado"
          });
        }
      }
    }
    
    if (sheetCaja && sheetCaja.getLastRow() > 1) {
      var dataCaja = sheetCaja.getDataRange().getValues();
      for (var cIdx = 1; cIdx < dataCaja.length; cIdx++) {
        if (dataCaja[cIdx][0]) {
          var valVentasV = parseFloat(dataCaja[cIdx][1]) || 0;
          var valPremiosV = parseFloat(dataCaja[cIdx][2]) || 0;
          var valQuedaV = parseFloat(dataCaja[cIdx][3]) || 0;
          var valComisionV = parseFloat(dataCaja[cIdx][4]) || 0;
          var valNetoV = parseFloat(dataCaja[cIdx][5]) || 0;

          listMetricas.push({
            fecha: hoyStr,
            taquilla: dataCaja[cIdx][0], vendedor: dataCaja[cIdx][0],
            ventas: valVentasV, total_ventas: valVentasV,
            premios: valPremiosV, total_premios: valPremiosV,
            laQueda: valQuedaV, la_queda: valQuedaV,
            comision: valComisionV, neto: valNetoV, neto_sistema: valNetoV,
            referencia: dataCaja[cIdx][6] || "", estatus: dataCaja[cIdx][7] || "PENDIENTE"
          });
        }
      }
    }
    
    resultado = { 
      success: true, 
      data: listMetricas,
      global: {
        totalVentas: globalTotalVentas,
        totalPremios: globalTotalPremios,
        comisiones: globalComisiones,
        gananciaNeta: globalGananciaNeta
      }
    };
  } else if (action === "getUsuarios") {
    var dataUsr = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios").getDataRange().getValues();
    var listUsuarios = [];
    for (var u = 1; u < dataUsr.length; u++) {
      if (dataUsr[u][1]) {
        listUsuarios.push({ id: dataUsr[u][0], usuario: dataUsr[u][1], password: dataUsr[u][2], rol: dataUsr[u][3], estado: dataUsr[u][5] || "Activo" });
      }
    }
    resultado = { success: true, data: listUsuarios };
  } else if (action === "guardarUsuario") {
    var sheetUsr = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
    var dataUsr = sheetUsr.getDataRange().getValues();
    var encontradaFila = -1;
    for (var f = 1; f < dataUsr.length; f++) {
      if (String(dataUsr[f][0]).trim() === String(params.editId).trim()) {
        encontradaFila = f + 1;
        break;
      }
    }
    if (encontradaFila !== -1) {
      sheetUsr.getRange(encontradaFila, 2, 1, 4).setValues([[params.usuario, params.password, params.rol, params.estado || "Activo"]]);
      resultado = { success: true, message: "¡Usuario actualizado!" };
    } else {
      sheetUsr.appendRow(["MAFRAS-" + String(dataUsr.length).padStart(2, '0'), params.usuario, params.password, params.rol, "", params.estado || "Activo"]);
      resultado = { success: true, message: "¡Usuario creado!" };
    }
  } else if (action === "cambiarEstadoUsuario") {
    var sheetUsr = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
    var dataUsr = sheetUsr.getDataRange().getValues();
    var encontrado = false;
    for (var e = 1; e < dataUsr.length; e++) {
      if (String(dataUsr[e][0]).trim() === String(params.id).trim()) {
        sheetUsr.getRange(e + 1, 6).setValue(params.estado);
        encontrado = true;
        break;
      }
    }
    resultado = encontrado ? { success: true, message: "¡Estatus actualizado!" } : { success: false, message: "Usuario no encontrado." };
  } else if (action === "getCarreras") {
    var dataCar = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Carreras").getDataRange().getValues();
    var listCarreras = [];
    var zonaHorariaCar = Session.getScriptTimeZone();
    for (var k = 1; k < dataCar.length; k++) {
      if (dataCar[k][0] || dataCar[k][1]) {
        var fechaRaw = dataCar[k][5];
        var fechaStr = "";
        if (fechaRaw instanceof Date) {
          fechaStr = Utilities.formatDate(fechaRaw, zonaHorariaCar, "yyyy-MM-dd");
        } else {
          var fVal = String(fechaRaw || "").trim();
          if (fVal.indexOf('T') !== -1) fVal = fVal.split('T')[0];
          fechaStr = fVal;
        }
        
        listCarreras.push({
          carrera: dataCar[k][0], numero: dataCar[k][1], ejemplar: dataCar[k][2],
          estado: dataCar[k][3], hipodromo: dataCar[k][4], fecha: fechaStr,
          hora: String(dataCar[k][6] || ""), distancia: dataCar[k][7], reunion: String(dataCar[k][8] || "")
        });
      }
    }
    resultado = { success: true, data: listCarreras };
  } else if (action === "getResultados") {
    var dataRes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Resultados").getDataRange().getValues();
    var listRes = [];
    for (var r = 1; r < dataRes.length; r++) {
      if (dataRes[r][0]) {
        listRes.push({ Valida: dataRes[r][0], CaballoGanador: dataRes[r][1], '2do lugar': dataRes[r][2], '3er lugar': dataRes[r][3], puestosValidos: dataRes[r][4] || "" });
      }
    }
    resultado = { success: true, data: listRes };
  } else if (action === "getCajas") {
    recalcularCajasVendedores();
    var dataCaja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Caja_Vendedores").getDataRange().getValues();
    var listCajas = [];
    for (var m = 1; m < dataCaja.length; m++) {
      if (dataCaja[m][0]) {
        listCajas.push({
          taquilla: dataCaja[m][0], ventas: dataCaja[m][1], premios: dataCaja[m][2],
          laQueda: dataCaja[m][3], comision: dataCaja[m][4], neto: dataCaja[m][5],
          referencia: dataCaja[m][6] || "", estatus: dataCaja[m][7] || "PENDIENTE"
        });
      }
    }
    resultado = { success: true, data: listCajas };
  } else {
    resultado = { status: "API Activa y sincronizada" };
  }

  var jsonString = JSON.stringify(resultado);
  if (callback) {
    jsonString = callback + "(" + jsonString + ");";
    return ContentService.createTextOutput(jsonString).setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(jsonString).setMimeType(ContentService.MimeType.JSON);
}

function verificarLogin(usuario, password) {
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios").getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(usuario).trim() && String(data[i][2]).trim() === String(password).trim()) {
      if (String(data[i][5] || "Activo").trim().toLowerCase() === "suspendido") {
        return { success: false, message: "Su cuenta se encuentra suspendida." };
      }
      return { success: true, rol: data[i][3], usuario: data[i][1] };
    }
  }
  return { success: false, message: "Usuario o contraseña incorrectos." };
}

function procesarResultadosYCalcularGanadores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetResultados = ss.getSheetByName("Resultados");
  var sheetTickets = ss.getSheetByName("Tickets");
  var sheetCarreras = ss.getSheetByName("Carreras");
  
  var resultadosData = sheetResultados.getDataRange().getValues();
  var ticketsData = sheetTickets.getDataRange().getValues();
  var carrerasData = sheetCarreras.getDataRange().getValues();
  
  var resultadosOficiales = {};
  var carrerasOficialesCorridas = {}; 
  
  var ventasPorCaballoPorValida = {};
  for (var t = 1; t < ticketsData.length; t++) {
    var estatusT = String(ticketsData[t][6]).trim().toLowerCase();
    if (estatusT === "anulado") continue;
    var selStr = String(ticketsData[t][3]).trim();
    if (!selStr) continue;
    var montoT = parseFloat(ticketsData[t][4]) || 0;
    
    var subJugadas = selStr.includes(" | ") ? selStr.split(" | ") : [selStr];
    var montoPorSub = montoT / subJugadas.length;
    
    for (var sj = 0; sj < subJugadas.length; sj++) {
      var item = subJugadas[sj].trim();
      var combPart = item;
      var montoSub = montoPorSub;
      var matchSub = item.match(/(.*?)\s*\(Bs\s*([\d\.]+)\)/);
      if (matchSub) {
        combPart = matchSub[1].trim();
        montoSub = parseFloat(matchSub[2]) || 0;
      }
      var hElegidos = combPart.split(" - ");
      
      for (var v = 0; v < 6; v++) {
        var valKey = "Valida " + (v + 1);
        var hNum = parseInt(hElegidos[v]);
        if (!isNaN(hNum)) {
          if (!ventasPorCaballoPorValida[valKey]) ventasPorCaballoPorValida[valKey] = {};
          ventasPorCaballoPorValida[valKey][hNum] = (ventasPorCaballoPorValida[valKey][hNum] || 0) + montoSub;
        }
      }
    }
  }

  var caballosRetirados = {};
  var caballosActivosPorValida = {};
  for (var c = 1; c < carrerasData.length; c++) {
    var valCar = String(carrerasData[c][0]).trim();
    var numCar = parseInt(carrerasData[c][1]);
    var estCar = String(carrerasData[c][3]).trim().toLowerCase();
    if (valCar && !isNaN(numCar)) {
      if (estCar === "retirado") {
        if (!caballosRetirados[valCar]) caballosRetirados[valCar] = {};
        caballosRetirados[valCar][numCar] = true;
      } else {
        if (!caballosActivosPorValida[valCar]) caballosActivosPorValida[valCar] = [];
        caballosActivosPorValida[valCar].push(numCar);
      }
    }
  }

  var favoritosPorValida = {};
  for (var valKey in caballosActivosPorValida) {
    var activos = caballosActivosPorValida[valKey];
    if (activos.length === 0) continue;
    var ventasH = ventasPorCaballoPorValida[valKey] || {};
    var maxVentas = -1;
    var mejorCaballo = activos[0];
    for (var i = 0; i < activos.length; i++) {
      var hNum = activos[i];
      var vAcum = ventasH[hNum] || 0;
      if (vAcum > maxVentas) {
        maxVentas = vAcum;
        mejorCaballo = hNum;
      }
    }
    favoritosPorValida[valKey] = mejorCaballo;
  }
  
  for (var r = 1; r < resultadosData.length; r++) {
    if (resultadosData[r][0]) {
      var valName = resultadosData[r][0]; 
      var ganadoresVal = [];
      var cuantosPuestos = parseInt(resultadosData[r][4]) || 3;
      
      var tieneGanadorReal = (resultadosData[r][1] !== "" && resultadosData[r][1] !== null && resultadosData[r][1] !== undefined);
      
      if (tieneGanadorReal) {
        for (var col = 1; col <= cuantosPuestos; col++) {
          if (resultadosData[r][col] !== "" && resultadosData[r][col] !== null && resultadosData[r][col] !== undefined) {
            ganadoresVal.push(parseInt(resultadosData[r][col]));
          }
        }
        resultadosOficiales[valName] = ganadoresVal;
        carrerasOficialesCorridas[valName] = true;
      } else {
        carrerasOficialesCorridas[valName] = false;
      }
      
      for (var c = 1; c < carrerasData.length; c++) {
        if (String(carrerasData[c][0]).trim() === valName) {
          if (String(carrerasData[c][3]).trim().toLowerCase() !== "retirado") {
            var numCarrera = parseInt(carrerasData[c][1]);
            if (tieneGanadorReal) {
              sheetCarreras.getRange(c + 1, 4).setValue(ganadoresVal.indexOf(numCarrera) !== -1 ? "GANO" : "CORRIO");
            } else {
              sheetCarreras.getRange(c + 1, 4).setValue("PENDIENTE");
            }
          }
        }
      }
    }
  }
  
  for (var t = 1; t < ticketsData.length; t++) {
    var estatusActual = String(ticketsData[t][6]).trim();
    if (estatusActual === "Pagado" || estatusActual === "Anulado") continue; 
    
    var seleccionStr = String(ticketsData[t][3]).trim();
    if (!seleccionStr) continue;
    
    var subJugadas = seleccionStr.includes(" | ") ? seleccionStr.split(" | ") : [seleccionStr];
    var subJugadasResultados = [];
    
    for (var sj = 0; sj < subJugadas.length; sj++) {
      var item = subJugadas[sj].trim();
      var combPart = item;
      var matchSub = item.match(/(.*?)\s*\(Bs\s*([\d\.]+)\)/);
      if (matchSub) {
        combPart = matchSub[1].trim();
      }
      var caballosElegidos = combPart.split(" - ");
      
      var subPerdido = false;
      var subTodasCorrieronYAcertadas = true;
      
      for (var v = 0; v < 6; v++) {
        var validaKey = "Valida " + (v + 1);
        var seleccionCaballo = parseInt(caballosElegidos[v]);
        
        if (caballosRetirados[validaKey] && caballosRetirados[validaKey][seleccionCaballo]) {
          var favoritoValida = favoritosPorValida[validaKey];
          if (favoritoValida !== undefined) {
            seleccionCaballo = favoritoValida;
          }
        }
        
        if (carrerasOficialesCorridas[validaKey]) {
          var ganadoresValida = resultadosOficiales[validaKey] || [];
          if (ganadoresValida.indexOf(seleccionCaballo) === -1) {
            subPerdido = true;
            break; 
          }
        } else {
          subTodasCorrieronYAcertadas = false;
        }
      }
      
      if (subPerdido) {
        subJugadasResultados.push("Perdedor");
      } else if (subTodasCorrieronYAcertadas) {
        subJugadasResultados.push("Ganador");
      } else {
        subJugadasResultados.push("Pendiente");
      }
    }
    
    var todasPerdidas = subJugadasResultados.every(function(st) { return st === "Perdedor"; });
    var algunaGanadora = subJugadasResultados.some(function(st) { return st === "Ganador"; });
    var algunaPendiente = subJugadasResultados.some(function(st) { return st === "Pendiente"; });
    
    var celdaEstatusTicket = sheetTickets.getRange(t + 1, 7);
    if (todasPerdidas) {
      celdaEstatusTicket.setValue("Perdedor");
    } else if (algunaPendiente) {
      celdaEstatusTicket.setValue("Pendiente");
    } else if (algunaGanadora) {
      celdaEstatusTicket.setValue("Ganador");
    } else {
      celdaEstatusTicket.setValue("Pendiente");
    }
  }
  recalcularCajasVendedores();
}

function recalcularCajasVendedores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetTickets = ss.getSheetByName("Tickets");
  var sheetCaja = ss.getSheetByName("Caja_Vendedores");
  var sheetCarreras = ss.getSheetByName("Carreras");
  var sheetResultados = ss.getSheetByName("Resultados");
  
  if (!sheetTickets || !sheetCaja) return;
  
  var ticketsData = sheetTickets.getDataRange().getValues();
  var cajaData = sheetCaja.getDataRange().getValues();
  var carrerasData = sheetCarreras ? sheetCarreras.getDataRange().getValues() : [];
  var resultadosData = sheetResultados ? sheetResultados.getDataRange().getValues() : [];
  
  var resultadosOficiales = {};
  var carrerasOficialesCorridas = {};
  for (var r = 1; r < resultadosData.length; r++) {
    if (resultadosData[r][0]) {
      var valName = resultadosData[r][0];
      var ganadoresVal = [];
      var cuantosPuestos = parseInt(resultadosData[r][4]) || 3;
      var tieneGanadorReal = (resultadosData[r][1] !== "" && resultadosData[r][1] !== null && resultadosData[r][1] !== undefined);
      if (tieneGanadorReal) {
        for (var col = 1; col <= cuantosPuestos; col++) {
          if (resultadosData[r][col] !== "" && resultadosData[r][col] !== null && resultadosData[r][col] !== undefined) {
            ganadoresVal.push(parseInt(resultadosData[r][col]));
          }
        }
        resultadosOficiales[valName] = ganadoresVal;
        carrerasOficialesCorridas[valName] = true;
      }
    }
  }

  var caballosRetirados = {};
  var ventasPorCaballoPorValida = {};
  for (var t = 1; t < ticketsData.length; t++) {
    var estT = String(ticketsData[t][6]).trim().toLowerCase();
    if (estT === "anulado") continue;
    var selStr = String(ticketsData[t][3]).trim();
    if (!selStr) continue;
    var montoT = parseFloat(ticketsData[t][4]) || 0;
    var subJugadas = selStr.includes(" | ") ? selStr.split(" | ") : [selStr];
    var montoPorSub = montoT / subJugadas.length;
    for (var sj = 0; sj < subJugadas.length; sj++) {
      var item = subJugadas[sj].trim();
      var combPart = item;
      var matchSub = item.match(/(.*?)\s*\(Bs\s*([\d\.]+)\)/);
      if (matchSub) combPart = matchSub[1].trim();
      var hElegidos = combPart.split(" - ");
      for (var v = 0; v < 6; v++) {
        var valKey = "Valida " + (v + 1);
        var hNum = parseInt(hElegidos[v]);
        if (!isNaN(hNum)) {
          if (!ventasPorCaballoPorValida[valKey]) ventasPorCaballoPorValida[valKey] = {};
          ventasPorCaballoPorValida[valKey][hNum] = (ventasPorCaballoPorValida[valKey][hNum] || 0) + montoPorSub;
        }
      }
    }
  }

  for (var c = 1; c < carrerasData.length; c++) {
    var valCar = String(carrerasData[c][0]).trim();
    var numCar = parseInt(carrerasData[c][1]);
    var estCar = String(carrerasData[c][3]).trim().toLowerCase();
    if (valCar && !isNaN(numCar) && estCar === "retirado") {
      if (!caballosRetirados[valCar]) caballosRetirados[valCar] = {};
      caballosRetirados[valCar][numCar] = true;
    }
  }

  var caballosActivosPorValida = {};
  for (var c = 1; c < carrerasData.length; c++) {
    var valCar = String(carrerasData[c][0]).trim();
    var numCar = parseInt(carrerasData[c][1]);
    var estCar = String(carrerasData[c][3]).trim().toLowerCase();
    if (valCar && !isNaN(numCar) && estCar !== "retirado") {
      if (!caballosActivosPorValida[valCar]) caballosActivosPorValida[valCar] = [];
      caballosActivosPorValida[valCar].push(numCar);
    }
  }

  var favoritosPorValida = {};
  for (var valKey in caballosActivosPorValida) {
    var activos = caballosActivosPorValida[valKey];
    if (activos.length === 0) continue;
    var ventasH = ventasPorCaballoPorValida[valKey] || {};
    var maxVentas = -1;
    var mejorCaballo = activos[0];
    for (var i = 0; i < activos.length; i++) {
      var hNum = activos[i];
      var vAcum = ventasH[hNum] || 0;
      if (vAcum > maxVentas) {
        maxVentas = vAcum;
        mejorCaballo = hNum;
      }
    }
    favoritosPorValida[valKey] = mejorCaballo;
  }
  
  var estadosPrevios = {};
  for (var c = 1; c < cajaData.length; c++) {
    var taqKey = String(cajaData[c][0]).trim().toLowerCase();
    if (taqKey) {
      estadosPrevios[taqKey] = {
        referencia: cajaData[c][6] || "",
        estatus: cajaData[c][7] || "PENDIENTE"
      };
    }
  }
  
  if (cajaData.length > 1) {
    sheetCaja.getRange(2, 1, cajaData.length - 1, sheetCaja.getLastColumn()).clearContent();
  }
  
  if (ticketsData.length <= 1) return;
  
  var resumenVendedores = {};
  for (var i = 1; i < ticketsData.length; i++) {
    var vendedor = ticketsData[i][1];
    var montoVenta = parseFloat(ticketsData[i][4]) || 0;
    var estatusTicket = String(ticketsData[i][6]).trim();
    
    if (!vendedor) continue;
    if (!resumenVendedores[vendedor]) { resumenVendedores[vendedor] = { ventas: 0, premiosPagados: 0 }; }
    
    resumenVendedores[vendedor].ventas += montoVenta;
    
    if (estatusTicket === "Ganador" || estatusTicket === "Ganadores" || estatusTicket === "Pagado") {
      var selStr = String(ticketsData[i][3]).trim();
      var subJugadas = selStr.includes(" | ") ? selStr.split(" | ") : [selStr];
      var montoPorSub = montoVenta / subJugadas.length;
      
      for (var sj = 0; sj < subJugadas.length; sj++) {
        var item = subJugadas[sj].trim();
        var combPart = item;
        var montoSub = montoPorSub;
        var matchSub = item.match(/(.*?)\s*\(Bs\s*([\d\.]+)\)/);
        if (matchSub) {
          combPart = matchSub[1].trim();
          montoSub = parseFloat(matchSub[2]) || 0;
        }
        var caballosElegidos = combPart.split(" - ");
        var subGanadora = true;
        
        for (var v = 0; v < 6; v++) {
          var validaKey = "Valida " + (v + 1);
          var seleccionCaballo = parseInt(caballosElegidos[v]);
          
          if (caballosRetirados[validaKey] && caballosRetirados[validaKey][seleccionCaballo]) {
            var favoritoValida = favoritosPorValida[validaKey];
            if (favoritoValida !== undefined) seleccionCaballo = favoritoValida;
          }
          
          if (carrerasOficialesCorridas[validaKey]) {
            var ganadoresValida = resultadosOficiales[validaKey] || [];
            if (ganadoresValida.indexOf(seleccionCaballo) === -1) {
              subGanadora = false;
              break;
            }
          } else {
            subGanadora = false;
          }
        }
        
        if (subGanadora) {
          resumenVendedores[vendedor].premiosPagados += (montoSub * MULTIPLICADOR_PREMIO);
        }
      }
    }
  }
  
  var filaDestino = 2;
  var filasAEscribir = [];
  for (var v in resumenVendedores) {
    var totalVentas = resumenVendedores[v].ventas;
    var totalPremios = resumenVendedores[v].premiosPagados;
    var laQueda = totalVentas - totalPremios;
    var comision = laQueda > 0 ? laQueda * COMISION_QUEDA : 0;
    var netoEntregar = laQueda > 0 ? laQueda * 0.70 : laQueda;
    
    var vKeyNormalized = String(v).trim().toLowerCase();
    var refAnterior = (estadosPrevios[vKeyNormalized] && estadosPrevios[vKeyNormalized].referencia) ? estadosPrevios[vKeyNormalized].referencia : "";
    var estAnterior = (estadosPrevios[vKeyNormalized] && estadosPrevios[vKeyNormalized].estatus) ? estadosPrevios[vKeyNormalized].estatus : "PENDIENTE";
    
    filasAEscribir.push([v, totalVentas, totalPremios, laQueda, comision, netoEntregar, refAnterior, estAnterior]);
  }
  
  if (filasAEscribir.length > 0) {
    sheetCaja.getRange(2, 1, filasAEscribir.length, filasAEscribir[0].length).setValues(filasAEscribir);
  }
}