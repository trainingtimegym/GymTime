/**
 * SISTEMA DE GESTIÓN DE GIMNASIO
 * Backend Google Apps Script
 * Hojas requeridas: Socios, Pagos, Config
 * Config debe tener UNA sola fila: Mensual | 30 | 30000
 */

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function doGet(e) {
  return manejarPeticion(e);
}
function doPost(e) {
  const params = JSON.parse(e.postData.contents);
  return manejarPeticion({ parameter: params });
}

function manejarPeticion(e) {
  try {
    const accion = e.parameter.accion;
    let resultado;

    switch (accion) {
      case 'buscarSocio':
        resultado = buscarSocio(e.parameter.query);
        break;
      case 'altaSocio':
        resultado = altaSocio(JSON.parse(e.parameter.datos));
        break;
      case 'registrarPago':
        resultado = registrarPago(JSON.parse(e.parameter.datos));
        break;
      case 'cambiarEstado':
        resultado = cambiarEstado(e.parameter.dni, e.parameter.nuevoEstado);
        break;
      case 'listarPorVencer':
        resultado = listarPorVencer();
        break;
      case 'listarTodos':
        resultado = listarTodosLosSocios();
        break;
      case 'obtenerConfiguracion':
        resultado = obtenerConfiguracion();
        break;
      case 'actualizarConfiguracion':
        resultado = actualizarConfiguracion(JSON.parse(e.parameter.datos));
        break;
      case 'historialPagos':
        resultado = historialPagos(e.parameter.dni);
        break;
      default:
        resultado = { error: 'Acción no reconocida' };
    }

    return ContentService.createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---------- HELPERS ----------

function getSheet(nombre) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(nombre);
}

function normalizar(txt) {
  return (txt || '').toString().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
}

function calcularEstadoSemaforo(fechaVencimiento) {
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  const venc = new Date(fechaVencimiento);
  venc.setHours(0, 0, 0, 0);
  const diffMs = venc - hoy;
  const diasRestantes = Math.round(diffMs / (1000 * 60 * 60 * 24));

  let color;
  if (diasRestantes < 0) color = 'vencido';
  else if (diasRestantes <= 5) color = 'rojo';
  else if (diasRestantes <= 15) color = 'amarillo';
  else color = 'verde';

  return { diasRestantes, color };
}

function mensajeAliento(diasRestantes) {
  const frasesVerde = [
    '¡Seguís firme! Así se construyen resultados 💪',
    'Constancia es la clave, ¡vamos con todo!',
    '¡A puro entrenamiento! Nos vemos en el gym.'
  ];
  const frasesAmarillo = [
    'Se acerca la fecha de pago, ¡no perdés el ritmo!',
    'Ya casi llega tu próximo pago, seguí entrenando fuerte.'
  ];
  const frasesRojo = [
    'Tu pago está por vencer, ¡no dejes que se corte la racha!',
    '¡Últimos días! Renová y seguí sumando entrenamientos.'
  ];
  const frasesVencido = [
    'Tu cuota venció, te esperamos para renovar y volver a entrenar.',
    'Se venció tu plan, ¡extrañamos verte entrenar! Renová cuando quieras.'
  ];

  let lista;
  if (diasRestantes < 0) lista = frasesVencido;
  else if (diasRestantes <= 5) lista = frasesRojo;
  else if (diasRestantes <= 15) lista = frasesAmarillo;
  else lista = frasesVerde;

  return lista[Math.floor(Math.random() * lista.length)];
}

function generarLinkWhatsapp(celular, nombre, diasRestantes, fechaVenc, monto) {
  let cel = (celular || '').toString().replace(/\D/g, '');
  if (cel && !cel.startsWith('54')) cel = '54' + cel;

  const montoTxt = monto ? ` (cuota: $${Number(monto).toLocaleString('es-AR')})` : '';
  let textoEstado;
  if (diasRestantes < 0) {
    textoEstado = `tu cuota venció hace ${Math.abs(diasRestantes)} día(s) (${fechaVenc})${montoTxt}`;
  } else if (diasRestantes === 0) {
    textoEstado = `tu cuota vence hoy (${fechaVenc})${montoTxt}`;
  } else {
    textoEstado = `te quedan ${diasRestantes} día(s) de entrenamiento (vence el ${fechaVenc})${montoTxt}`;
  }

  const mensaje = `¡Hola ${nombre}! 👋 Te escribimos del gym: ${textoEstado}. ${mensajeAliento(diasRestantes)}`;
  const url = `https://wa.me/${cel}?text=${encodeURIComponent(mensaje)}`;
  return url;
}

// ---------- CONFIGURACIÓN (plan único) ----------

function obtenerConfiguracion() {
  const sheet = getSheet('Config');
  const datos = sheet.getDataRange().getValues();
  if (datos.length < 2) return { plan: 'Mensual', duracionDias: 30, precio: 0 };
  const fila = datos[1];
  return { plan: fila[0] || 'Mensual', duracionDias: fila[1] || 30, precio: fila[2] || 0 };
}

function actualizarConfiguracion(d) {
  const sheet = getSheet('Config');
  if (sheet.getLastRow() < 2) {
    sheet.appendRow(['Mensual', d.duracionDias || 30, d.precio || 0]);
  } else {
    if (d.duracionDias !== undefined) sheet.getRange(2, 2).setValue(d.duracionDias);
    if (d.precio !== undefined) sheet.getRange(2, 3).setValue(d.precio);
  }
  return { ok: true };
}

// ---------- SOCIOS ----------

function buscarSocio(query) {
  const sheet = getSheet('Socios');
  const datos = sheet.getDataRange().getValues();
  const config = obtenerConfiguracion();
  const q = normalizar(query);
  const resultados = [];

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const dni = fila[0].toString();
    const nombre = fila[1];
    const apellido = fila[2];
    const celular = fila[3].toString();
    const nombreCompleto = normalizar(nombre + ' ' + apellido);

    if (dni.includes(q) || celular.includes(q) || nombreCompleto.includes(q)) {
      resultados.push(filaASocio(fila, config));
    }
  }
  return { resultados };
}

function filaASocio(fila, config) {
  const [dni, nombre, apellido, celular, fechaAlta, plan, duracionDias, estado, fechaUltimoPago, fechaVencimiento, notas] = fila;
  const semaforo = fechaVencimiento ? calcularEstadoSemaforo(fechaVencimiento) : { diasRestantes: null, color: 'sin_pagos' };
  const fechaVencTxt = fechaVencimiento ? Utilities.formatDate(new Date(fechaVencimiento), 'GMT-3', 'dd/MM/yyyy') : '';

  return {
    dni: dni.toString(),
    nombre, apellido, celular: celular.toString(),
    fechaAlta: fechaAlta ? Utilities.formatDate(new Date(fechaAlta), 'GMT-3', 'dd/MM/yyyy') : '',
    plan: plan || config.plan,
    duracionDias: duracionDias || config.duracionDias,
    estado,
    fechaUltimoPago: fechaUltimoPago ? Utilities.formatDate(new Date(fechaUltimoPago), 'GMT-3', 'dd/MM/yyyy') : '',
    fechaVencimiento: fechaVencTxt,
    notas: notas || '',
    diasRestantes: semaforo.diasRestantes,
    colorSemaforo: semaforo.color,
    montoCuota: config.precio,
    linkWhatsapp: (celular && fechaVencimiento) ? generarLinkWhatsapp(
      celular, nombre, semaforo.diasRestantes, fechaVencTxt, config.precio
    ) : null
  };
}

function altaSocio(d) {
  const sheet = getSheet('Socios');
  const config = obtenerConfiguracion();
  const dni = d.dni.toString();
  const datos = sheet.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0].toString() === dni) return { error: 'Ya existe un socio con ese DNI' };
  }
  sheet.appendRow([
    dni, d.nombre, d.apellido, d.celular, new Date(),
    config.plan, config.duracionDias, 'Activo', '', '', d.notas || ''
  ]);
  return { ok: true };
}

function cambiarEstado(dni, nuevoEstado) {
  const sheet = getSheet('Socios');
  const datos = sheet.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0].toString() === dni.toString()) {
      sheet.getRange(i + 1, 8).setValue(nuevoEstado);
      return { ok: true };
    }
  }
  return { error: 'Socio no encontrado' };
}

function listarTodosLosSocios() {
  const sheet = getSheet('Socios');
  const datos = sheet.getDataRange().getValues();
  const config = obtenerConfiguracion();
  const resultados = [];
  for (let i = 1; i < datos.length; i++) resultados.push(filaASocio(datos[i], config));
  return { resultados };
}

function listarPorVencer() {
  const sheet = getSheet('Socios');
  const datos = sheet.getDataRange().getValues();
  const config = obtenerConfiguracion();
  const resultados = [];

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const estado = fila[7];
    const fechaVenc = fila[9];
    if (estado !== 'Activo' || !fechaVenc) continue;

    const semaforo = calcularEstadoSemaforo(fechaVenc);
    if (semaforo.color === 'rojo' || semaforo.color === 'vencido') {
      resultados.push(filaASocio(fila, config));
    }
  }
  return { resultados };
}

// ---------- PAGOS ----------

function registrarPago(d) {
  const sheetSocios = getSheet('Socios');
  const sheetPagos = getSheet('Pagos');
  const config = obtenerConfiguracion();
  const dni = d.dni.toString();
  const datos = sheetSocios.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0].toString() === dni) {
      const hoy = new Date();
      const dias = parseInt(d.dias, 10) || config.duracionDias;
      const monto = d.monto !== undefined && d.monto !== '' ? d.monto : config.precio;
      const fechaVencimiento = new Date(hoy);
      fechaVencimiento.setDate(fechaVencimiento.getDate() + dias);

      sheetSocios.getRange(i + 1, 6).setValue(config.plan);
      sheetSocios.getRange(i + 1, 7).setValue(dias);
      sheetSocios.getRange(i + 1, 8).setValue('Activo');
      sheetSocios.getRange(i + 1, 9).setValue(hoy);
      sheetSocios.getRange(i + 1, 10).setValue(fechaVencimiento);

      sheetPagos.appendRow([
        Utilities.getUuid(), dni, hoy, monto, dias, fechaVencimiento, d.registradoPor || ''
      ]);

      return { ok: true, nuevoVencimiento: Utilities.formatDate(fechaVencimiento, 'GMT-3', 'dd/MM/yyyy') };
    }
  }
  return { error: 'Socio no encontrado' };
}

function historialPagos(dni) {
  const sheet = getSheet('Pagos');
  const datos = sheet.getDataRange().getValues();
  const resultados = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1].toString() === dni.toString()) {
      resultados.push({
        fecha: Utilities.formatDate(new Date(datos[i][2]), 'GMT-3', 'dd/MM/yyyy'),
        monto: datos[i][3],
        dias: datos[i][4]
      });
    }
  }
  return { resultados: resultados.reverse() };
}
