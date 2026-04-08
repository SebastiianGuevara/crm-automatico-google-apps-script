// ══════════════════════════════════════════════
//  MENÚ PERSONALIZADO
// ══════════════════════════════════════════════
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('📊 CRM')
      .addItem('🎨 Aplicar colores por estado', 'colorearClientes')
      .addItem('📧 Enviar alertas de seguimiento', 'enviarAlertas')
      .addItem('📈 Actualizar Dashboard', 'actualizarDashboard')
      .addItem('📝 Registrar seguimiento', 'registrarSeguimiento')
      .addSeparator()
      .addItem('⚙️ Configurar validaciones', 'configurarValidaciones')
      .addToUi();
  } catch (e) {
    Logger.log('Error en onOpen: ' + e.message);
  }
}


// ══════════════════════════════════════════════
//  COLOREAR FILAS SEGÚN ESTADO
// ══════════════════════════════════════════════
function colorearClientes() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clientes');
  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) return;

  for (let fila = 2; fila <= ultimaFila; fila++) {
    const estado = hoja.getRange(fila, 6).getValue(); // Columna F = Estado
    const rango = hoja.getRange(fila, 1, 1, 8);

    if (estado === 'Caliente') {
      rango.setBackground('#d4edda'); // Verde suave
      hoja.getRange(fila, 6).setFontColor('#155724').setFontWeight('bold');
    } else if (estado === 'Tibio') {
      rango.setBackground('#fff3cd'); // Amarillo suave
      hoja.getRange(fila, 6).setFontColor('#856404').setFontWeight('bold');
    } else if (estado === 'Frío') {
      rango.setBackground('#f8d7da'); // Rojo suave
      hoja.getRange(fila, 6).setFontColor('#721c24').setFontWeight('bold');
    }
  }

  SpreadsheetApp.getUi().alert('✅ Colores aplicados correctamente.');
}


// ══════════════════════════════════════════════
//  ENVIAR ALERTAS DE CLIENTES SIN CONTACTO
// ══════════════════════════════════════════════
function enviarAlertas() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clientes');
  const ultimaFila = hoja.getLastRow();
  const hoy = new Date();
  let alertas = [];

  for (let fila = 2; fila <= ultimaFila; fila++) {
    const nombre = hoja.getRange(fila, 1).getValue();
    const email = hoja.getRange(fila, 2).getValue();
    const ultimoContacto = hoja.getRange(fila, 7).getValue();

    if (!nombre || !ultimoContacto) continue;

    const fechaContacto = new Date(ultimoContacto);
    const diasSin = Math.floor((hoy - fechaContacto) / (1000 * 60 * 60 * 24));

    if (diasSin >= CONFIG.DIAS_SIN_CONTACTO) {
      alertas.push({ nombre, email, diasSin, fila });
    }
  }

  if (alertas.length === 0) {
    SpreadsheetApp.getUi().alert('✅ Todos los clientes tienen seguimiento reciente.');
    return;
  }

  // Construir el cuerpo del correo
  let cuerpoCorreo = `Hola,\n\nEstos clientes llevan más de ${CONFIG.DIAS_SIN_CONTACTO} días sin contacto:\n\n`;
  alertas.forEach(c => {
    cuerpoCorreo += `• ${c.nombre} — ${c.diasSin} días sin contacto\n`;
  });
  cuerpoCorreo += `\nTotal: ${alertas.length} clientes requieren seguimiento.\n\nSistema CRM Automático`;

  // Enviar correo al administrador
  GmailApp.sendEmail(
    CONFIG.EMAIL_ADMIN,
    `⚠️ CRM: ${alertas.length} clientes sin seguimiento`,
    cuerpoCorreo
  );

  // Registrar en la hoja Registro-Alertas
  const hojaAlertas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Registro-Alertas');
  alertas.forEach(c => {
    hojaAlertas.appendRow([new Date(), c.nombre, c.email, c.diasSin]);
  });

  // Colorear en rojo brillante las filas con alerta
  alertas.forEach(c => {
    hoja.getRange(c.fila, 1, 1, 8).setBackground('#f5c6cb');
  });

  SpreadsheetApp.getUi().alert(`⚠️ Alerta enviada. ${alertas.length} clientes sin contacto.`);
}


// ══════════════════════════════════════════════
//  ACTUALIZAR DASHBOARD
// ══════════════════════════════════════════════
function actualizarDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaClientes = ss.getSheetByName('Clientes');
  const hojaDashboard = ss.getSheetByName('Dashboard');
  const ultimaFila = hojaClientes.getLastRow();

  // Limpiar dashboard anterior
  hojaDashboard.clearContents();

  // Contar clientes por estado
  let calientes = 0, tibios = 0, frios = 0, total = 0;

  for (let fila = 2; fila <= ultimaFila; fila++) {
    const estado = hojaClientes.getRange(fila, 6).getValue();
    if (estado === 'Caliente') calientes++;
    else if (estado === 'Tibio') tibios++;
    else if (estado === 'Frío') frios++;
    if (estado) total++;
  }

  // Contar por ciudad
  const ciudades = {};
  for (let fila = 2; fila <= ultimaFila; fila++) {
    const ciudad = hojaClientes.getRange(fila, 4).getValue();
    if (ciudad) ciudades[ciudad] = (ciudades[ciudad] || 0) + 1;
  }

  // Contar por servicio
  const servicios = {};
  for (let fila = 2; fila <= ultimaFila; fila++) {
    const servicio = hojaClientes.getRange(fila, 5).getValue();
    if (servicio) servicios[servicio] = (servicios[servicio] || 0) + 1;
  }

  // ── Escribir KPIs principales
  hojaDashboard.getRange('A1').setValue('📊 DASHBOARD CRM');
  hojaDashboard.getRange('A1').setFontSize(18).setFontWeight('bold').setFontColor('#1a73e8');

  hojaDashboard.getRange('A3').setValue('RESUMEN GENERAL');
  hojaDashboard.getRange('A3').setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  hojaDashboard.getRange('A3:B3').merge();

  const kpis = [
    ['Total Clientes', total],
    ['🟢 Calientes', calientes],
    ['🟡 Tibios', tibios],
    ['🔴 Fríos', frios],
  ];
  kpis.forEach((kpi, i) => {
    hojaDashboard.getRange(4 + i, 1).setValue(kpi[0]);
    hojaDashboard.getRange(4 + i, 2).setValue(kpi[1]);
    hojaDashboard.getRange(4 + i, 2).setHorizontalAlignment('center').setFontWeight('bold');
  });

  // ── Tabla de clientes por ciudad
  hojaDashboard.getRange('A9').setValue('CLIENTES POR CIUDAD');
  hojaDashboard.getRange('A9').setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  hojaDashboard.getRange('A9:B9').merge();

  hojaDashboard.getRange('A10').setValue('Ciudad');
  hojaDashboard.getRange('B10').setValue('Cantidad');
  hojaDashboard.getRange('A10:B10').setFontWeight('bold').setBackground('#e8f0fe');

  let filaCiudad = 11;
  Object.entries(ciudades).sort((a, b) => b[1] - a[1]).forEach(([ciudad, count]) => {
    hojaDashboard.getRange(filaCiudad, 1).setValue(ciudad);
    hojaDashboard.getRange(filaCiudad, 2).setValue(count);
    filaCiudad++;
  });

  // ── Tabla de clientes por servicio
  hojaDashboard.getRange('D3').setValue('CLIENTES POR SERVICIO');
  hojaDashboard.getRange('D3').setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  hojaDashboard.getRange('D3:E3').merge();

  hojaDashboard.getRange('D4').setValue('Servicio');
  hojaDashboard.getRange('E4').setValue('Cantidad');
  hojaDashboard.getRange('D4:E4').setFontWeight('bold').setBackground('#e8f0fe');

  let filaServicio = 5;
  Object.entries(servicios).sort((a, b) => b[1] - a[1]).forEach(([servicio, count]) => {
    hojaDashboard.getRange(filaServicio, 4).setValue(servicio);
    hojaDashboard.getRange(filaServicio, 5).setValue(count);
    filaServicio++;
  });

  // ── Ajustar columnas
  hojaDashboard.autoResizeColumns(1, 5);

  SpreadsheetApp.getUi().alert('✅ Dashboard actualizado correctamente.');
}


// ══════════════════════════════════════════════
//  CONFIGURAR VALIDACIONES
// ══════════════════════════════════════════════
function configurarValidaciones() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clientes');

  // Estado (columna F)
  const reglaEstado = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Caliente', 'Tibio', 'Frío'], true)
    .setAllowInvalid(false)
    .build();
  hoja.getRange('F2:F1000').setDataValidation(reglaEstado);

  // Ciudad (columna D)
  const reglaCiudad = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Bogotá', 'Medellín', 'Cali', 'Barranquilla', 'Cartagena', 'Otra'], true)
    .setAllowInvalid(false)
    .build();
  hoja.getRange('D2:D1000').setDataValidation(reglaCiudad);

  // Servicio (columna E)
  const reglaServicio = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Desarrollo web', 'Diseño logo', 'Consultoría SEO', 'App móvil', 'Mantenimiento web', 'Otro'], true)
    .setAllowInvalid(false)
    .build();
  hoja.getRange('E2:E1000').setDataValidation(reglaServicio);

  // Vendedor (columna H)
  const reglaVendedor = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Sebastián', 'Ana Torres', 'Carlos Ruiz'], true)
    .setAllowInvalid(false)
    .build();
  hoja.getRange('H2:H1000').setDataValidation(reglaVendedor);

  // Último Contacto (columna G) — formato fecha
  hoja.getRange('G2:G1000').setNumberFormat('dd/MM/yyyy');

  // Encabezados
  hoja.getRange('A1:H1')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  hoja.autoResizeColumns(1, 8);
  SpreadsheetApp.getUi().alert('✅ Validaciones configuradas correctamente.');
}

// ══════════════════════════════════════════════
//  REGISTRAR SEGUIMIENTO
// ══════════════════════════════════════════════
function registrarSeguimiento() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaClientes = ss.getSheetByName('Clientes');
  const hojaSeguimientos = ss.getSheetByName('Seguimientos');

  // Paso 1 — Pedir nombre del cliente
  const respNombre = ui.prompt(
    '📝 Registrar Seguimiento',
    'Escribe el nombre del cliente:',
    ui.ButtonSet.OK_CANCEL
  );
  if (respNombre.getSelectedButton() !== ui.Button.OK) return;
  const nombreCliente = respNombre.getResponseText().trim();
  if (!nombreCliente) { ui.alert('⚠️ Debes escribir un nombre.'); return; }

  // Paso 2 — Pedir tipo de contacto
  const respTipo = ui.prompt(
    '📝 Tipo de Contacto',
    'Tipo de contacto:\n1. Llamada\n2. Email\n3. Reunión\n4. WhatsApp\n\nEscribe el tipo:',
    ui.ButtonSet.OK_CANCEL
  );
  if (respTipo.getSelectedButton() !== ui.Button.OK) return;
  const tipoContacto = respTipo.getResponseText().trim();
  if (!tipoContacto) { ui.alert('⚠️ Debes escribir el tipo de contacto.'); return; }

  // Paso 3 — Pedir notas
  const respNotas = ui.prompt(
    '📝 Notas del Seguimiento',
    'Escribe las notas o resultado del contacto:',
    ui.ButtonSet.OK_CANCEL
  );
  if (respNotas.getSelectedButton() !== ui.Button.OK) return;
  const notas = respNotas.getResponseText().trim();

  // Paso 4 — Obtener vendedor (email del usuario activo)
  // ✅ DESPUÉS — lo toma automáticamente de la hoja Clientes
  let vendedor = 'Sin asignar';
  for (let fila = 2; fila <= hojaClientes.getLastRow(); fila++) {
    const nombre = hojaClientes.getRange(fila, 1).getValue();
    if (nombre.toLowerCase() === nombreCliente.toLowerCase()) {
      vendedor = hojaClientes.getRange(fila, 8).getValue(); // Columna H = Vendedor
      break;
    }
  }

  // Paso 5 — Registrar en hoja Seguimientos
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  hojaSeguimientos.appendRow([fecha, nombreCliente, vendedor, tipoContacto, notas]);

  // Paso 6 — Actualizar "Último Contacto" en hoja Clientes
  const ultimaFila = hojaClientes.getLastRow();
  let clienteEncontrado = false;

  for (let fila = 2; fila <= ultimaFila; fila++) {
    const nombre = hojaClientes.getRange(fila, 1).getValue();
    if (nombre.toLowerCase() === nombreCliente.toLowerCase()) {
      hojaClientes.getRange(fila, 7).setValue(new Date()); // Columna G = Último Contacto
      clienteEncontrado = true;
      break;
    }
  }

  // Paso 7 — Confirmar resultado
  if (clienteEncontrado) {
    ui.alert(`✅ Seguimiento registrado y "Último Contacto" actualizado para ${nombreCliente}.`);
  } else {
    ui.alert(`✅ Seguimiento registrado.\n⚠️ No se encontró "${nombreCliente}" en la lista de clientes — verifica el nombre.`);
  }
}
