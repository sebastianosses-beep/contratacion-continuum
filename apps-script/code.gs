// ============================================================
// Continuum — Portal de Incorporación → Google Sheets
// Apps Script Web App: recibe POST del formulario HTML
// y escribe en la pestaña correcta según tipo de contrato.
//
// Pestañas requeridas:
//   "Nómina Chile" | "Nómina Perú" | "Prestación de Servicios" | "Empresa Jurídica"
// ============================================================

const SHEETS = {
  NOMINA_CL:   'Nómina Chile',
  NOMINA_PE:   'Nómina Perú',
  PRESTACION:  'Prestación de Servicios',
  EMPRESA:     'Empresa Jurídica'
};

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const contract = data.contrato?.tipo;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let sheetName, row;
    switch (contract) {
      case 'Chile-Nomina':
        sheetName = SHEETS.NOMINA_CL;
        row = rowNominaChile(data);
        break;
      case 'Peru-Nomina':
        sheetName = SHEETS.NOMINA_PE;
        row = rowNominaPeru(data);
        break;
      case 'Chile-Honorarios':
      case 'Peru-Honorarios':
      case 'Internacional-Servicios':
        sheetName = SHEETS.PRESTACION;
        row = rowPrestacion(data);
        break;
      case 'Chile-Empresa':
      case 'Peru-Empresa':
        sheetName = SHEETS.EMPRESA;
        row = rowEmpresa(data);
        break;
      default:
        throw new Error('Tipo de contrato desconocido: ' + contract);
    }

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error('Pestaña no encontrada: ' + sheetName);
    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, sheet: sheetName }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ————————————————————————————
// Helpers
// ————————————————————————————
const ts = () => new Date().toLocaleString('es-CL', { timeZone: 'America/Santiago' });
const v  = (x) => (x !== null && x !== undefined) ? String(x) : '';

// ————————————————————————————
// Nómina Chile
// Columnas (orden debe coincidir exactamente con el header de la pestaña)
// ————————————————————————————
function rowNominaChile(d) {
  const r   = d.recruiter  || {};
  const c   = d.contrato   || {};
  const rol = d.rol        || {};
  const f   = d.ficha      || {};
  const em  = f.emergencia || {};
  const bn  = f.banco      || {};
  const pv  = f.prevision  || {};
  const ub  = f.ubicacion  || {};
  const m   = d.metadata   || {};

  return [
    ts(),                          // Timestamp
    v(m.correo_corporativo),       // Correo Corporativo
    'Pendiente',                   // Estado
    v(r.nombre),                   // Nombre(s)
    v(f.apellidos?.primero),       // Primer Apellido
    v(f.apellidos?.segundo),       // Segundo Apellido
    v(r.email_personal),           // Dirección de correo electrónico
    v(r.tipo_documento),           // Tipo de documento
    v(r.numero_documento),         // Número de Identificación
    v(r.sexo),                     // Género
    v(r.fecha_nacimiento),         // Fecha de Nacimiento
    v(r.estado_civil),             // Estado Civil
    v(r.nacionalidad),             // Nacionalidad
    v(f.profesion),                // Profesión
    v(ub.ciudad),                  // Ciudad
    v(ub.comuna),                  // Comuna
    v(ub.direccion),               // Dirección
    v(pv.afp),                     // AFP
    v(pv.salud),                   // Sistema Salud
    v(bn.banco),                   // Banco
    v(bn.tipo),                    // Tipo de cuenta
    v(bn.numero),                  // Número de cuenta
    v(f.cargas),                   // Cargas familiares
    v(em.nombre),                  // Emergencia - Nombre
    v(em.parentesco),              // Emergencia - Parentesco
    v(em.direccion),               // Emergencia - Dirección
    v(em.telefono),                // Emergencia - Teléfono
    v(f.medico?.sangre),           // Tipo sanguíneo
    v(f.talla),                    // Talla
    v(rol.cargo),                  // Cargo
    v(rol.seniority),              // Seniority
    v(rol.proyecto),               // Proyecto / Líderes
    v(rol.buddy),                  // Buddy
    v(c.sueldo_liquido),           // Sueldo líquido
    v(c.moneda),                   // Moneda
    v(c.duracion),                 // Duración
    v(c.fecha_ingreso),            // Fecha de ingreso Continuum
    v(c.costo_empresa_usd),        // Costo empresa USD
  ];
}

// ————————————————————————————
// Nómina Perú
// ————————————————————————————
function rowNominaPeru(d) {
  const r   = d.recruiter  || {};
  const c   = d.contrato   || {};
  const rol = d.rol        || {};
  const f   = d.ficha      || {};
  const em  = f.emergencia || {};
  const bn  = f.banco      || {};
  const pv  = f.prevision  || {};
  const ub  = f.ubicacion  || {};
  const dom = ub.domicilio || {};
  const fam = f.familiares || [];
  const edu = f.educacion  || {};
  const m   = d.metadata   || {};

  return [
    ts(),
    v(m.correo_corporativo),
    'Pendiente',
    v(r.nombre),
    v(f.apellidos?.primero),
    v(f.apellidos?.segundo),
    v(r.email_personal),
    v(r.tipo_documento),
    v(r.numero_documento),         // DNI
    v(r.sexo),
    v(r.fecha_nacimiento),
    v(r.estado_civil),
    v(r.nacionalidad),
    v(f.profesion),
    v(ub.lugar_nacimiento),        // Lugar de nacimiento
    v(ub.departamento),
    v(ub.provincia),
    v(ub.distrito),
    v(dom.calle),                  // Domicilio actual
    v(dom.urbanizacion),
    v(dom.distrito),
    v(dom.provincia),
    v(dom.departamento),
    v(bn.banco),                   // Banco remuneración
    v(bn.numero),
    v(bn.cci),                     // CCI
    v(bn.cts_banco),               // Banco CTS
    v(bn.cts_numero),
    v(bn.cts_cci),
    v(pv.afp),                     // AFP / ONP
    v(pv.fecha_afiliacion),
    v(pv.cuspp),
    v(pv.eps),
    // Familiar 1
    v(fam[0]?.parentesco), v(fam[0]?.nombre), v(fam[0]?.fecha_nacimiento), v(fam[0]?.dni),
    // Familiar 2
    v(fam[1]?.parentesco), v(fam[1]?.nombre), v(fam[1]?.fecha_nacimiento), v(fam[1]?.dni),
    // Familiar 3
    v(fam[2]?.parentesco), v(fam[2]?.nombre), v(fam[2]?.fecha_nacimiento), v(fam[2]?.dni),
    v(edu.institucion),
    v(edu.fecha_egreso),
    v(f.impuestos?.empleo_anterior_planilla ? 'Sí' : 'No'),
    v(em.nombre),
    v(em.parentesco),
    v(em.direccion),
    v(em.telefono),
    v(f.talla),
    v(rol.cargo),
    v(rol.seniority),
    v(rol.proyecto),
    v(rol.buddy),
    v(c.sueldo_liquido),
    v(c.moneda),
    v(c.duracion),
    v(c.fecha_ingreso),
    v(c.costo_empresa_usd),
  ];
}

// ————————————————————————————
// Prestación de Servicios (Chile / Perú / Internacional)
// ————————————————————————————
function rowPrestacion(d) {
  const r   = d.recruiter  || {};
  const c   = d.contrato   || {};
  const rol = d.rol        || {};
  const f   = d.ficha      || {};
  const em  = f.emergencia || {};
  const ub  = f.ubicacion  || {};
  const m   = d.metadata   || {};
  const isIntl = c.tipo === 'Internacional-Servicios';
  const bn  = isIntl ? (f.banco_internacional || {}) : (f.banco || {});

  return [
    ts(),
    v(m.correo_corporativo),
    'Pendiente',
    v(c.tipo),                     // Tipo de contrato (distingue Chile/Perú/Intl)
    v(r.nombre),
    v(f.apellidos?.primero),
    v(f.apellidos?.segundo),
    v(r.email_personal),
    v(r.tipo_documento),
    v(r.numero_documento),
    v(r.sexo),
    v(r.fecha_nacimiento),
    v(r.estado_civil),
    v(r.nacionalidad),
    v(f.profesion),
    // Ubicación — aplana las 3 variantes en columnas comunes
    isIntl ? v(ub.direccion)    : v(ub.ciudad),      // Dirección / Ciudad
    isIntl ? v(ub.ciudad)       : v(ub.comuna),      // Ciudad / Comuna
    isIntl ? v(ub.codigo_postal): '',                // Código postal
    isIntl ? v(ub.pais)         : '',                // País
    // Banco
    v(bn.banco),
    isIntl ? v(bn.swift)  : '',                      // SWIFT / ABA / IBAN
    isIntl ? v(bn.cuenta) : v(bn.numero),            // Número de cuenta
    isIntl ? v(bn.moneda) : '',                      // Moneda de pago
    v(em.nombre),
    v(em.parentesco),
    v(em.direccion),
    v(em.telefono),
    v(f.talla),
    v(rol.cargo),
    v(rol.seniority),
    v(rol.proyecto),
    v(rol.buddy),
    v(c.sueldo_liquido),
    v(c.moneda),
    v(c.duracion),
    v(c.fecha_ingreso),
    v(c.costo_empresa_usd),
  ];
}

// ————————————————————————————
// Empresa Jurídica (Chile / Perú)
// ————————————————————————————
function rowEmpresa(d) {
  const r   = d.recruiter || {};
  const c   = d.contrato  || {};
  const rol = d.rol       || {};
  const m   = d.metadata  || {};
  // Datos de empresa vienen del paso reclutador (d.empresa) O del paso ficha (d.ficha.empresa)
  const emp = (d.empresa && d.empresa.razon_social) ? d.empresa : (d.ficha?.empresa || {});

  return [
    ts(),
    v(m.correo_corporativo),
    'Pendiente',
    v(c.tipo),                                        // Chile-Empresa / Peru-Empresa
    v(emp.razon_social),
    v(emp.rut),
    v(emp.representante_legal || emp.representante),
    v(emp.direccion),
    v(emp.descripcion_proyecto),
    v(emp.personeria),
    v(r.nombre),                                      // Contacto / Representante Continuum
    v(r.email_personal),
    v(r.tipo_documento),
    v(r.numero_documento),
    v(rol.cargo),
    v(rol.seniority),
    v(rol.proyecto),
    v(rol.buddy),
    v(c.sueldo_liquido),
    v(c.moneda),
    v(c.duracion),
    v(c.fecha_ingreso),
    v(c.costo_empresa_usd),
  ];
}

// ————————————————————————————
// Test manual desde el editor de Apps Script
// Ejecuta esta función para verificar que el script
// puede encontrar las pestañas y escribir una fila de prueba.
// ————————————————————————————
function testWrite() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = Object.values(SHEETS);
  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      Logger.log('⚠️  Pestaña no encontrada: ' + name);
    } else {
      Logger.log('✅  Pestaña OK: ' + name + ' (' + sheet.getLastRow() + ' filas)');
    }
  });
}
