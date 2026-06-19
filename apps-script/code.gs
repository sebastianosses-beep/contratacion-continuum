// ============================================================
// Continuum — Portal de Incorporación → Google Sheets
// Apps Script Web App: recibe POST del formulario HTML
// y escribe en la pestaña correcta según tipo de contrato.
//
// Pestañas requeridas:
//   nomina_chile | nomina_peru | prestacion_servicios | empresa_juridica
// ============================================================

const SHEETS = {
  NOMINA_CL:   'nomina_chile',
  NOMINA_PE:   'nomina_peru',
  PRESTACION:  'prestacion_servicios',
  EMPRESA:     'empresa_juridica'
};

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Checkpoint parcial — registra estado en tab seguimiento
    if (data._checkpoint) {
      let tab = ss.getSheetByName('seguimiento');
      if (!tab) {
        tab = ss.insertSheet('seguimiento');
        tab.appendRow(['Timestamp','Session ID','Correo corporativo','Email personal','Nombre','Tipo contrato','Estado']);
        tab.getRange(1, 1, 1, 7).setFontWeight('bold');
      }
      const r = data.recruiter || {};
      const m = data.metadata  || {};
      tab.appendRow([
        new Date().toLocaleString('es-CL', { timeZone: 'America/Santiago' }),
        v(m.session_id),
        v(m.correo_corporativo),
        v(r.email_personal),
        v(r.nombre),
        v(data.contrato?.tipo),
        v(data._estado)
      ]);
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, checkpoint: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const contract = data.contrato?.tipo;
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
const ts  = () => new Date().toLocaleString('es-CL', { timeZone: 'America/Santiago' });
const v   = (x) => (x !== null && x !== undefined) ? String(x) : '';

function firmasCols(d) {
  const f = d.firmas || {};
  const p = f.propuesta || {};
  const n = f.nda || {};
  return [
    v(p.nombre_confirmacion),   // Propuesta - nombre confirmación
    p.acepta ? 'Sí' : 'No',    // Propuesta - aceptó
    v(p.timestamp),             // Propuesta - timestamp
    v(n.firma),                 // NDA - firma (nombre escrito)
    n.acepta ? 'Sí' : 'No',    // NDA - aceptó
    v(n.timestamp),             // NDA - timestamp
  ];
}

// ————————————————————————————
// Nómina Chile
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
    ts(),
    v(m.correo_corporativo),
    'Pendiente',
    v(f.nombres),
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
    v(ub.ciudad),
    v(ub.comuna),
    v(ub.direccion),
    v(pv.afp),
    v(pv.salud),
    v(bn.banco),
    v(bn.tipo),
    v(bn.numero),
    v(f.cargas),
    v(em.nombre),
    v(em.parentesco),
    v(em.direccion),
    v(em.telefono),
    v(f.medico?.sangre),
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
    ...firmasCols(d),
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
    v(f.nombres),
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
    v(ub.lugar_nacimiento),
    v(ub.departamento),
    v(ub.provincia),
    v(ub.distrito),
    v(dom.calle),
    v(dom.urbanizacion),
    v(dom.distrito),
    v(dom.provincia),
    v(dom.departamento),
    v(bn.banco),
    v(bn.numero),
    v(bn.cci),
    v(bn.cts_banco),
    v(bn.cts_numero),
    v(bn.cts_cci),
    v(pv.afp),
    v(pv.fecha_afiliacion),
    v(pv.cuspp),
    v(pv.eps),
    v(fam[0]?.parentesco), v(fam[0]?.nombre), v(fam[0]?.fecha_nacimiento), v(fam[0]?.dni),
    v(fam[1]?.parentesco), v(fam[1]?.nombre), v(fam[1]?.fecha_nacimiento), v(fam[1]?.dni),
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
    ...firmasCols(d),
  ];
}

// ————————————————————————————
// Prestación de Servicios (Chile / Perú / Internacional)
// ————————————————————————————
function rowPrestacion(d) {
  const r      = d.recruiter  || {};
  const c      = d.contrato   || {};
  const rol    = d.rol        || {};
  const f      = d.ficha      || {};
  const em     = f.emergencia || {};
  const ub     = f.ubicacion  || {};
  const m      = d.metadata   || {};
  const isIntl = c.tipo === 'Internacional-Servicios';
  const bn     = isIntl ? (f.banco_internacional || {}) : (f.banco || {});

  return [
    ts(),
    v(m.correo_corporativo),
    'Pendiente',
    v(c.tipo),
    v(f.nombres),
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
    isIntl ? v(ub.direccion)     : v(ub.ciudad),
    isIntl ? v(ub.ciudad)        : v(ub.comuna),
    isIntl ? v(ub.codigo_postal) : '',
    isIntl ? v(ub.pais)          : '',
    v(bn.banco),
    isIntl ? v(bn.swift)  : '',
    isIntl ? v(bn.cuenta) : v(bn.numero),
    isIntl ? ''           : v(bn.cci),
    isIntl ? v(bn.moneda) : '',
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
    ...firmasCols(d),
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
  const emp = (d.empresa && d.empresa.razon_social) ? d.empresa : (d.ficha?.empresa || {});

  return [
    ts(),
    v(m.correo_corporativo),
    'Pendiente',
    v(c.tipo),
    v(emp.razon_social),
    v(emp.rut),
    v(emp.representante_legal || emp.representante),
    v(emp.direccion),
    v(emp.descripcion_proyecto),
    v(emp.personeria),
    v(r.nombre),
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
    ...firmasCols(d),
  ];
}

// ————————————————————————————
// Test manual desde el editor de Apps Script
// ————————————————————————————
function testWrite() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.values(SHEETS).forEach(name => {
    const sheet = ss.getSheetByName(name);
    Logger.log(sheet
      ? '✅  ' + name + ' (' + sheet.getLastRow() + ' filas)'
      : '⚠️  No encontrada: ' + name
    );
  });
}
