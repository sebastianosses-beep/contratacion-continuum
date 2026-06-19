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

    // Envío de email al candidato con link pre-cargado
    if (data._action === 'send_email') {
      const r    = data.recruiter || {};
      const c    = data.contrato  || {};
      const rol  = data.rol       || {};
      const m    = data.metadata  || {};
      const firstName = v(r.nombres || r.nombre).split(' ')[0] || 'equipo';
      const link = v(m.portal_link);

      MailApp.sendEmail({
        to: v(r.email_personal),
        subject: `Tu propuesta de incorporación — ${v(rol.cargo)} en Continuum`,
        htmlBody: `
<div style="font-family:Inter,sans-serif;max-width:560px;margin:0 auto;color:#1A1916;">
  <div style="background:#1A1916;padding:20px 32px;border-radius:8px 8px 0 0;">
    <span style="color:#fff;font-weight:600;font-size:16px;letter-spacing:-0.01em;">Continuum</span>
  </div>
  <div style="background:#F7F6F3;padding:32px;border-radius:0 0 8px 8px;border:1px solid #E4E2DC;border-top:none;">
    <h2 style="font-size:22px;margin:0 0 8px;">Hola, ${firstName}</h2>
    <p style="color:#5C5A54;margin:0 0 24px;">Tienes una propuesta de incorporación esperándote. Revísala y completa tu proceso en unos minutos.</p>
    <div style="background:#fff;border:1px solid #E4E2DC;border-radius:8px;padding:20px;margin-bottom:24px;">
      <table style="width:100%;border-collapse:collapse;font-size:13px;">
        <tr><td style="padding:4px 0;color:#5C5A54;width:140px;">Cargo</td><td style="padding:4px 0;font-weight:500;">${v(rol.cargo)} · ${v(rol.seniority)}</td></tr>
        <tr><td style="padding:4px 0;color:#5C5A54;">Proyecto</td><td style="padding:4px 0;font-weight:500;">${v(rol.proyecto)}</td></tr>
        <tr><td style="padding:4px 0;color:#5C5A54;">Modalidad</td><td style="padding:4px 0;font-weight:500;">${v(c.label)}</td></tr>
        <tr><td style="padding:4px 0;color:#5C5A54;">Fecha de ingreso</td><td style="padding:4px 0;font-weight:500;">${v(c.fecha_ingreso)}</td></tr>
      </table>
    </div>
    <a href="${link}" style="background:#1A1916;color:#fff;text-decoration:none;padding:12px 28px;border-radius:8px;display:inline-block;font-weight:500;font-size:14px;">Ver mi propuesta →</a>
    <p style="color:#9B9890;font-size:12px;margin-top:28px;margin-bottom:0;">¿Tienes dudas? Escríbenos a <a href="mailto:people@continuumhq.com" style="color:#5C5A54;">people@continuumhq.com</a></p>
  </div>
</div>`
      });

      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, email_sent: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Checkpoint parcial — registra estado en tab seguimiento
    if (data._checkpoint) {
      let tab = ss.getSheetByName('seguimiento');
      if (!tab) {
        tab = ss.insertSheet('seguimiento');
        tab.appendRow(['Timestamp','Session ID','Correo corporativo','Email personal','Nombre','Tipo contrato','Estado','Portal Link']);
        tab.getRange(1, 1, 1, 8).setFontWeight('bold');
      }
      const r = data.recruiter || {};
      const m = data.metadata  || {};
      tab.appendRow([
        new Date(),
        v(m.session_id),
        v(m.correo_corporativo),
        v(r.email_personal),
        v(r.nombre),
        v(data.contrato?.tipo),
        v(data._estado),
        v(m.portal_link)
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

    // PandaDoc — solo para contratos Perú e Internacional
    const pandaContracts = ['Chile-Nomina','Peru-Nomina','Peru-Honorarios','Peru-Empresa','Internacional-Servicios']; // Chile-Nomina temporal para test, luego va a Talana
    if (pandaContracts.includes(contract)) {
      try { enviarContratoPandaDoc(data); } catch(pe) { Logger.log('PandaDoc error: ' + pe.message); }
    }

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
// Recordatorio automático — se ejecuta diariamente vía trigger
// ————————————————————————————
function enviarRecordatorios() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const tab = ss.getSheetByName('seguimiento');
  if (!tab || tab.getLastRow() < 2) return;

  const rows = tab.getDataRange().getValues().slice(1); // sin header
  const ESTADOS_INCOMPLETOS = ['iniciado', 'propuesta_aceptada', 'nda_firmado'];
  const DONDE = {
    iniciado:           'revisar y aceptar tu propuesta de incorporación',
    propuesta_aceptada: 'firmar el NDA',
    nda_firmado:        'completar tu ficha de ingreso'
  };

  // Por session_id, quedarse con el estado más reciente
  const sessions = {};
  rows.forEach(row => {
    const sid = String(row[1] || '');
    if (!sid) return;
    const ts = new Date(row[0]);
    if (!sessions[sid] || ts > sessions[sid].ts) {
      sessions[sid] = { ts, email: row[3], nombre: row[4], estado: row[6], link: row[7] };
    }
  });

  const now         = new Date();
  const MIN_ESPERA  =  3 * 60 * 60 * 1000; // 3h mínimo antes de enviar
  const MAX_ESPERA  = 48 * 60 * 60 * 1000; // no molestar después de 48h

  Object.entries(sessions).forEach(([sid, s]) => {
    if (!ESTADOS_INCOMPLETOS.includes(s.estado)) return;
    if (!s.email || !s.link) return;
    const diff = now - s.ts;
    if (diff < MIN_ESPERA || diff > MAX_ESPERA) return;

    const firstName = String(s.nombre || '').split(' ')[0] || 'candidato';
    const donde     = DONDE[s.estado] || 'completar tu proceso';

    MailApp.sendEmail({
      to: s.email,
      subject: 'Recordatorio: completa tu proceso de incorporación en Continuum',
      htmlBody: `
<div style="font-family:Inter,sans-serif;max-width:560px;margin:0 auto;color:#1A1916;">
  <div style="background:#1A1916;padding:20px 32px;border-radius:8px 8px 0 0;">
    <span style="color:#fff;font-weight:600;font-size:16px;letter-spacing:-0.01em;">Continuum</span>
  </div>
  <div style="background:#F7F6F3;padding:32px;border-radius:0 0 8px 8px;border:1px solid #E4E2DC;border-top:none;">
    <h2 style="font-size:22px;margin:0 0 8px;">Hola, ${firstName}</h2>
    <p style="color:#5C5A54;margin:0 0 24px;">Quedaste a un paso de completar tu incorporación en Continuum. Solo te falta <strong>${donde}</strong> — toma menos de 5 minutos.</p>
    <a href="${s.link}" style="background:#1A1916;color:#fff;text-decoration:none;padding:12px 28px;border-radius:8px;display:inline-block;font-weight:500;font-size:14px;">Continuar mi proceso →</a>
    <p style="color:#9B9890;font-size:12px;margin-top:28px;margin-bottom:0;">¿Tienes dudas? Escríbenos a <a href="mailto:people@continuumhq.com" style="color:#5C5A54;">people@continuumhq.com</a></p>
  </div>
</div>`
    });

    // Marcar como recordatorio enviado para no volver a enviar
    tab.appendRow([new Date(), sid, '', s.email, s.nombre, '', 'recordatorio_enviado', s.link]);
    Logger.log('Recordatorio enviado a ' + s.email);
  });
}

// ————————————————————————————
// PandaDoc — genera contrato y envía a firma (Perú / Internacional)
// ————————————————————————————
function enviarContratoPandaDoc(data) {
  const PANDADOC_API_KEY  = PropertiesService.getScriptProperties().getProperty('PANDADOC_API_KEY');
  const TEMPLATE_ID       = 'gWLdyf3ERaWCetL9wAtg2b';
  const r   = data.recruiter || {};
  const c   = data.contrato  || {};
  const rol = data.rol       || {};
  const f   = data.ficha     || {};

  const firstName = v(f.nombres || r.nombres || r.nombre).split(' ')[0];
  const lastName  = [v(f.apellidos?.primero || r.apellido1), v(f.apellidos?.segundo || r.apellido2)].filter(Boolean).join(' ');
  const email     = v(r.email_personal);

  const body = {
    name:          `Contrato — ${firstName} ${lastName} — ${v(rol.cargo)}`,
    template_uuid: TEMPLATE_ID,
    recipients: [{
      email:      email,
      first_name: firstName,
      last_name:  lastName,
      role:       'Candidate'
    }],
    tokens: [
      { name: 'Recruiter.Company', value: 'Continuum' },
      { name: 'Job Title',         value: v(rol.cargo) + ' · ' + v(rol.seniority) },
      { name: 'Salary',            value: v(c.sueldo_liquido) },
      { name: 'CurrencyType',      value: v(c.moneda) },
      { name: 'Start Date',        value: v(c.fecha_ingreso) }
    ],
    metadata: { session_id: v(data.metadata?.session_id) }
  };

  const response = UrlFetchApp.fetch('https://api.pandadoc.com/public/v1/documents', {
    method:  'post',
    headers: {
      'Authorization': 'API-Key ' + PANDADOC_API_KEY,
      'Content-Type':  'application/json'
    },
    payload:              JSON.stringify(body),
    muteHttpExceptions:   true
  });

  const result = JSON.parse(response.getContentText());
  Logger.log('PandaDoc documento creado: ' + JSON.stringify(result));

  if (!result.id) throw new Error('PandaDoc no retornó id: ' + response.getContentText());

  // Polling hasta que el documento esté en draft (máx 30s)
  let ready = false;
  for (let i = 0; i < 10; i++) {
    Utilities.sleep(3000);
    const statusResp = UrlFetchApp.fetch(`https://api.pandadoc.com/public/v1/documents/${result.id}`, {
      headers: { 'Authorization': 'API-Key ' + PANDADOC_API_KEY },
      muteHttpExceptions: true
    });
    const statusData = JSON.parse(statusResp.getContentText());
    Logger.log('PandaDoc status poll: ' + statusData.status);
    if (statusData.status === 'document.draft') { ready = true; break; }
  }

  if (!ready) throw new Error('PandaDoc: documento no llegó a draft después de 30s');

  const sendResp = UrlFetchApp.fetch(`https://api.pandadoc.com/public/v1/documents/${result.id}/send`, {
    method:  'post',
    headers: {
      'Authorization': 'API-Key ' + PANDADOC_API_KEY,
      'Content-Type':  'application/json'
    },
    payload:            JSON.stringify({ message: `Hola ${firstName}, aquí está tu contrato para revisar y firmar.`, silent: false }),
    muteHttpExceptions: true
  });

  Logger.log('PandaDoc enviado a firma: ' + sendResp.getContentText());
}

// ————————————————————————————
// Test directo PandaDoc desde el editor
// ————————————————————————————
function testPandaDoc() {
  const data = {
    recruiter: { nombres: 'Francisco', apellido1: 'Morales', apellido2: 'Monsalves', email_personal: 'sebastian.osses@continuumhq.com' },
    contrato:  { tipo: 'Peru-Nomina', sueldo_liquido: '3000', moneda: 'PEN', fecha_ingreso: '2026-07-01' },
    rol:       { cargo: 'Software Engineer', seniority: 'Senior', proyecto: 'Proyecto Test' },
    ficha:     { nombres: 'Francisco', apellidos: { primero: 'Morales', segundo: 'Monsalves' } },
    metadata:  { session_id: 'test-123' }
  };
  enviarContratoPandaDoc(data);
  Logger.log('testPandaDoc completado');
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
