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

function dbgLog(ss, msg) {
  try {
    const t = ss.getSheetByName('_debug') || ss.insertSheet('_debug');
    t.appendRow([new Date(), msg]);
  } catch(_) {}
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    dbgLog(ss, 'doPost | tipo=' + data.contrato?.tipo + ' | _action=' + data._action + ' | _checkpoint=' + data._checkpoint);

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

    // PandaDoc — genera y firma el documento del contrato (todos los tipos, incluido Chile-Nómina)
    const pandaContracts = ['Chile-Nomina','Peru-Nomina','Peru-Honorarios','Peru-Empresa','Internacional-Servicios'];
    if (pandaContracts.includes(contract)) {
      dbgLog(ss, 'PandaDoc: disparando para ' + contract);
      try { enviarContratoPandaDoc(data); } catch(pe) { dbgLog(ss, 'PandaDoc ERROR: ' + pe.message); }
    } else {
      dbgLog(ss, 'PandaDoc: contrato no en lista (' + contract + ')');
    }

    // Talana — crea la ficha del trabajador (Persona + Contrato) en paralelo, solo Chile-Nómina.
    // No genera ni firma documentos — eso lo hace PandaDoc arriba. Si Talana falla, no bloquea
    // el flujo del candidato (queda registrado en _debug para revisión manual).
    if (contract === 'Chile-Nomina') {
      dbgLog(ss, 'Talana: disparando creación de ficha para ' + contract);
      try { enviarContratoTalana(data); } catch(te) { dbgLog(ss, 'Talana ERROR: ' + te.message); }
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

// Reconstruye el RUT en formato "NNNNNNNN-N" sin importar puntos, guion o espacios en el input
function normalizarRut(rut) {
  const limpio = v(rut).replace(/[^0-9kK]/g, '');
  if (limpio.length < 2) return limpio;
  const cuerpo = limpio.slice(0, -1);
  const verificador = limpio.slice(-1).toUpperCase();
  return cuerpo + '-' + verificador;
}

// Convierte un número entero a su representación en palabras (español)
function numeroAPalabras(num) {
  const n = parseInt(String(num).replace(/[^\d]/g, ''), 10);
  if (!n || isNaN(n)) return '';

  const UNIDADES = ['', 'uno', 'dos', 'tres', 'cuatro', 'cinco', 'seis', 'siete', 'ocho', 'nueve'];
  const DECENAS  = ['', 'diez', 'veinte', 'treinta', 'cuarenta', 'cincuenta', 'sesenta', 'setenta', 'ochenta', 'noventa'];
  const ESPECIALES = {
    11: 'once', 12: 'doce', 13: 'trece', 14: 'catorce', 15: 'quince',
    16: 'dieciséis', 17: 'diecisiete', 18: 'dieciocho', 19: 'diecinueve',
    21: 'veintiuno', 22: 'veintidós', 23: 'veintitrés', 24: 'veinticuatro',
    25: 'veinticinco', 26: 'veintiséis', 27: 'veintisiete', 28: 'veintiocho', 29: 'veintinueve'
  };
  const CENTENAS = ['', 'ciento', 'doscientos', 'trescientos', 'cuatrocientos', 'quinientos', 'seiscientos', 'setecientos', 'ochocientos', 'novecientos'];

  function tresDigitos(num3) {
    if (num3 === 0) return '';
    if (num3 === 100) return 'cien';
    if (ESPECIALES[num3]) return ESPECIALES[num3];
    const c = Math.floor(num3 / 100);
    const resto = num3 % 100;
    let str = CENTENAS[c];
    if (resto) {
      if (ESPECIALES[resto]) {
        str += (str ? ' ' : '') + ESPECIALES[resto];
      } else {
        const d = Math.floor(resto / 10);
        const u = resto % 10;
        const du = d ? (DECENAS[d] + (u ? ' y ' + UNIDADES[u] : '')) : UNIDADES[u];
        str += (str ? ' ' : '') + du;
      }
    }
    return str;
  }

  function grupo(num3, singular, plural) {
    if (num3 === 0) return '';
    if (num3 === 1) return singular;
    return tresDigitos(num3) + ' ' + plural;
  }

  if (n === 0) return 'cero pesos';

  const millones = Math.floor(n / 1000000);
  const miles    = Math.floor((n % 1000000) / 1000);
  const resto    = n % 1000;

  let partes = [];
  if (millones) partes.push(grupo(millones, 'un millón', 'millones'));
  if (miles)    partes.push((miles === 1 ? 'mil' : tresDigitos(miles) + ' mil'));
  if (resto)    partes.push(tresDigitos(resto));

  return partes.join(' ').trim() + ' pesos';
}

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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const PANDADOC_API_KEY  = PropertiesService.getScriptProperties().getProperty('PANDADOC_API_KEY');
  const TEMPLATE_ID       = 'nyeZBmKQXnWYF4xPr55b6F'; // Plantilla Contrato Nómina Chile v1 (Trabajador.*/Contrato.*, roles Signer/Recruiter)
  const r   = data.recruiter || {};
  const c   = data.contrato  || {};
  const rol = data.rol       || {};
  const f   = data.ficha     || {};

  const firstName = v(f.nombres || r.nombres || r.nombre).split(' ')[0];
  const lastName  = [v(f.apellidos?.primero || r.apellido1), v(f.apellidos?.segundo || r.apellido2)].filter(Boolean).join(' ');
  const email     = v(r.email_personal);

  const ub = f.ubicacion || {};
  const domicilio = [v(ub.direccion), v(ub.comuna), v(ub.ciudad)].filter(Boolean).join(', ');
  const fechaContrato = Utilities.formatDate(new Date(), 'America/Santiago', "d 'de' MMMM 'de' yyyy");

  const body = {
    name:          `Contrato — ${firstName} ${lastName} — ${v(rol.cargo)}`,
    template_uuid: TEMPLATE_ID,
    recipients: [
      {
        email:      email,
        first_name: firstName,
        last_name:  lastName,
        role:       'Trabajador'
      }
    ],
    tokens: [
      { name: 'Contrato.Fecha',              value: fechaContrato },
      { name: 'Trabajador.NombreCompleto',   value: `${firstName} ${lastName}` },
      { name: 'Trabajador.RUT',              value: v(r.numero_documento) },
      { name: 'Trabajador.Nacionalidad',     value: v(r.nacionalidad) },
      { name: 'Trabajador.FechaNacimiento',  value: v(r.fecha_nacimiento) },
      { name: 'Trabajador.EstadoCivil',      value: v(r.estado_civil) },
      { name: 'Trabajador.Email',            value: email },
      { name: 'Trabajador.Domicilio',        value: domicilio },
      { name: 'Contrato.Cargo',              value: v(rol.cargo) + ' ' + v(rol.seniority) },
      { name: 'Contrato.FuncionesEspecificas', value: v(rol.descripcion) },
      { name: 'Contrato.SueldoBase',         value: v(c.sueldo_liquido) + ' ' + v(c.moneda) },
      { name: 'Contrato.SueldoEnPalabras',   value: numeroAPalabras(c.sueldo_liquido) },
      { name: 'Contrato.FechaIngreso',       value: v(c.fecha_ingreso) }
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
  dbgLog(ss, 'PandaDoc POST /documents → ' + response.getResponseCode() + ' | id=' + result.id);

  if (!result.id) throw new Error('PandaDoc no retornó id: ' + response.getContentText());

  // Polling hasta que el documento esté en draft (máx 30s)
  let ready = false;
  let ultimoStatus = '';
  for (let i = 0; i < 10; i++) {
    Utilities.sleep(3000);
    const statusResp = UrlFetchApp.fetch(`https://api.pandadoc.com/public/v1/documents/${result.id}`, {
      headers: { 'Authorization': 'API-Key ' + PANDADOC_API_KEY },
      muteHttpExceptions: true
    });
    const statusData = JSON.parse(statusResp.getContentText());
    ultimoStatus = statusData.status;
    if (statusData.status === 'document.draft') { ready = true; break; }
  }
  dbgLog(ss, 'PandaDoc polling terminó — último status: ' + ultimoStatus + ' (ready=' + ready + ')');

  if (!ready) throw new Error('PandaDoc: documento no llegó a draft después de 30s (status final: ' + ultimoStatus + ')');

  const sendResp = UrlFetchApp.fetch(`https://api.pandadoc.com/public/v1/documents/${result.id}/send`, {
    method:  'post',
    headers: {
      'Authorization': 'API-Key ' + PANDADOC_API_KEY,
      'Content-Type':  'application/json'
    },
    payload:            JSON.stringify({ message: `Hola ${firstName}, aquí está tu contrato para revisar y firmar.`, silent: false }),
    muteHttpExceptions: true
  });

  const sendStatus = sendResp.getResponseCode();
  dbgLog(ss, 'PandaDoc POST /send → ' + sendStatus + ' | ' + sendResp.getContentText());
  if (sendStatus < 200 || sendStatus >= 300) {
    throw new Error('PandaDoc falló al enviar a firma: ' + sendResp.getContentText());
  }
}

// ————————————————————————————
// Test directo PandaDoc desde el editor
// ————————————————————————————
function testPandaDoc() {
  const data = {
    recruiter: {
      nombres: 'Mario', apellido1: 'Aliaga', apellido2: 'Guzmán',
      email_personal: 'sebastian.osses@continuumhq.com',
      numero_documento: '13.735.003-3', nacionalidad: 'Chilena',
      fecha_nacimiento: '7 de Noviembre de 1979', estado_civil: 'Soltero'
    },
    contrato:  { tipo: 'Chile-Nomina', sueldo_liquido: '3.612.686', moneda: 'CLP', fecha_ingreso: '22-06-2026' },
    rol:       { cargo: 'Software Engineer', seniority: 'Senior', proyecto: 'Proyecto Test', descripcion: 'Diseñar y desarrollar microservicios en Node.js / Java / Go.' },
    ficha:     {
      nombres: 'Mario', apellidos: { primero: 'Aliaga', segundo: 'Guzmán' },
      ubicacion: { ciudad: 'Cartagena', comuna: 'Valparaíso', direccion: 'Echaurren N°7' }
    },
    metadata:  { session_id: 'test-123' }
  };
  enviarContratoPandaDoc(data);
  Logger.log('testPandaDoc completado');
}

// ————————————————————————————
// Talana — catálogos (solo lectura, sin riesgo)
// ————————————————————————————
// IMPORTANTE: a pesar del nombre, "sandbox.talana.dev" NO acepta el token de producción.
// El token entregado por Talana SAC es solo para el dominio productivo real: talana.com
const TALANA_BASE_URL = 'https://talana.com/es/api';

function talanaGet(path) {
  const TALANA_API_KEY = PropertiesService.getScriptProperties().getProperty('TALANA_API_KEY');
  const resp = UrlFetchApp.fetch(TALANA_BASE_URL + path, {
    method: 'get',
    headers: { 'Authorization': 'Token ' + TALANA_API_KEY },
    muteHttpExceptions: true
  });
  return JSON.parse(resp.getContentText());
}

function talanaPost(path, body) {
  const TALANA_API_KEY = PropertiesService.getScriptProperties().getProperty('TALANA_API_KEY');
  const resp = UrlFetchApp.fetch(TALANA_BASE_URL + path, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Token ' + TALANA_API_KEY },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  return { status: resp.getResponseCode(), body: JSON.parse(resp.getContentText() || '{}') };
}

// Busca el id de un registro de catálogo por nombre (AFP, Banco, Isapre soportan ?nombre=)
function talanaBuscarIdPorNombre(path, nombre) {
  if (!nombre) return null;
  const lista = talanaGet(path + '?nombre=' + encodeURIComponent(nombre));
  if (!Array.isArray(lista) || lista.length === 0) return null;
  return lista[0].id;
}

// Imprime el catálogo de Razón Social — necesario para 'empleadorRazonSocial' en /contrato/
function testTalanaRazonSocial() {
  Logger.log(JSON.stringify(talanaGet('/razonSocial/'), null, 2));
}

// Imprime los 5 catálogos en el Logger — no crea ni modifica nada en Talana
function testTalanaCatalogos() {
  const catalogos = {
    'AFP':            '/afp/',
    'Isapre/Prevision': '/prevision/',
    'Banco':          '/banco/',
    'TipoContrato':   '/tipoContrato/',
    'JornadaLaboral': '/jornadaLaboral/'
  };

  Object.entries(catalogos).forEach(([nombre, path]) => {
    try {
      const data = talanaGet(path);
      Logger.log('— ' + nombre + ' —');
      Logger.log(JSON.stringify(data, null, 2));
    } catch (err) {
      Logger.log('ERROR en ' + nombre + ': ' + err.message);
    }
  });
}

// ————————————————————————————
// Talana — crea Persona y Contrato (Chile-Nómina)
//
// Endpoints reales confirmados en developers.talana.com (NO inventados):
//   POST /persona/   — crea el trabajador, identificado por 'rut'
//   GET  /persona/?rut=...  — recupera el 'id' numérico interno
//   POST /contrato/  — usa ese 'id' en el campo 'empleado'
//
// PENDIENTE (no implementado todavía): firma digital y generación de PDF.
// Talana no genera el PDF del contrato desde una plantilla vía API — solo
// guarda los datos estructurados. Para usar /document/requestSignature
// primero hay que subir un PDF ya armado vía /documentos/, lo cual requiere
// que resolvamos cómo generamos ese PDF (Google Docs + merge, o PandaDoc).
// ————————————————————————————
function enviarContratoTalana(data) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const r   = data.recruiter || {};
  const c   = data.contrato  || {};
  const rol = data.rol       || {};
  const f   = data.ficha     || {};
  const m   = data.metadata  || {};
  const pv  = f.prevision    || {};
  const bn  = f.banco        || {};

  const rut = normalizarRut(r.numero_documento); // Talana espera "NNNNNNNN-N", sin importar cómo venga el input
  const sexoMap = { 'Masculino': 'M', 'Femenino': 'F' };

  // 1) Crear Persona
  const personaBody = {
    rut: rut,
    nombre: v(f.nombres || r.nombres || r.nombre),
    apellidoPaterno: v(f.apellidos?.primero || r.apellido1),
    apellidoMaterno: v(f.apellidos?.segundo || r.apellido2),
    email: v(m.correo_corporativo),
    fechaNacimiento: v(r.fecha_nacimiento),
    sexo: sexoMap[r.sexo] || 'N',
    detalles: [{ emailPersonal: v(r.email_personal) }]
  };

  const respPersona = talanaPost('/persona/', personaBody);
  dbgLog(ss, 'Talana POST /persona/ → ' + respPersona.status + ' | ' + JSON.stringify(respPersona.body));
  const yaExiste = respPersona.status !== 201 && /ya existe/i.test(respPersona.body?.detail || '');
  if (respPersona.status !== 201 && !yaExiste) {
    throw new Error('Talana falló al crear persona: ' + JSON.stringify(respPersona.body));
  }
  if (yaExiste) {
    dbgLog(ss, 'Talana: persona ya existía, continuando con la existente — ' + respPersona.body.detail);
  }

  // 2) Recuperar el id numérico interno de esa persona
  const listaPersona = talanaGet('/persona/?rut=' + encodeURIComponent(rut));
  if (!Array.isArray(listaPersona) || listaPersona.length === 0) {
    throw new Error('Talana: no se encontró la persona recién creada con rut ' + rut);
  }
  const idPersona = listaPersona[0].id;
  dbgLog(ss, 'Talana: id numérico de persona = ' + idPersona);

  // 3) Crear Contrato — catálogos resueltos en vivo por nombre
  const afpId    = talanaBuscarIdPorNombre('/afp/', pv.afp);
  const isapreId = talanaBuscarIdPorNombre('/prevision/', pv.salud);
  const bancoId  = talanaBuscarIdPorNombre('/banco/', bn.banco);

  // 'empleadorRazonSocial' es obligatorio: la razón social bajo la cual se contrata.
  // Hay 2 en la cuenta (CONTINUUM SPA y CONTINUUM HOLDING SPA) — filtramos por el RUT
  // real de Continuum SpA (76091977-2), NUNCA tomar la primera del array a ciegas.
  const CONTINUUM_SPA_RUT = '76091977-2';
  const razonesSociales = talanaGet('/razonSocial/');
  const razonSocial = Array.isArray(razonesSociales)
    ? razonesSociales.find(rs => rs.rut === CONTINUUM_SPA_RUT)
    : null;
  if (!razonSocial) throw new Error('Talana: no se encontró la razón social CONTINUUM SPA (rut ' + CONTINUUM_SPA_RUT + ')');
  const razonSocialId = razonSocial.id;

  const contratoBody = {
    empleado: idPersona,
    empleadorRazonSocial: razonSocialId,
    cargo: v(rol.cargo) + ' ' + v(rol.seniority),
    desde: v(c.fecha_ingreso),
    fechaContratacion: v(c.fecha_ingreso),
    sueldoBase: parseFloat(String(c.sueldo_liquido).replace(/[^\d]/g, '')) || 0,
    afp: afpId,
    isapre: isapreId,
    sueldoBanco: bancoId,
    sueldoCuentaCorriente: v(bn.numero),
    descripcionDelCargo: v(rol.descripcion),
    documentoEsContratoOAnexo: 'contrato',
    disabilities: 'no'
  };

  const respContrato = talanaPost('/contrato/', contratoBody);
  dbgLog(ss, 'Talana POST /contrato/ → ' + respContrato.status + ' | ' + JSON.stringify(respContrato.body));
  if (respContrato.status !== 201) {
    throw new Error('Talana falló al crear contrato: ' + JSON.stringify(respContrato.body));
  }

  dbgLog(ss, 'Talana: persona + contrato creados OK. Firma digital pendiente (falta resolver generación de PDF).');
}

// Prueba manual con datos ficticios claramente marcados — ojo: esto escribe en la base PRODUCTIVA real de Talana
function testTalanaCrearPersonaYContrato() {
  const data = {
    recruiter: { nombres: 'PRUEBA', apellido1: 'INTEGRACION', apellido2: 'TEST', email_personal: 'prueba.integracion@example.com', numero_documento: '11111111-1', fecha_nacimiento: '1990-01-01', sexo: 'Masculino' },
    contrato:  { tipo: 'Chile-Nomina', sueldo_liquido: '1000000', moneda: 'CLP', fecha_ingreso: '2026-07-01' },
    rol:       { cargo: 'Cargo Prueba', seniority: 'Test', descripcion: 'Descripción de prueba — no es un trabajador real.' },
    ficha:     { nombres: 'PRUEBA', apellidos: { primero: 'INTEGRACION', segundo: 'TEST' }, prevision: {}, banco: {} },
    metadata:  { correo_corporativo: 'prueba.integracion@continuumhq.com' }
  };
  enviarContratoTalana(data);
  Logger.log('testTalanaCrearPersonaYContrato completado — revisa la pestaña _debug');
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
