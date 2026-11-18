/**
 * Backend del onboarding TT Audit (Google Apps Script).
 *
 * PARTE 1 (doPost sin "action"): recibe el formulario, guarda archivos en Drive
 *   y agrega/actualiza la fila en el Excel correcto (por NOMBRE de encabezado).
 *   Crea al final las columnas: "Perfil", "Aceptó horario", "Alta Buk",
 *   "Contrato (link Drive)", "Capacitación Buk", "Contrato descargado".
 *
 * RRHH, cuando termine el alta en Buk y tenga el contrato:
 *   1) escribe "Sí" en la columna "Alta Buk"
 *   2) pega el link de Drive del contrato en "Contrato (link Drive)"
 *      (el archivo puede quedar PRIVADO: el script lo lee y lo entrega solo
 *       a quien verifique DNI + fecha de nacimiento).
 *
 * PARTE 2 (doPost con action:"estado"): el ingresante pone DNI + fecha de
 *   nacimiento; si coincide y "Alta Buk" = Sí, devuelve el contrato y habilita
 *   la capacitación.
 *
 * Despliegue (en cada URL que uses): Implementar > Administrar implementaciones >
 *   editar > Versión: Nueva versión.  Ejecutar como: Yo | Acceso: Cualquier usuario.
 *   El dueño del script debe tener acceso a AMBOS Excel y a las carpetas de Drive.
 */
const SHEETS = {
  contratado:  {id: "1Ff5Q09Y9akdetiyu5ywvLiBB894ZvO2GxL3rfYKOZvc", tab: "Respuestas de formulario 1", dniCol: "Nº DNI", nameCol: "Nombre y Apellidos Completos"},
  practicante: {id: "1SCifr4EwgftkWMcC5SzQDQnMTCEyzaoyRFeeDbA7DdE", tab: "Respuestas de formulario 1", dniCol: "N° de DNI/CE", nameCol: "Nombre y Apellidos completos"},
};
const CARPETA_ARCHIVOS_ID = "PEGAR_ID_CARPETA_DRIVE"; // carpeta privada: DNI, recibos, fotos
const COL_ALTA = "Alta Buk", COL_CONTRATO = "Contrato (link Drive)";
const MAX_CONTRATO_BYTES = 8 * 1024 * 1024;

function doGet() { return ContentService.createTextOutput("TT Audit onboarding OK"); }

function doPost(e) {
  const p = JSON.parse(e.postData.contents);
  if (p.action === "estado") return json(estado(p));
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { guardar(p); return ContentService.createTextOutput("ok"); }
  finally { lock.releaseLock(); }
}

function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function norm(s) { return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").trim().toLowerCase(); }
function pad(n) { return ("0" + n).slice(-2); }
function fechaISO(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const s = String(v).trim(); let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + "-" + pad(m[2]) + "-" + pad(m[3]);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? m[3] + "-" + pad(m[2]) + "-" + pad(m[1]) : s;
}

/* ---------- PARTE 1: guardar ---------- */
function guardar(p) {
  const cfg = SHEETS[p.hoja]; if (!cfg) throw new Error("hoja inválida");
  const sh = SpreadsheetApp.openById(cfg.id).getSheetByName(cfg.tab);
  const row = Object.assign({}, p.data);

  if (Object.keys(p.files || {}).length) {
    const nombre = row[cfg.nameCol] || "SIN_NOMBRE";
    const sub = DriveApp.getFolderById(CARPETA_ARCHIVOS_ID).createFolder((p.dni || "sin-dni") + " - " + nombre);
    Object.keys(p.files).forEach(col => {
      row[col] = p.files[col].map(f => {
        const blob = Utilities.newBlob(Utilities.base64Decode(f.b64), f.type, col.slice(0, 20) + " - " + f.name);
        return sub.createFile(blob).getUrl();
      }).join(", ");
    });
  }

  let headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const nuevo = !!p.data["Aceptó horario"]; // viene de la Parte 1 (formulario completo)
  if (nuevo) { if (row[COL_ALTA] === undefined) row[COL_ALTA] = ""; if (row[COL_CONTRATO] === undefined) row[COL_CONTRATO] = ""; }
  Object.keys(row).forEach(k => { if (headers.indexOf(k) < 0) { sh.getRange(1, headers.length + 1).setValue(k); headers.push(k); } });
  if (nuevo) row["Marca temporal"] = new Date();

  // misma persona re-enviando (o Parte 2 actualizando seguimiento) => actualiza su fila
  const dniCol = headers.indexOf(cfg.dniCol); let target = -1;
  if (dniCol >= 0 && p.dni && sh.getLastRow() > 1) {
    const vals = sh.getRange(2, dniCol + 1, sh.getLastRow() - 1, 1).getValues().flat().map(v => String(v).replace(/\s/g, ""));
    const i = vals.lastIndexOf(String(p.dni)); if (i >= 0) target = i + 2;
  }
  const out = headers.map(h => row[h] !== undefined ? row[h] : "");
  if (target > 0) {
    const old = sh.getRange(target, 1, 1, headers.length).getValues()[0];
    sh.getRange(target, 1, 1, headers.length).setValues([out.map((v, i) => v === "" ? old[i] : v)]);
  } else if (nuevo) sh.appendRow(out);
}

/* ---------- PARTE 2: estado + contrato ---------- */
function estado(p) {
  const dni = String(p.dni || "").replace(/\s/g, ""), nac = String(p.nac || "");
  if (!dni || !nac) return {found: false};
  for (const hoja of Object.keys(SHEETS)) {
    const cfg = SHEETS[hoja], sh = SpreadsheetApp.openById(cfg.id).getSheetByName(cfg.tab);
    if (sh.getLastRow() < 2) continue;
    const data = sh.getDataRange().getValues(), H = data[0];
    const iDni = H.indexOf(cfg.dniCol), iNac = H.indexOf("Fecha de Nacimiento");
    if (iDni < 0 || iNac < 0) continue;
    for (let r = data.length - 1; r >= 1; r--) {
      const row = data[r];
      if (String(row[iDni]).replace(/\s/g, "") !== dni || fechaISO(row[iNac]) !== nac) continue;
      const out = {found: true, hoja: hoja, perfil: String(row[H.indexOf("Perfil")] || ""), nombre: String(row[H.indexOf(cfg.nameCol)] || ""),
                   alta: norm(row[H.indexOf(COL_ALTA)]) === "si"};
      const link = String(row[H.indexOf(COL_CONTRATO)] || "");
      const m = out.alta && link.match(/[-\w]{25,}/);
      if (m) {
        try {
          const f = DriveApp.getFileById(m[0]), blob = f.getBlob();
          if (blob.getBytes().length <= MAX_CONTRATO_BYTES) {
            out.contratoB64 = Utilities.base64Encode(blob.getBytes());
            out.contratoName = f.getName(); out.contratoMime = blob.getContentType();
          }
        } catch (err) { /* sin acceso al archivo: se entrega sin contrato */ }
      }
      return out;
    }
  }
  return {found: false};
}
