/**
 * ============================================================================
 *  DEPOAPP — Backend de sincronización + notificaciones push
 *  (versión con ENTREGA que viaja entre celulares para la liquidación)
 * ============================================================================
 *  NOVEDAD de esta versión:
 *   - Sincroniza el estado de ENTREGA de cada partido entre todos los celulares.
 *     Cuando el camarógrafo marca "Entregado" (o el editor marca su parte),
 *     eso queda guardado en el Excel y lo ve el jefe en la liquidación.
 *
 *  ANTES DE PEGAR ESTE CÓDIGO, en la hoja "Asignaciones" del Excel:
 *   - Ya tenés una columna "ENTREGADO" (la usamos para el camarógrafo).
 *   - AGREGÁ una columna nueva con el título EXACTO:  EDITADO
 *     (ponela al lado de ENTREGADO; el título va en la fila 1).
 *
 *  Después: Implementar > Gestionar implementaciones > (lápiz) editar >
 *     Versión: "Nueva versión" > Implementar.
 *  La URL NO cambia. No hace falta tocar nada más.
 * ============================================================================
 */

/* Tu Google Sheet "Sistema Partidos" */
var SHEET_ID = "1Nz2Sy43E3jDausJBCsijzq7QzZu9t6jV5Ha3YhS1Ow4";
var HOJA = "Asignaciones";

/* Columnas de la hoja Asignaciones que usa la app */
var COL = {
  fecha: "FECHA",
  hora: "HORA",
  cancha: "CANCHA",
  club: "CLUB",
  cat: "CATEGORIA",
  precio: "PRECIO",
  emp: "EMPLEADO ASIGNADO",
  estado: "ESTADO DE ASIGNACION",
  obs: "OBSERVACIONES",
  id: "ID PARTIDO",
  edi: "EDICION",
  entregado: "ENTREGADO",   /* camarógrafo entregó su parte  -> SI / NO */
  editado: "EDITADO",       /* editor entregó su parte       -> SI / NO  (columna NUEVA) */
  pagado: "PAGADO"
};

function prop(nombre) {
  return PropertiesService.getScriptProperties().getProperty(nombre) || "";
}

function hoja() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var h = ss.getSheetByName(HOJA);
  if (!h) throw new Error("No encuentro la hoja " + HOJA);
  return h;
}

function encabezados(h) {
  return h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map(function (x) {
    return String(x).trim();
  });
}

function idx(heads, nombre) {
  return heads.indexOf(nombre);
}

/* --------------------------------------------------------------------------
 *  LEER  ->  GET  ?action=partidos&token=TU_TOKEN
 * ------------------------------------------------------------------------ */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.token !== prop("TOKEN")) return json({ ok: false, error: "token invalido" });

  if (p.action === "partidos") {
    var h = hoja();
    var heads = encabezados(h);
    var last = h.getLastRow();
    if (last < 2) return json({ ok: true, partidos: [] });

    var datos = h.getRange(2, 1, last - 1, h.getLastColumn()).getValues();

    /* Auto-ID: a toda fila con FECHA pero sin ID PARTIDO le pone un ID estable
       y lo escribe de vuelta en la hoja. */
    var cId = idx(heads, COL.id), cFe = idx(heads, COL.fecha);
    if (cId >= 0) {
      var faltan = false, colId = [], base = new Date().getTime();
      for (var k = 0; k < datos.length; k++) {
        var v = String(datos[k][cId] || "").trim();
        if (!v && datos[k][cFe]) { v = "DA" + (base + k); datos[k][cId] = v; faltan = true; }
        colId.push([datos[k][cId]]);
      }
      if (faltan) { h.getRange(2, cId + 1, datos.length, 1).setValues(colId); SpreadsheetApp.flush(); }
    }

    var cEnt = idx(heads, COL.entregado), cEdt = idx(heads, COL.editado);
    var out = [];
    for (var i = 0; i < datos.length; i++) {
      var r = datos[i];
      if (!r[idx(heads, COL.fecha)]) continue;
      out.push({
        fila: i + 2,
        id: String(r[idx(heads, COL.id)] || ""),
        fecha: fechaISO(r[idx(heads, COL.fecha)]),
        hora: String(r[idx(heads, COL.hora)] || ""),
        cancha: String(r[idx(heads, COL.cancha)] || ""),
        club: String(r[idx(heads, COL.club)] || ""),
        cat: String(r[idx(heads, COL.cat)] || ""),
        precio: numero(r[idx(heads, COL.precio)]),
        emp: String(r[idx(heads, COL.emp)] || ""),
        estado: String(r[idx(heads, COL.estado)] || ""),
        edi: String(r[idx(heads, COL.edi)] || "").toUpperCase() === "SI",
        entregado: cEnt >= 0 ? String(r[cEnt] || "").toUpperCase() === "SI" : false,
        editado: cEdt >= 0 ? String(r[cEdt] || "").toUpperCase() === "SI" : false,
        pagado: String(r[idx(heads, COL.pagado)] || "").toUpperCase() === "SI"
      });
    }
    return json({ ok: true, partidos: out, total: out.length });
  }

  /* GUARDAR por GET (más confiable que POST desde el navegador).
     La app manda ?action=guardar&token=...&data=<json del partido> */
  if (p.action === "guardar") {
    try {
      var partido = JSON.parse(p.data || "{}");
      var res = guardarPartido(partido);
      var aviso = null;
      if (partido && partido.empId) {
        aviso = enviarPush(partido.empId, textoAviso(partido, res.creado));
      }
      /* avisar también al jefe (Gastón). Scorify manda avisaJefe:1 en cada
         alta/cambio; la app lo manda solo cuando NO es el jefe quien guarda. */
      if (partido && partido.avisaJefe && String(partido.empId) !== "jefe") {
        enviarPush("jefe", textoAvisoJefe(partido, res.creado));
      }
      return json({ ok: true, fila: res.fila, creado: res.creado, push: aviso });
    } catch (err) {
      return json({ ok: false, error: String(err) });
    }
  }

  /* LEER PINES  ->  GET ?action=pins  → { ok, pins:{ empId: hash } }
     Sirve para que el PIN de cada empleado no se pierda si el celular
     borra el almacenamiento o entra desde otro navegador. */
  if (p.action === "pins") {
    return json({ ok: true, pins: leerPins() });
  }

  /* GUARDAR UN PIN  ->  GET ?action=guardarPin&empId=..&hash=..
     hash vacío = resetear (queda sin PIN). */
  if (p.action === "guardarPin") {
    try {
      guardarPin(String(p.empId || ""), String(p.hash || ""));
      return json({ ok: true });
    } catch (err) {
      return json({ ok: false, error: String(err) });
    }
  }

  return json({ ok: false, error: "accion no reconocida" });
}

/* --------------------------------------------------------------------------
 *  PINES: hoja "Pins" (EMPID | HASH). Se crea sola si no existe.
 * ------------------------------------------------------------------------ */
function hojaPins() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var h = ss.getSheetByName("Pins");
  if (!h) {
    h = ss.insertSheet("Pins");
    h.getRange(1, 1, 1, 2).setValues([["EMPID", "HASH"]]);
  }
  return h;
}
function leerPins() {
  var h = hojaPins(), last = h.getLastRow(), o = {};
  if (last < 2) return o;
  var d = h.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < d.length; i++) {
    var id = String(d[i][0] || "").trim();
    if (id) o[id] = String(d[i][1] || "");
  }
  return o;
}
function guardarPin(empId, hash) {
  if (!empId) throw new Error("falta empId");
  var h = hojaPins(), last = h.getLastRow(), fila = 0;
  if (last > 1) {
    var ids = h.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === String(empId)) { fila = i + 2; break; }
    }
  }
  if (!fila) { fila = last + 1; h.getRange(fila, 1).setValue(empId); }
  h.getRange(fila, 2).setValue(hash || "");
  return fila;
}

/* --------------------------------------------------------------------------
 *  GUARDAR  ->  POST  {token, action:"guardarPartido", partido:{...}}
 * ------------------------------------------------------------------------ */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.token !== prop("TOKEN")) return json({ ok: false, error: "token invalido" });

    if (body.action === "guardarPartido") {
      var res = guardarPartido(body.partido);
      var aviso = null;
      if (body.partido && body.partido.emp) {
        aviso = enviarPush(body.partido.emp, textoAviso(body.partido, res.creado));
      }
      return json({ ok: true, fila: res.fila, creado: res.creado, push: aviso });
    }

    if (body.action === "borrarPartido") {
      var h = hoja();
      if (body.fila && body.fila > 1) h.deleteRow(body.fila);
      return json({ ok: true });
    }

    /* SCORIFY: sync_depoapp.py manda los partidos de la semana scrapeados.
       Crea los nuevos como SIN ASIGNAR y, en los que ya existen, corrige
       SOLO fecha/hora/cancha. Nunca pisa empleado, editor, precio, etc. */
    if (body.action === "scorify") {
      return json(sincronizarScorify(body.partidos || []));
    }

    return json({ ok: false, error: "accion no reconocida" });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/* --------------------------------------------------------------------------
 *  SCORIFY -> escribe/actualiza los partidos scrapeados en "Asignaciones"
 * ------------------------------------------------------------------------ */
function sincronizarScorify(partidos) {
  var h = hoja();
  var heads = encabezados(h);
  var cId = idx(heads, COL.id), cFe = idx(heads, COL.fecha), cHo = idx(heads, COL.hora),
      cCa = idx(heads, COL.cancha), cCl = idx(heads, COL.club), cCt = idx(heads, COL.cat),
      cEm = idx(heads, COL.emp), cEs = idx(heads, COL.estado);

  var last = h.getLastRow();
  var idFila = {};
  if (last > 1 && cId >= 0) {
    var ids = h.getRange(2, cId + 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) idFila[String(ids[i][0]).trim()] = i + 2;
  }

  var filaLibre = last + 1; /* dónde va la próxima alta (sin depender de getLastRow en el loop) */
  var detalle = [], procesados = 0;
  for (var k = 0; k < partidos.length; k++) {
    var p = partidos[k];
    if (!p || !p.id || !p.fecha) continue;
    procesados++;

    var fila = idFila[String(p.id).trim()] || 0;
    var creado = false, cambios = false, push = null;

    if (!fila) {
      /* ALTA: partido nuevo, SIN ASIGNAR (para que vos le pongas el empleado) */
      fila = filaLibre; filaLibre++;
      setCell(h, fila, cFe, p.fecha);
      setCell(h, fila, cHo, p.hora || "");
      setCell(h, fila, cCa, p.cancha || "");
      setCell(h, fila, cCl, p.club || "");
      setCell(h, fila, cCt, p.cat || "");
      setCell(h, fila, cEs, "SIN ASIGNAR");
      setCell(h, fila, cId, p.id);
      idFila[String(p.id).trim()] = fila;
      creado = true;
    } else {
      /* Ya existe: corrige SOLO fecha/hora/cancha. No toca lo que cargaste vos. */
      var actual = h.getRange(fila, 1, 1, h.getLastColumn()).getValues()[0];
      if (fechaISO(actual[cFe]) !== p.fecha)          { setCell(h, fila, cFe, p.fecha);        cambios = true; }
      if (String(actual[cHo] || "") !== String(p.hora || ""))     { setCell(h, fila, cHo, p.hora || "");   cambios = true; }
      if (String(actual[cCa] || "") !== String(p.cancha || ""))   { setCell(h, fila, cCa, p.cancha || ""); cambios = true; }

      /* Si cambió algo y el partido ya estaba asignado, avisar al empleado y al jefe */
      if (cambios && cEm >= 0) {
        var nombreEmp = String(actual[cEm] || "").trim();
        var empId = empIdDeNombre(nombreEmp);
        var datos = { club: p.club, cat: p.cat, fecha: p.fecha, hora: p.hora, cancha: p.cancha, emp: nombreEmp };
        if (empId) {
          push = enviarPush(empId, textoAviso(datos, false));
          enviarPush("jefe", textoAvisoJefe(datos, false));
        }
      }
    }
    detalle.push({ id: p.id, creado: creado, cambios: cambios, push: push });
  }
  SpreadsheetApp.flush();
  return { ok: true, procesados: procesados, detalle: detalle };
}

function setCell(h, fila, c, valor) {
  if (c >= 0) h.getRange(fila, c + 1).setValue(valor);
}

/* "ENZO NUÑEZ" -> "enzo"  ·  "GASTON CHECHI" -> "jefe" (para el push) */
function empIdDeNombre(nombre) {
  var t = String(nombre || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase();
  t = t.replace(/[̀-ͯ]/g, "");
  if (!t) return "";
  if (t.indexOf("GASTON") > -1) return "jefe";
  var mapa = { ENZO: "enzo", LUCA: "luca", AGUS: "agus", SANTI: "santi", FACU: "facu", GIU: "giu", DANA: "dana" };
  for (var key in mapa) { if (t.indexOf(key) > -1) return mapa[key]; }
  return "";
}

/* Escribe (o actualiza) una fila en Asignaciones. Devuelve la fila y si fue alta. */
function guardarPartido(p) {
  var h = hoja();
  var heads = encabezados(h);
  var last = h.getLastRow();

  /* Buscar por ID PARTIDO */
  var fila = 0;
  if (p.id) {
    var ids = last > 1 ? h.getRange(2, idx(heads, COL.id) + 1, last - 1, 1).getValues() : [];
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(p.id)) { fila = i + 2; break; }
    }
  }
  var creado = false;
  if (!fila) { fila = last + 1; creado = true; if (!p.id) p.id = "DA" + new Date().getTime(); }

  function set(col, valor) {
    var c = idx(heads, col);
    if (c >= 0) h.getRange(fila, c + 1).setValue(valor);
  }

  /* Entrega: la app manda entregado/editado (bool) o el estado en texto */
  var ent = (p.entregado === true) || (p.estado === "entreg");
  var edt = (p.editado === true) || (p.estadoEdit === "entreg");

  set(COL.fecha, p.fecha || "");
  set(COL.hora, p.hora || "");
  set(COL.cancha, p.cancha || "");
  set(COL.club, p.club || "");
  set(COL.cat, p.cat || "");
  set(COL.precio, p.precio || 0);
  set(COL.emp, p.emp || "");
  set(COL.estado, p.emp ? "ASIGNADO" : "SIN ASIGNAR");
  set(COL.edi, p.edi ? "SI" : "NO");
  set(COL.entregado, ent ? "SI" : "NO");
  set(COL.editado, edt ? "SI" : "NO");
  set(COL.pagado, p.pagado ? "SI" : "NO");
  set(COL.id, p.id);

  return { fila: fila, creado: creado };
}

function textoAviso(p, esNuevo) {
  var d = p.fecha ? String(p.fecha).split("-").reverse().slice(0, 2).join("/") : "";
  return (esNuevo ? "Nuevo partido: " : "Cambió tu partido: ") +
    (p.club || "") + (p.cat ? " (" + p.cat + ")" : "") +
    (d ? " · " + d : "") + (p.hora ? " " + p.hora : "") +
    (p.cancha ? " · " + p.cancha : "");
}

/* Aviso para el jefe: incluye a qué empleado corresponde el partido. */
function textoAvisoJefe(p, esNuevo) {
  var d = p.fecha ? String(p.fecha).split("-").reverse().slice(0, 2).join("/") : "";
  return (esNuevo ? "Nuevo partido" : "Cambió un partido") +
    (p.emp ? " · " + p.emp : "") + ": " + (p.club || "") +
    (p.cat ? " (" + p.cat + ")" : "") +
    (d ? " · " + d : "") + (p.hora ? " " + p.hora : "") +
    (p.cancha ? " · " + p.cancha : "");
}

/* --------------------------------------------------------------------------
 *  PUSH: le llega SOLO al empleado indicado (external_id = id del empleado)
 * ------------------------------------------------------------------------ */
function enviarPush(empId, mensaje) {
  var appId = prop("ONESIGNAL_APP_ID");
  var restKey = prop("ONESIGNAL_REST_KEY");
  if (!appId || !restKey) return "faltan claves de OneSignal";

  var payload = {
    app_id: appId,
    target_channel: "push",
    include_aliases: { external_id: [String(empId)] },
    headings: { en: "DepoApp", es: "DepoApp" },
    contents: { en: mensaje, es: mensaje }
  };

  try {
    var resp = UrlFetchApp.fetch("https://api.onesignal.com/notifications", {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Basic " + restKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    return resp.getResponseCode() === 200 ? "enviado" : "error " + resp.getContentText();
  } catch (err) {
    return "error " + String(err);
  }
}

/* -------------------------- utilidades ---------------------------------- */
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function numero(v) {
  var n = parseFloat(String(v).replace(/[$\s]/g, "").replace(/,/g, ""));
  return isNaN(n) ? 0 : n;
}
function fechaISO(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  var s = String(v).trim();
  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    var a = m[3].length === 2 ? "20" + m[3] : m[3];
    return a + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[1]).slice(-2);
  }
  return s;
}

function testPush() {
  Logger.log(enviarPush("luca", "Prueba desde DepoApp ✅"));
}
