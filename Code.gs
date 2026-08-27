/* =========================================================
   CLÍNICA ODONTOLÓGICA — Ma. Eugenia Blanco
   Backend Apps Script — patrón JSONP (mismo esquema que
   BC Salud y GG Taller: Google Sheets como base de datos,
   Apps Script como API, GitHub Pages como frontend).

   INSTALACIÓN:
   1) Creá una Google Sheet nueva (o abrí la que ya armaste con
      las 4 hojas).
   2) Extensiones > Apps Script.
   3) Borrá el contenido de Code.gs y pegá todo este archivo.
   4) Ejecutá la función "inicializarHojas" una vez (desde el
      selector de funciones arriba > Ejecutar). Esto crea las
      4 hojas con los encabezados correctos si no existen.
   5) Implementar > Nueva implementación > tipo "Aplicación web".
      Ejecutar como: Yo. Quién tiene acceso: Cualquiera.
   6) Copiá la URL que te da ("Web app URL") y pegala en
      config.js del frontend (CLINICA_API_URL).
   ========================================================= */

const HOJAS = {
  PACIENTES: 'Pacientes',
  TRATAMIENTOS: 'Tratamientos',
  CONSULTAS: 'Consultas',
  PLAN: 'Plan'
};

const ENCABEZADOS = {
  Pacientes: ['id','historiaNro','nombre','apellido','dni','fechaNacimiento','sexo','telefono','whatsapp','email','domicilio','localidad','provincia','contactoEmergencia','obraSocial','numeroAfiliado','observaciones','diabetes','hipertension','cardiopatias','coagulacion','embarazo','alergiasTiene','alergiasDetalle','medicacionTiene','medicacionDetalle','cirugiasTiene','cirugiasDetalle','otrasTiene','otrasDetalle','observacionesMedicas','piezasAusentes','createdAt','updatedAt'],
  Tratamientos: ['id','pacienteId','pieza','fecha','procedimiento','superficies','profesional','diagnostico','descripcion','material','estado','observaciones','costo','garantiaTiene','garantiaInicio','garantiaVencimiento','createdAt'],
  Consultas: ['id','pacienteId','fecha','motivo','diagnostico','piezas','observaciones','indicaciones','proximoControl','createdAt'],
  Plan: ['id','pacienteId','pieza','procedimiento','estado','createdAt']
};

function inicializarHojas(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(ENCABEZADOS).forEach(nombre=>{
    let hoja = ss.getSheetByName(nombre);
    if(!hoja){ hoja = ss.insertSheet(nombre); }
    if(hoja.getLastRow() === 0){
      hoja.getRange(1,1,1,ENCABEZADOS[nombre].length).setValues([ENCABEZADOS[nombre]]);
      hoja.setFrozenRows(1);
    }
  });
  // Elimina la hoja "Hoja 1" / "Sheet1" default si quedó vacía
  const def = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if(def && def.getLastRow() === 0 && ss.getSheets().length > 1){ ss.deleteSheet(def); }
}

/* ---------- Helpers genéricos ---------- */
function sh(nombre){ return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre); }

function headersDe(hoja){ return hoja.getRange(1,1,1,hoja.getLastColumn()).getValues()[0]; }

/* Google Sheets convierte solo texto que "parece" número o fecha (DNI, teléfonos,
   fechas) en tipos nativos. Esto normaliza esos valores de vuelta a texto legible
   antes de mandarlos al frontend, sin importar cómo haya quedado guardada la celda. */
function normalizarValor(valor, header){
  if(Object.prototype.toString.call(valor) === '[object Date]'){
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    if(header === 'createdAt' || header === 'updatedAt'){
      return Utilities.formatDate(valor, tz, "yyyy-MM-dd'T'HH:mm:ss");
    }
    return Utilities.formatDate(valor, tz, "yyyy-MM-dd");
  }
  if(typeof valor === 'boolean') return valor;
  if(typeof valor === 'number') return String(valor);
  return valor;
}

function filasAObjetos(hoja){
  const datos = hoja.getDataRange().getValues();
  if(datos.length < 2) return [];
  const heads = datos[0];
  return datos.slice(1)
    .filter(f => f[0] !== '' && f[0] !== null)
    .map(f => { const o = {}; heads.forEach((h,i)=> o[h] = normalizarValor(f[i], h)); return o; });
}

function agregarObjeto(hoja, obj){
  const heads = headersDe(hoja);
  const fila = heads.map(h => obj[h] !== undefined ? obj[h] : '');
  hoja.appendRow(fila);
}

function buscarFilaPorId(hoja, id){
  const datos = hoja.getDataRange().getValues();
  for(let i=1;i<datos.length;i++){
    if(String(datos[i][0]) === String(id)) return i+1; // 1-based row number
  }
  return -1;
}

function actualizarObjetoPorId(hoja, id, patch){
  const heads = headersDe(hoja);
  const fila = buscarFilaPorId(hoja, id);
  if(fila < 0) return false;
  const rango = hoja.getRange(fila,1,1,heads.length);
  const actual = rango.getValues()[0];
  heads.forEach((h,i)=>{ if(patch[h] !== undefined) actual[i] = patch[h]; });
  rango.setValues([actual]);
  return true;
}

function generarId(prefijo){
  return prefijo + Utilities.getUuid().slice(0,8);
}

function siguienteHistoriaNro(hoja){
  const datos = hoja.getDataRange().getValues();
  let max = 0;
  for(let i=1;i<datos.length;i++){
    const n = Number(datos[i][1]) || 0;
    if(n > max) max = n;
  }
  return max + 1;
}

function tocarPaciente(id){
  actualizarObjetoPorId(sh(HOJAS.PACIENTES), id, {updatedAt: new Date().toISOString()});
}

/* ---------- Acciones ---------- */
function getTodo(){
  return {
    pacientes: filasAObjetos(sh(HOJAS.PACIENTES)),
    tratamientos: filasAObjetos(sh(HOJAS.TRATAMIENTOS)),
    consultas: filasAObjetos(sh(HOJAS.CONSULTAS)),
    plan: filasAObjetos(sh(HOJAS.PLAN))
  };
}

function crearPaciente(data){
  const hoja = sh(HOJAS.PACIENTES);
  const ahora = new Date().toISOString();
  const obj = Object.assign({
    id: generarId('pac_'),
    historiaNro: siguienteHistoriaNro(hoja),
    piezasAusentes: '',
    createdAt: ahora,
    updatedAt: ahora
  }, data);
  agregarObjeto(hoja, obj);
  return obj;
}

function actualizarPaciente(id, data){
  data.updatedAt = new Date().toISOString();
  const ok = actualizarObjetoPorId(sh(HOJAS.PACIENTES), id, data);
  if(!ok) return {error:'Paciente no encontrado.'};
  return {ok:true};
}

function agregarTratamiento(data){
  const hoja = sh(HOJAS.TRATAMIENTOS);
  const obj = Object.assign({id: generarId('tr_'), createdAt: new Date().toISOString()}, data);
  agregarObjeto(hoja, obj);
  if(data.pacienteId) tocarPaciente(data.pacienteId);
  return obj;
}

function agregarConsulta(data){
  const hoja = sh(HOJAS.CONSULTAS);
  const obj = Object.assign({id: generarId('con_'), createdAt: new Date().toISOString()}, data);
  agregarObjeto(hoja, obj);
  if(data.pacienteId) tocarPaciente(data.pacienteId);
  return obj;
}

function agregarPlanItem(data){
  const hoja = sh(HOJAS.PLAN);
  const obj = Object.assign({id: generarId('pl_'), createdAt: new Date().toISOString()}, data);
  agregarObjeto(hoja, obj);
  if(data.pacienteId) tocarPaciente(data.pacienteId);
  return obj;
}

function actualizarPlanEstado(id, estado){
  const ok = actualizarObjetoPorId(sh(HOJAS.PLAN), id, {estado});
  if(!ok) return {error:'Ítem de plan no encontrado.'};
  return {ok:true};
}

function setPiezaAusente(pacienteId, pieza, valor){
  const hoja = sh(HOJAS.PACIENTES);
  const fila = buscarFilaPorId(hoja, pacienteId);
  if(fila < 0) return {error:'Paciente no encontrado.'};
  const heads = headersDe(hoja);
  const col = heads.indexOf('piezasAusentes') + 1;
  const celda = hoja.getRange(fila, col);
  let lista = (celda.getValue() || '').toString().split(',').map(s=>s.trim()).filter(Boolean);
  const piezaStr = String(pieza);
  if(valor){
    if(!lista.includes(piezaStr)) lista.push(piezaStr);
  } else {
    lista = lista.filter(x => x !== piezaStr);
  }
  celda.setValue(lista.join(','));
  tocarPaciente(pacienteId);
  return {ok:true, piezasAusentes: lista.join(',')};
}

/* ---------- Entrada HTTP (JSONP) ---------- */
function doGet(e){
  let resultado;
  try{
    const accion = e.parameter.accion;
    const data = e.parameter.data ? JSON.parse(e.parameter.data) : null;
    switch(accion){
      case 'ping': resultado = {ok:true, ts:new Date().toISOString()}; break;
      case 'getTodo': resultado = getTodo(); break;
      case 'crearPaciente': resultado = crearPaciente(data); break;
      case 'actualizarPaciente': resultado = actualizarPaciente(e.parameter.id, data); break;
      case 'agregarTratamiento': resultado = agregarTratamiento(data); break;
      case 'agregarConsulta': resultado = agregarConsulta(data); break;
      case 'agregarPlanItem': resultado = agregarPlanItem(data); break;
      case 'actualizarPlanEstado': resultado = actualizarPlanEstado(e.parameter.id, e.parameter.estado); break;
      case 'setPiezaAusente': resultado = setPiezaAusente(e.parameter.pacienteId, e.parameter.pieza, e.parameter.valor === 'true'); break;
      default: resultado = {error: 'Acción desconocida: ' + accion};
    }
  } catch(err){
    resultado = {error: err.toString()};
  }
  return construirRespuesta(resultado, e.parameter.callback);
}

function construirRespuesta(data, callback){
  const json = JSON.stringify(data);
  if(callback){
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
