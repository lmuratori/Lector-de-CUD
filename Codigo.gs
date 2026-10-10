// === SISTEMA MÉDICO MURATORI — Apps Script API ===
// Esta versión conserva TODO lo que ya tenías (Pacientes, Historia_Clínica, Recetas,
// getAll/search/ping, addPaciente/addReceta/addHistoria) y agrega:
//   - el puente real con RIDM (para que los datos lleguen solos a OCM)
//   - el arreglo del backup, que venía fallando en silencio (mandaba 'backup'
//     y el script no tenía esa acción — por eso OCM decía "OK" sin haber guardado nada)
//   - la puerta lista para kinesiología, todavía sin usar

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName('Recetas') === null) ss.getSheets()[0].setName('Recetas');
  const recetasSheet = ss.getSheetByName('Recetas');
  recetasSheet.getRange(1,1,1,recetasSheet.getLastColumn())
    .setBackground('#6B3FA0').setFontColor('white').setFontWeight('bold');

  if (ss.getSheetByName('Pacientes') === null) {
    const ps = ss.insertSheet('Pacientes');
    const ph = ['ID','Apellido','Nombre','DNI','Fecha_Nac','Obra_Social',
                'N_Afiliado','Plan','Telefono','Email','Domicilio',
                'Lugar_Atencion','Diagnostico_Principal','Estado_Tratamiento',
                'Localidad','Trabajos_JSON','Vivienda_JSON'];
    ps.getRange(1,1,1,ph.length).setValues([ph])
      .setBackground('#1A73E8').setFontColor('white').setFontWeight('bold');
    ps.setFrozenRows(1);
  }
  if (ss.getSheetByName('Historia_Clinica') === null) {
    const hs = ss.insertSheet('Historia_Clinica');
    const hh = ['ID_HC','ID_Paciente','Fecha','Diagnostico','Medicamentos',
                'Estudios_Realizados','Estudios_Pendientes','Indicaciones',
                'Observaciones','Proximo_Control'];
    hs.getRange(1,1,1,hh.length).setValues([hh])
      .setBackground('#0F9D58').setFontColor('white').setFontWeight('bold');
    hs.setFrozenRows(1);
  }
  // NUEVO: hoja de espera para los formularios RIDM que completa el paciente
  if (ss.getSheetByName('RIDM_Pendientes') === null) {
    const rs = ss.insertSheet('RIDM_Pendientes');
    const rh = ['Timestamp','DNI','Apellido','Nombre','Telefono','Email','Datos_JSON','Procesado'];
    rs.getRange(1,1,1,rh.length).setValues([rh])
      .setBackground('#D97706').setFontColor('white').setFontWeight('bold');
    rs.setFrozenRows(1);
  }
  // NUEVO: registro de backups (antes se perdían porque la acción no existía)
  if (ss.getSheetByName('Backups') === null) {
    const bs = ss.insertSheet('Backups');
    bs.getRange(1,1,1,2).setValues([['Timestamp','Resumen_JSON']])
      .setBackground('#64748B').setFontColor('white').setFontWeight('bold');
    bs.setFrozenRows(1);
  }
  // NUEVO: evoluciones de kinesiología — todavía no se usa, queda lista para cuando se sume KSM
  if (ss.getSheetByName('Kinesiologia') === null) {
    const ks = ss.insertSheet('Kinesiologia');
    ks.getRange(1,1,1,5).setValues([['Timestamp','DNI','Paciente','Region','Evolucion_JSON']])
      .setBackground('#16A34A').setFontColor('white').setFontWeight('bold');
    ks.setFrozenRows(1);
  }
  // NUEVO: evaluaciones semanales que completa el PACIENTE (para el médico o el kinesiólogo)
  if (ss.getSheetByName('Evaluaciones_Pacientes') === null) {
    const es = ss.insertSheet('Evaluaciones_Pacientes');
    es.getRange(1,1,1,6).setValues([['Timestamp','DNI','Paciente','Profesional','Region','Valores_JSON']])
      .setBackground('#7C3AED').setFontColor('white').setFontWeight('bold');
    es.setFrozenRows(1);
  }
  // NUEVO: pagos de Mercado Pago confirmados de verdad contra la API de MP (no contra una imagen)
  if (ss.getSheetByName('Pagos_Confirmados') === null) {
    const pc = ss.insertSheet('Pagos_Confirmados');
    const pch = ['Timestamp','Payment_ID','DNI','Tipo_Consulta','Monto','Estado_MP','Atendido'];
    pc.getRange(1,1,1,pch.length).setValues([pch])
      .setBackground('#0EA5E9').setFontColor('white').setFontWeight('bold');
    pc.setFrozenRows(1);
  }
  // NUEVO: registros del Lector CUD, conectados con Pacientes (agenda familiar) e Historia_Clinica
  if (ss.getSheetByName('Discapacidad_CUD') === null) {
    const dc = ss.insertSheet('Discapacidad_CUD');
    const dch = ['Timestamp','DNI','Nombre','Telefono','Fecha_Nac','Diagnostico','Codigos_CIF','Tipo',
                 'Subtipo','Modalidad','Orientacion','Fecha_Emision','Fecha_Vencimiento','Junta',
                 'Firmantes_JSON','Efectores','Tipo_Pension','Trazabilidad'];
    dc.getRange(1,1,1,dch.length).setValues([dch])
      .setBackground('#244C46').setFontColor('white').setFontWeight('bold');
    dc.setFrozenRows(1);
  }
  // NUEVO: Agenda de Turnos — reservas presenciales reales, compartidas por
  // la Consulta Virtual y por Turnos_Secretaria.html (ver reservarTurno/
  // getTurnosDisponibles). Si ya existe (porque alguna de esas dos ya la
  // creó sola en su primer uso), no se toca.
  if (ss.getSheetByName('Turnos_Agenda') === null) {
    const ta = ss.insertSheet('Turnos_Agenda');
    const tah = ['ID_Turno','Timestamp_Reserva','Consultorio','Fecha','Hora','Duracion_Min',
                 'DNI','Nombre','Telefono','Email','Origen','Estado','Notas'];
    ta.getRange(1,1,1,tah.length).setValues([tah])
      .setBackground('#0c4a6e').setFontColor('white').setFontWeight('bold');
    ta.setFrozenRows(1);
  }
  recetasSheet.setFrozenRows(1);
  SpreadsheetApp.getUi().alert('✅ ¡Hojas configuradas! Ahora: Implementar > Nueva implementación');
}

// ══════════════════════════════════════════════════════════════════
// PARCHE — Token compartido para las acciones que devuelven datos
// masivos/sensibles de pacientes. Antes CUALQUIERA con la URL del Web App
// podía llamar get_vista_consolidada_todos, search, get_historial, etc.
// sin identificarse. No es autenticación por usuario (todas las apps
// comparten el mismo token), pero corta el riesgo de que la URL se filtre
// y cualquiera pueda leer datos de todos los pacientes.
//
// A PROPÓSITO queda "abierto" (sin exigir token) hasta que corras
// configurarAppToken() una vez — así no se rompe nada mientras vas
// actualizando cada app con el token nuevo. Una vez configurado, se exige
// en todas las acciones de la lista de abajo.
// ══════════════════════════════════════════════════════════════════
const ACCIONES_GET_PROTEGIDAS_ = [
  'getAll','search','get_pendientes','get_historial','get_pagos_pendientes',
  'getRecomendaciones','getAutoevaluacionesRMYT','getConsultas','get_pacientes_incompletos',
  'get_vista_consolidada','get_vista_consolidada_todos','getMensajesPendientes',
  'get_ridm_paciente',
  'get_preguntas_kine','get_resumen_facturacion_sistema',
  'get_permisos_especialista','get_lista_especialistas',
  'get_investigacion_paciente','get_discapacidad_paciente','get_plantillas_paciente',
  'get_minirimp','ssq_obtener_seguimiento_posqx','ssq_resumen_seguridad_quirofano',
  'get_neuroortopedia','get_todas_plantillas_sync',
  // AGREGADA (revisión sep 2026): confirma/anula una fila del puente RIDM →
  // OCM. Antes cualquiera con la URL podía "marcar como procesada" una
  // filiación real sin haberla visto, haciendo que se pierda de la lista
  // de pendientes del médico.
  'marcar_procesado',
  'get_gonio_mediciones',
  // AGREGADA (Membresía de Seguimiento): consulta el estado de la
  // membresía de seguimiento de un DNI (patología, próxima renovación,
  // mensajes usados esta semana) — la usa el botón "📆 Membresía de
  // seguimiento" de OCM. La llama siempre el médico.
  'get_membresia_seguimiento',
  'get_programacion_qx','get_seguimiento','get_membresias_lista','get_rmyt_dashboard','get_pendientes_qx','get_panel_hc',
  // KEP (oct 2026): evaluaciones, mediciones y derivaciones de kinesiología — datos de pacientes
  'get_kep','get_kep_dashboard',
  // Página de salud (oct 2026): temas y códigos de acceso de cada paciente
  'get_pagina_salud_temas','get_pagina_salud_accesos',
  'get_kin_hc','get_kin_ordenes'
  // NO van acá get_resumen_para_firma / get_estado_firma /
  // get_consentimiento_para_firma / get_estado_firma_consentimiento: esas
  // las llama el PACIENTE (sin token compartido) desde el link que recibe
  // por mail, así que exigirles el token de esta lista las rompería en
  // cuanto se configure. Tienen su PROPIO mecanismo de protección —un
  // Firma_Token distinto por cada documento— ver el comentario arriba de
  // guardarResumenParaFirma en la sección de firma táctil.
];
function configurarAppToken(){
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('APP_TOKEN');
  if(!token){
    token = Utilities.getUuid();
    props.setProperty('APP_TOKEN', token);
  }
  MailApp.sendEmail('lmuratori@gmail.com', 'SIM-M — Token de acceso de las apps',
    'Este es el token que hay que pegar UNA vez en cada app (OCM, KSM, Panel de Especialistas, ' +
    'Agenda de Mensajes, Colegas) para que sigan funcionando una vez que el token quede exigido:\n\n' +
    token + '\n\nHasta que lo pegues en todas, el sistema sigue funcionando igual que antes (no se corta nada).');
  Logger.log('APP_TOKEN: ' + token);
}
function getAppToken_(){
  return PropertiesService.getScriptProperties().getProperty('APP_TOKEN') || '';
}
// Si todavía no se configuró ningún token (getAppToken_ vacío), no bloquea
// nada — es el modo "todavía no migrado". Una vez configurado, exige que
// coincida exactamente.
function tokenValido_(tokenRecibido){
  const real = getAppToken_();
  if(!real) return true;
  return String(tokenRecibido||'')===String(real);
}

// ══════════════════════════════════════════════════════════════════
// TRIAL — vencimiento controlado desde ESTE backend (no desde el HTML que
// se le entrega a un colega). Por default está DESACTIVADO: no cambia nada
// en el sistema del Dr. Muratori a menos que él mismo corra configurarTrial()
// desde el editor de Apps Script. Pensado para dar una copia de prueba: se
// clona este proyecto (Sheet + Apps Script) para el destinatario, y en ESA
// copia (no en la principal) se activa el trial con fecha de vencimiento.
// Una vez vencido, TODAS las acciones (menos check_trial) quedan bloqueadas
// server-side — no alcanza con editar el HTML del lado del cliente para
// esquivarlo, porque el backend es el que corta.
// ══════════════════════════════════════════════════════════════════
function configurarTrial(fechaVencimientoISO){
  // fechaVencimientoISO: 'AAAA-MM-DD', ej: '2026-09-15' — vence a las 23:59:59
  // de ese día. Para un trial de N horas EXACTAS a partir de ahora (ej: las
  // típicas 72hs de evaluación), usar iniciarTrialPorHoras() en su lugar.
  const props = PropertiesService.getScriptProperties();
  props.setProperty('TRIAL_ACTIVO', 'true');
  props.setProperty('TRIAL_VENCIMIENTO', fechaVencimientoISO);
  Logger.log('Trial activado. Vence el ' + fechaVencimientoISO + '. Para desactivarlo, correr desactivarTrial().');
}
// ── Trial de N horas exactas desde AHORA (por defecto 72) ──
// Pensado para dar una copia de evaluación de 72hs: se corre UNA vez, desde
// el editor de Apps Script DE LA COPIA de demo (nunca en el proyecto
// principal), y a partir de ese momento arranca la cuenta regresiva —
// no hace falta calcular una fecha a mano. Guarda el vencimiento como
// fecha+hora completa (no solo el día), así el corte es preciso a la hora,
// no "antes de las 23:59 del 3er día".
function iniciarTrialPorHoras(horas){
  const h = Number(horas) || 72;
  const props = PropertiesService.getScriptProperties();
  const vencimiento = new Date(Date.now() + h*60*60*1000);
  props.setProperty('TRIAL_ACTIVO', 'true');
  props.setProperty('TRIAL_VENCIMIENTO', vencimiento.toISOString());
  Logger.log('Trial de ' + h + ' horas activado. Vence el ' + vencimiento.toISOString() + '. Para desactivarlo, correr desactivarTrial().');
}
function desactivarTrial(){
  PropertiesService.getScriptProperties().deleteProperty('TRIAL_ACTIVO');
  PropertiesService.getScriptProperties().deleteProperty('TRIAL_VENCIMIENTO');
  Logger.log('Trial desactivado — el sistema vuelve a funcionar sin límite de fecha.');
}
function estadoTrial_(){
  const props = PropertiesService.getScriptProperties();
  const activo = props.getProperty('TRIAL_ACTIVO')==='true';
  const fechaVenc = props.getProperty('TRIAL_VENCIMIENTO') || '';
  let vencido = false;
  if(activo && fechaVenc){
    const hoy = new Date();
    // Acepta tanto 'AAAA-MM-DD' (configurarTrial, vence 23:59:59 de ese día)
    // como una fecha+hora ISO completa (iniciarTrialPorHoras, vence a la
    // hora exacta) — se distinguen por si trae una 'T' de hora o no.
    const venc = fechaVenc.indexOf('T') > -1 ? new Date(fechaVenc) : new Date(fechaVenc + 'T23:59:59');
    vencido = hoy.getTime() > venc.getTime();
  }
  return {activo, vencido, fecha_vencimiento: fechaVenc};
}

function doGet(e) {
  const action=e.parameter.action||'',sheet=e.parameter.sheet||'',query=e.parameter.query||'';
  if(action==='check_trial') return jsonOut(Object.assign({success:true}, estadoTrial_()));
  const trial = estadoTrial_();
  if(trial.vencido){
    return jsonOut({success:false, error:'trial_vencido', vencido:true, mensaje:'El período de prueba finalizó el '+trial.fecha_vencimiento+'.'});
  }
  if(ACCIONES_GET_PROTEGIDAS_.indexOf(action)>-1 && !tokenValido_(e.parameter.token)){
    return jsonOut({success:false, error:'Token inválido o faltante — pedile al médico el token de acceso'});
  }
  let r;
  if(action==='getAll') r=getAllRows(sheet);
  else if(action==='search') r=searchPatient(query);
  else if(action==='ping') r={success:true,message:'API activa'};
  // ── NUEVO: puente RIDM → OCM ──
  else if(action==='get_pendientes') r=getRidmPendientes();
  else if(action==='get_ridm_paciente') r=getRidmPaciente(e.parameter.dni||'');
  else if(action==='get_historial') r=getHistorialPaciente(e.parameter.dni||'');
  else if(action==='marcar_procesado') r=marcarRidmProcesado(e.parameter.row||'');
  // ── NUEVO: Resumen de Historia Clínica con firma táctil del paciente ──
  // (sep 2026: ahora se les pasa también el token compartido —parámetro
  // "token", lo manda el médico/OCM— y el token propio del documento
  // —parámetro "t", lo manda el paciente desde el link que recibió—;
  // alcanza con que uno de los dos sea válido. Ver getResumenParaFirma.)
  else if(action==='get_resumen_para_firma') r=getResumenParaFirma(e.parameter.dni||'', e.parameter.token||'', e.parameter.t||'');
  else if(action==='get_estado_firma') r=getEstadoFirma(e.parameter.dni||'', e.parameter.token||'', e.parameter.t||'');
  // ── NUEVO: Consentimiento Informado quirúrgico con firma táctil (extiende el mismo mecanismo) ──
  else if(action==='get_consentimiento_para_firma') r=getConsentimientoParaFirma(e.parameter.dni||'', e.parameter.token||'', e.parameter.t||'');
  else if(action==='get_estado_firma_consentimiento') r=getEstadoFirmaConsentimiento(e.parameter.dni||'', e.parameter.token||'', e.parameter.t||'');
  // ── NUEVO: cola de pagos de Mercado Pago confirmados, esperando que el médico los atienda ──
  else if(action==='get_pagos_pendientes') r=getPagosPendientes();
  else if(action==='getRecomendaciones') r=getRecomendaciones(e.parameter.dni||'');
  else if(action==='getAutoevaluacionesRMYT') r=getAutoevaluacionesRMYT(e.parameter.dni||'');
  else if(action==='getConsultas') r=getConsultas(e.parameter.estado||'');
  else if(action==='getTarifas') r=getTarifas();
  else if(action==='get_pacientes_incompletos') r=getPacientesCamposIncompletos();
  else if(action==='get_vista_consolidada') r=getVistaConsolidada(e.parameter.dni||'');
  else if(action==='get_vista_consolidada_todos') r=getVistaConsolidadaTodos();
  else if(action==='getMensajesPendientes') r=getMensajesPendientes();
  else if(action==='get_preguntas_kine') r=getPreguntasKine(e.parameter.dni||'');
  else if(action==='get_peaje_sistema_pct') r=getPeajeSistemaPct();
  else if(action==='get_resumen_facturacion_sistema') r=getResumenFacturacionSistema(e.parameter.profesionalEmail||'', e.parameter.mes||'', e.parameter.origen||'');
  else if(action==='get_lista_quirurgico') r=getListaQuirurgico(e.parameter.tipo||'');
  // ── NUEVO: permisos de especialistas + módulos compartidos (Investigación/CUD/Plantillas) ──
  else if(action==='get_permisos_especialista') r=getPermisosEspecialista(e.parameter.email||'');
  else if(action==='get_lista_especialistas') r=getListaEspecialistas();
  else if(action==='get_investigacion_paciente') r=getInvestigacionPacienteGateado_(e.parameter.dni||'', e.parameter.profesionalEmail||'');
  else if(action==='get_discapacidad_paciente') r=getDiscapacidadPacienteGateado_(e.parameter.dni||'', e.parameter.profesionalEmail||'');
  else if(action==='get_plantillas_paciente') r=getPlantillasPacienteGateado_(e.parameter.dni||'', e.parameter.profesionalEmail||'');
  else if(action==='get_todas_plantillas_sync') r=getTodasPlantillasSync();
  // ── NUEVO: Mini-RIMP — interconsultas pre-quirúrgicas por especialidad ──
  else if(action==='get_minirimp_campos') r=getCamposMiniRIMP(e.parameter.especialidad||'');
  else if(action==='get_minirimp') r=getMiniRIMP(e.parameter.especialidad||'', e.parameter.dni||'');
  // ── NUEVO: Seguimiento Posquirúrgico (SSQ) — ver Revision_Integracion_SSQ_Posquirurgico.md ──
  else if(action==='ssq_obtener_seguimiento_posqx') r=getSeguimientoPosqx(e.parameter.dni||'');
  else if(action==='ssq_resumen_seguridad_quirofano') r=getResumenSeguridadQuirofano();
  // ── NUEVO: GONIO — reconecta Costa-Bártani/Valgo-Varo/Ángulos de Frente/Podoscopia con su guardado ──
  else if(action==='get_gonio_mediciones') r=getGonioMediciones(e.parameter.dni||'');
  // ── NUEVO: Membresía de Seguimiento (mensual, por patología puntual) ──
  else if(action==='get_membresia_seguimiento') r=getMembresiaSeguimiento(e.parameter.dni||'');
  // ── NUEVO: Neuro-ortopedia / Parálisis cerebral (CAM) — GMFCS, MACS,
  // Índice de Reimers, Ashworth/Tardieu, Thomas-Stonell/Greenberg. Ver
  // guardarNeuroortopedia más abajo.
  else if(action==='get_neuroortopedia') r=getNeuroortopedia(e.parameter.dni||'');
  // ── NUEVO: Agenda de Turnos — disponibilidad real, compartida entre la
  // Consulta Virtual (médico) y Turnos_Secretaria.html (cada sede). Sin
  // token a propósito, igual que getTarifas: no expone datos de pacientes,
  // solo horarios libres/ocupados.
  else if(action==='get_turnos_disponibles') r=getTurnosDisponibles(e.parameter);
  else if(action==='get_agenda_mes') r=getAgendaMes(e.parameter);
  else if(action==='get_programacion_qx') r=getProgramacionQx(e.parameter);
  else if(action==='get_seguimiento') r=getSeguimiento();
  else if(action==='get_membresias_lista') r=getMembresiasLista();
  else if(action==='get_rmyt_dashboard') r=getRmytDashboard();
  // ── NUEVO: KEP (Evaluador de Parámetros Kinésicos) — con token ──
  else if(action==='get_kep') r=getKep(e.parameter.dni||'');
  else if(action==='get_kep_dashboard') r=getKepDashboard();
  // ── NUEVO: Página de salud (códigos A + C por tema, coordinados con RMYT) — con token ──
  else if(action==='get_pagina_salud_temas') r=getPaginaSaludTemas();
  else if(action==='get_pagina_salud_accesos') r=getPaginaSaludAccesos(e.parameter.dni||'');
  // ── NUEVO: Kinesiología (portada QR, historia clínica kinésica) — con token ──
  else if(action==='get_kin_hc') r=getKinHC(e.parameter.dni||'');
  else if(action==='get_kin_ordenes') r=getKinOrdenes();
  else if(action==='get_pendientes_qx') r=getPendientesQx();
  else if(action==='get_panel_hc') r=getPanelHC();
  else r={error:'Acción no reconocida'};
  return jsonOut(r);
}
// ── Igual que ACCIONES_GET_PROTEGIDAS_, pero para las acciones de ESCRITURA
// (guardar/enviar/configurar) que usan las apps de especialistas — KSM,
// Colegas, Panel de Especialistas, Lector CUD y Agenda de Mensajes.
//
// CORREGIDO (revisión sep 2026): antes, las acciones que usa el propio Dr.
// Muratori desde OCM (addPaciente, addReceta, addHistoria,
// addHistoriaExtendida, actualizarConsulta, backup) quedaban A PROPÓSITO
// afuera de esta lista, bajo el supuesto de que "OCM tiene su propio control
// de acceso físico (es la compu del consultorio)". Ese supuesto era falso:
// la URL de este Web App está escrita en el código fuente de las ~29
// páginas públicas del sistema (cualquiera puede verla con "ver código
// fuente" del navegador), así que "hace falta estar en la compu del
// consultorio" nunca fue una barrera real — cualquiera con la URL podía
// crear pacientes falsos, pisar un paciente real adivinando su DNI, o
// cargar recetas/historias clínicas falsas, sin loguearse nunca. Ahora SÍ
// quedan protegidas por token, igual que el resto de la lista. OCM, KSM y
// Agenda de Mensajes ya mandan el token en estas llamadas cuando está
// configurado (variable APP_TOKEN de cada página) — mientras no corras
// configurarAppToken(), tokenValido_() sigue dejando pasar todo igual que
// antes, así que subir este archivo no corta nada por sí solo.
//
// Siguen A PROPÓSITO afuera de esta lista:
//   - saveConsulta — la llama un paciente ANÓNIMO desde
//     Recomendaciones_Medicas_Terapeuticas_v1.html para pedir una consulta
//     paga; nunca va a tener el token, igual que qr_solicitud_consulta.
//     Solo crea una consulta en estado "pendiente_pago", no expone ni pisa
//     historia clínica de nadie.
//   - las acciones que llenan los pacientes desde un link que reciben por
//     mail (saveAutoevaluacionRMYT, evaluacion_submit, kine_pregunta_submit,
//     ssq_guardar_seguimiento_posqx) — el paciente nunca tiene ni va a
//     tener el token, así que gatearlas rompería el sistema para ellos. Por
//   - qr_plantillas_solicitud y qr_plantillas_rx — igual que
//     qr_solicitud_consulta: las llena un paciente anónimo desde
//     QR_Plantillas.html. No leen nada de nadie: solo mandan un mail con las
//     fotos al médico y dan de alta/actualizan una fila 'solicitada' en
//     Plantillas_Sync (límite de pedidos por DNI y por hora).
//     la misma razón queda afuera qr_solicitud_consulta — la carga un
//     paciente anónimo que entró desde un código QR (guardia, comercios,
//     clubes), tampoco tiene token.
//   - por la misma razón quedan afuera reservar_turno y cancelar_turno —
//     los usa tanto el médico desde la Consulta Virtual (que sí tiene
//     token) como la secretaria de cada sede desde Turnos_Secretaria.html
//     (que NO tiene ni va a tener token: es una página de mostrador para
//     el QR impreso en cada consultorio). No exponen historia clínica,
//     solo nombre/DNI/teléfono/horario — el mismo nivel de dato que ya
//     entra sin token por qr_solicitud_consulta.
//   - agregar_horario_libre — mismo caso que reservar_turno/cancelar_turno,
//     pero para que la SECRETARIA DEL HOSPITAL SÍCOLI cargue ella misma los
//     horarios que van quedando libres (el hospital no tiene un horario
//     semanal fijo como los consultorios privados, así que no alcanza con
//     un patrón fijo en CONSULTORIOS_HORARIOS_ — ver
//     CONSULTORIOS_AGENDA_LIBRE_ más abajo). Tampoco expone historia
//     clínica, solo fecha/hora.
//   - registrar_aceptacion_ksm — es el primer registro de un especialista
//     nuevo, que todavía no tiene el token (se lo damos recién después de
//     que el médico lo autorice desde el Panel de Especialistas). El
//     registro en sí no expone datos de pacientes, solo crea un pedido de
//     alta pendiente.
//   - el webhook de Mercado Pago (data.type==='payment') — lo autentica
//     Mercado Pago con su propio formato, no nuestro token, y se resuelve
//     antes de llegar a este chequeo.
const ACCIONES_POST_PROTEGIDAS_ = [
  // AGREGADAS (revisión sep 2026) — ver el comentario completo arriba de
  // esta lista: eran las acciones de escritura más sensibles de todo el
  // sistema (alta/edición de pacientes, recetas, historia clínica, backup)
  // y estaban sin ninguna protección.
  'addPaciente','addReceta','addHistoria','addHistoriaExtendida','enviar_rimp_paciente',
  'guardar_programacion_qx','eliminar_programacion_qx','sync_seguimiento_lote','sync_plantillas_lote','guardar_pendiente_qx',
  'actualizarConsulta','backup',
  'generarLinkMeet','kine_evolucion','kine_informe_enviar','kine_pregunta_responder',
  // KEP (oct 2026): todas las usa el kinesiólogo o el médico. kep_submit queda PÚBLICA a propósito
  // (la llena el paciente desde el link del mail, como saveAutoevaluacionRMYT): exige que el DNI
  // exista en Pacientes, tiene tope diario por DNI y no devuelve ningún dato clínico.
  'kep_medicion','kep_enviar_link','kep_derivar_medico','kep_responder_medico',
  // Página de salud: crean/extienden códigos con la clave SIMM_CLAVE (solo la usa el médico)
  'ps_crear_acceso','ps_habilitar_clave','ps_extender_acceso','ps_renovar_paciente',
  // Kinesiología (el médico/kinesiólogo): notas de sesión, envío de la portada, alta con bitácora
  'kin_nota_sesion','kin_enviar_portada','alta_paciente','kin_membresia_estado','kin_orden_enviar','kin_nota_compartida',
  'guardar_neuroortopedia',
  // AGREGADA (revisión sep 2026): esta MISMA acción ya estaba en la lista de
  // GET protegidas de arriba (ACCIONES_GET_PROTEGIDAS_) — alguien la había
  // marcado como sensible ahí, pero se quedó afuera acá. Como doPost la
  // despacha igual (línea de abajo, "get_pacientes_incompletos"), cualquiera
  // podía pedir por POST, sin token, el nombre/DNI/teléfono/mail de TODOS
  // los pacientes con datos incompletos — exactamente el tipo de fuga que
  // esta lista existe para cortar.
  'get_pacientes_incompletos',
  'registrar_facturacion_sistema','enviar_resumen_facturacion_sistema',
  'set_peaje_sistema_pct','set_permisos_especialista',
  'analizar_foto_cud','guardar_registro_cud',
  'analizar_foto_documento_discap',
  'analizar_foto_edad_osea','analizar_foto_clasificacion_visual',
  'crear_link_pago_mp',
  'asistente_consulta_ia',
  'analizar_texto_paciente','guardar_paciente_historial',
  'confirmarPagoMensaje','marcarRespondido',
  // AGREGADAS (revisión sep 2026, sabotaje/adversarial): saveTarifas y
  // set_membresia las llama únicamente el médico desde OCM (configuración
  // de precios y el interruptor de membresía de un paciente) — no tienen
  // ningún llamador anónimo. Sin token, cualquiera podía cambiar los
  // precios que ve un paciente antes de pagar, o darle/quitarle membresía
  // (RMYT, mail con contexto clínico, agendas) a cualquier DNI sin que el
  // médico lo decidiera.
  'saveTarifas','set_membresia',
  // AGREGADAS (revisión sep 2026, sabotaje/adversarial): guardar_resumen_para_firma
  // y guardar_consentimiento_para_firma las llama el MÉDICO desde OCM (arman
  // el resumen/consentimiento y generan el link) — van con el token
  // compartido, como el resto de las acciones de OCM. NO se agrega acá
  // guardar_firma_paciente ni guardar_firma_consentimiento (las llama el
  // PACIENTE sin ese token): esas dos tienen su propio Firma_Token por
  // documento, chequeado adentro de cada función — agregarlas a esta lista
  // las rompería en cuanto se configure el token compartido.
  // enviar_resumen_firmado también es del médico (manda por mail el
  // resumen ya firmado a quien él decida) — antes cualquiera con la URL
  // podía pedir que se reenvíe el resumen firmado de cualquier DNI a
  // cualquier mail que eligiera, sin ser el médico.
  'guardar_resumen_para_firma','guardar_consentimiento_para_firma','enviar_resumen_firmado',
  // AGREGADAS (Membresía de Seguimiento, sep 2026): las tres las llama
  // únicamente el médico desde OCM — iniciar/renovar el mes (genera el
  // link de pago), desactivar, y revisar a mano si ya se pagó. Ninguna
  // la llama nunca el paciente directamente.
  'iniciar_renovar_membresia_seguimiento','desactivar_membresia_seguimiento','activar_cortesia_membresia',
  'confirmar_renovacion_membresia',
  // AGREGADAS (integración drmuratorisalud.ar, sept 2026): crean/modifican
  // el acceso del paciente a los esquemas protegidos del sitio y llaman a
  // un servicio externo con una clave secreta (SIMM_CLAVE) — nada de esto
  // lo tiene que poder disparar alguien que no sea el médico desde CAM V.
  'registrarAccesoSalud','habilitarClaveAccesoSalud','extenderAccesoSalud',
  'listarAccesosSalud','obtenerTemasAccesoSalud','alertasVencimientoAccesoSalud'
];

// ═══════════════════════════════════════════════════════════════════
// INTEGRACIÓN CON drmuratorisalud.ar — Clave A (tratamiento médico) y
// Clave B (tratamiento quirúrgico) de la página educativa (sept 2026,
// pedido de Luis). Basado en la especificación e implementación de
// referencia del sitio (simm/SIMM_Acceso_QR.gs) — mismo contrato de API,
// pero acá la SIMM_CLAVE queda SOLO en las Propiedades del Script (nunca
// en el navegador de un consultorio ni en un archivo compartido), y cada
// código queda anotado en la historia clínica de este mismo sistema, en
// la hoja "Accesos_Pagina_Salud" — el sitio en sí NUNCA recibe nombre,
// DNI, mail ni teléfono del paciente, solo tema/días/claves/código.
//
// Configuración (una sola vez, en el editor de Apps Script → Configuración
// del proyecto → Propiedades del script):
//   SITIO_SALUD_URL = https://drmuratorisalud.ar
//   SIMM_CLAVE      = (la misma clave cargada como secreto SIMM_CLAVE en
//                      Cloudflare, del lado del sitio — NO la clave del
//                      panel del sitio, que es otra distinta)
// ═══════════════════════════════════════════════════════════════════
function saludApi_(ruta, cuerpo){
  var props = PropertiesService.getScriptProperties();
  var base = (props.getProperty('SITIO_SALUD_URL') || 'https://drmuratorisalud.ar').replace(/\/$/, '');
  var clave = props.getProperty('SIMM_CLAVE');
  if(!clave) throw new Error('Falta configurar SIMM_CLAVE en Propiedades del Script (Configuración del proyecto).');
  var opciones = {
    method: cuerpo ? 'post' : 'get',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + clave },
    muteHttpExceptions: true
  };
  if(cuerpo) opciones.payload = JSON.stringify(cuerpo);
  var r = UrlFetchApp.fetch(base + '/api/integracion/' + ruta, opciones);
  var texto = r.getContentText() || '{}';
  var datos;
  try{ datos = JSON.parse(texto); }catch(e){ throw new Error('El sitio respondió algo que no es JSON válido — revisá que el dominio esté bien apuntado.'); }
  if(r.getResponseCode() >= 400) throw new Error(datos.error || ('Error ' + r.getResponseCode() + ' del sitio.'));
  return datos;
}

function hojaAccesosSalud_(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Accesos_Pagina_Salud');
  if(!sheet){
    sheet = ss.insertSheet('Accesos_Pagina_Salud');
    sheet.appendRow(['ID','DNI','Nombre_Paciente','Codigo','Tema','Claves','Vence','Plan_Membresia','Enlace','Creado','Historial_JSON']);
  }
  return sheet;
}

// Busca la fila (1-indexada, con encabezado) de un código dado. -1 si no está.
function buscarFilaAccesoSalud_(sheet, codigo){
  var datos = sheet.getDataRange().getValues();
  for(var i=1;i<datos.length;i++){
    if(String(datos[i][3]||'').toUpperCase() === String(codigo||'').toUpperCase()) return i+1;
  }
  return -1;
}

function agregarHistorialAccesoSalud_(historialJson, accion, previo, nuevo){
  var lista = [];
  try{ lista = JSON.parse(historialJson || '[]'); }catch(e){ lista = []; }
  lista.push({fecha: new Date().toISOString(), accion: accion, previo: previo, nuevo: nuevo});
  return JSON.stringify(lista);
}

/** Lista de temas con esquema protegido: {success, temas:[{slug,titulo,claves}]} */
function obtenerTemasAccesoSalud(){
  try{
    return {success:true, temas: saludApi_('temas').temas};
  }catch(e){
    return {success:false, error:String(e.message||e)};
  }
}

/**
 * Genera un nuevo acceso para un paciente y lo registra en su historia
 * clínica. d = {dni, nombre, tema, dias, claves ('A'|'B'|'AB'), plan_membresia}
 */
function registrarAccesoSalud(d){
  if(!d || !String(d.dni||'').trim()) return {success:false, error:'Falta el DNI del paciente'};
  if(!String(d.tema||'').trim()) return {success:false, error:'Falta elegir un tema'};
  try{
    var r = saludApi_('codigo', {tema:d.tema, dias:d.dias||30, claves:d.claves||'A'});
    var sheet = hojaAccesosSalud_();
    var id = Utilities.getUuid();
    var historial = agregarHistorialAccesoSalud_('[]', 'crear', null, {claves:r.claves, vence:r.vence});
    sheet.appendRow([id, d.dni, d.nombre||'', r.codigo, r.tema, r.claves, r.vence, d.plan_membresia||'', r.enlace, new Date().toISOString(), historial]);
    return {success:true, id:id, codigo:r.codigo, tema:r.tema, claves:r.claves, vence:r.vence, enlace:r.enlace};
  }catch(e){
    return {success:false, error:String(e.message||e)};
  }
}

/**
 * Habilita una clave adicional (ej. 'B' al decidirse la cirugía) en un
 * código YA existente de ese paciente — no genera un QR nuevo.
 * d = {codigo, clave}
 */
function habilitarClaveAccesoSalud(d){
  if(!d || !String(d.codigo||'').trim()) return {success:false, error:'Falta el código'};
  try{
    var r = saludApi_('habilitar', {codigo:d.codigo, clave:d.clave||'B'});
    var sheet = hojaAccesosSalud_();
    var fila = buscarFilaAccesoSalud_(sheet, d.codigo);
    if(fila>0){
      var previo = sheet.getRange(fila,6).getValue();
      var historialActual = sheet.getRange(fila,11).getValue();
      sheet.getRange(fila,6).setValue(r.claves);
      sheet.getRange(fila,7).setValue(r.vence);
      sheet.getRange(fila,11).setValue(agregarHistorialAccesoSalud_(historialActual, 'habilitar', previo, r.claves));
    }
    return {success:true, codigo:r.codigo, tema:r.tema, claves:r.claves, vence:r.vence, enlace:r.enlace};
  }catch(e){
    return {success:false, error:String(e.message||e)};
  }
}

/** Extiende la vigencia de un código (membresía). d = {codigo, dias, plan_membresia} */
function extenderAccesoSalud(d){
  if(!d || !String(d.codigo||'').trim()) return {success:false, error:'Falta el código'};
  try{
    var r = saludApi_('extender', {codigo:d.codigo, dias:d.dias||30});
    var sheet = hojaAccesosSalud_();
    var fila = buscarFilaAccesoSalud_(sheet, d.codigo);
    if(fila>0){
      var previoVence = sheet.getRange(fila,7).getValue();
      var historialActual = sheet.getRange(fila,11).getValue();
      sheet.getRange(fila,7).setValue(r.vence);
      if(d.plan_membresia) sheet.getRange(fila,8).setValue(d.plan_membresia);
      sheet.getRange(fila,11).setValue(agregarHistorialAccesoSalud_(historialActual, 'extender', previoVence, r.vence));
    }
    return {success:true, codigo:r.codigo, tema:r.tema, claves:r.claves, vence:r.vence, enlace:r.enlace};
  }catch(e){
    return {success:false, error:String(e.message||e)};
  }
}

/** Lista los accesos ya generados para un DNI (para mostrarlos en CAM V). */
function listarAccesosSalud(d){
  var dni = (d && d.dni) || '';
  var sheet = hojaAccesosSalud_();
  var datos = sheet.getDataRange().getValues();
  var out = [];
  for(var i=1;i<datos.length;i++){
    if(String(datos[i][1]||'') === String(dni)){
      out.push({id:datos[i][0], codigo:datos[i][3], tema:datos[i][4], claves:datos[i][5], vence:datos[i][6], plan_membresia:datos[i][7], enlace:datos[i][8]});
    }
  }
  return {success:true, accesos: out};
}

/** Accesos que vencen entre hoy y los próximos `dias` días (por defecto 7) — para ofrecer la extensión en el próximo control. */
function alertasVencimientoAccesoSalud(d){
  var diasAlerta = (d && d.dias) || 7;
  var hoyD = new Date();
  var limite = new Date(); limite.setDate(limite.getDate() + diasAlerta);
  var sheet = hojaAccesosSalud_();
  var datos = sheet.getDataRange().getValues();
  var out = [];
  for(var i=1;i<datos.length;i++){
    var venceStr = datos[i][6];
    if(!venceStr) continue;
    var vence = new Date(venceStr);
    if(vence >= hoyD && vence <= limite){
      out.push({dni:datos[i][1], nombre:datos[i][2], codigo:datos[i][3], tema:datos[i][4], claves:datos[i][5], vence:venceStr});
    }
  }
  return {success:true, alertas: out};
}

function doPost(e) {
  const data=JSON.parse(e.postData.contents);

  // ── Notificación de Mercado Pago: no manda nuestro campo 'action', manda 'type' + 'data.id' ──
  if(data.type==='payment' && data.data && data.data.id){
    return jsonOut(confirmarPagoMP(String(data.data.id)));
  }

  const action = data.action || data.accion || '';
  const trial = estadoTrial_();
  if(trial.vencido){
    return jsonOut({success:false, error:'trial_vencido', vencido:true, mensaje:'El período de prueba finalizó el '+trial.fecha_vencimiento+'.'});
  }
  if(ACCIONES_POST_PROTEGIDAS_.indexOf(action)>-1 && !tokenValido_(data.token)){
    return jsonOut({success:false, error:'Token inválido o faltante — pedile al médico el token de acceso'});
  }
  let r;
  if(action==='addPaciente') r=addPaciente(data);
  else if(action==='set_membresia') r=setMembresiaPaciente(data);
  else if(action==='get_pacientes_incompletos') r=getPacientesCamposIncompletos();
  else if(action==='enviar_recordatorios_campos') r=enviarRecordatorioCamposIncompletos(data);
  else if(action==='addReceta') r=addReceta(data);
  else if(action==='addHistoria') r=addHistoria(data);
  else if(action==='addHistoriaExtendida') r=addHistoriaExtendida(data);
  else if(action==='enviarRecetaPorMail') r=enviarRecetaPorMail(data);
  else if(action==='enviar_rimp_paciente') r=enviarRimpPaciente(data);
  else if(action==='guardar_registro_cud') r=guardarRegistroCUD(data.registro||data);
  else if(action==='analizar_foto_cud') r=analizarFotoCUD(data);
  else if(action==='analizar_foto_documento_discap') r=analizarFotoDocumentoDiscap(data);
  else if(action==='analizar_foto_edad_osea') r=analizarFotoEdadOsea(data);
  else if(action==='analizar_foto_clasificacion_visual') r=analizarFotoClasificacionVisual(data);
  else if(action==='crear_link_pago_mp') r=crearLinkPagoMP(data);
  else if(action==='asistente_consulta_ia') r=asistenteConsultaIA(data);
  else if(action==='analizar_texto_paciente') r=analizarTextoPaciente(data);
  else if(action==='analizar_foto_historia') r=analizarFotoHistoria(data);
  else if(action==='guardar_paciente_historial') r=guardarPacienteConHistorial(data);
  // ── NUEVO ──
  else if(action==='backup') r=guardarBackup(data);
  else if(action==='ridm_submit') r=guardarRidmSubmit(data);
  else if(action==='ridm_antecedentes') r=guardarRidmAntecedentes(data);
  else if(action==='guardar_resumen_para_firma') r=guardarResumenParaFirma(data);
  else if(action==='guardar_firma_paciente') r=guardarFirmaPaciente(data);
  else if(action==='enviar_resumen_firmado') r=enviarResumenFirmadoPorMail(data);
  else if(action==='guardar_consentimiento_para_firma') r=guardarConsentimientoParaFirma(data);
  else if(action==='guardar_firma_consentimiento') r=guardarFirmaConsentimiento(data);
  else if(action==='qr_solicitud_consulta') r=registrarSolicitudQR(data);
  else if(action==='qr_plantillas_solicitud') r=registrarPlantillasQR(data);
  else if(action==='qr_plantillas_rx') r=registrarInformeRxQR(data);
  else if(action==='kine_evolucion') r=guardarKineEvolucion(data);
  else if(action==='guardar_neuroortopedia') r=guardarNeuroortopedia(data);
  else if(action==='evaluacion_submit') r=guardarEvaluacionPaciente(data);
  else if(action==='marcar_pago_atendido') r=marcarPagoAtendido(data.payment_id||'');
  else if(action==='saveRecomendaciones') r=saveRecomendaciones(data);
  else if(action==='saveAutoevaluacionRMYT') r=saveAutoevaluacionRMYT(data);
  // ── NUEVO: KEP — kep_submit lo llena el PACIENTE desde su link (sin token, igual que RMYT) ──
  else if(action==='kep_submit') r=guardarKepEvaluacion(data);
  else if(action==='kep_medicion') r=guardarKepMedicion(data);
  else if(action==='kep_enviar_link') r=kepEnviarLink(data);
  else if(action==='kep_derivar_medico') r=kepDerivarMedico(data);
  else if(action==='kep_responder_medico') r=kepResponderMedico(data);
  else if(action==='ps_crear_acceso') r=crearAccesoPaginaSalud(data);
  else if(action==='ps_habilitar_clave') r=habilitarClavePaginaSalud(data);
  else if(action==='ps_extender_acceso') r=extenderAccesoPaginaSalud(data);
  else if(action==='ps_renovar_paciente') r=renovarAccesosPaciente(data);
  // Kinesiología: estas 4 las usa el PACIENTE desde el QR del centro (sin token, como RMYT). Se identifica con DNI + fecha de
  // nacimiento, con tope de intentos; no devuelven datos clínicos. El resto exige token.
  else if(action==='kin_identificar') r=kinIdentificar(data);
  else if(action==='kin_inscribir') r=kinInscribir(data);
  else if(action==='kin_entrada') r=kinEntrada(data);
  else if(action==='kin_salida') r=kinSalida(data);
  else if(action==='kin_mensaje') r=kinMensaje(data);
  else if(action==='kin_membresia_info') r=kinMembresiaInfo();
  else if(action==='kin_membresia_solicitar') r=kinMembresiaSolicitar(data);
  else if(action==='kin_membresia_estado') r=kinMembresiaEstado(data);
  else if(action==='kin_orden_enviar') r=kinOrdenEnviar(data);
  else if(action==='kin_nota_compartida') r=kinNotaCompartida(data);
  else if(action==='kin_nota_sesion') r=kinNotaSesion(data);
  else if(action==='kin_enviar_portada') r=kinEnviarPortada(data);
  else if(action==='alta_paciente') r=altaPaciente(data);
  else if(action==='ssq_guardar_seguimiento_posqx') r=guardarSeguimientoPosqx(data);
  // ── NUEVO: GONIO — reconecta Costa-Bártani/Valgo-Varo/Ángulos de Frente/Podoscopia con su guardado ──
  else if(action==='guardar_gonio_medicion') r=guardarGonioMedicion(data);
  else if(action==='saveConsulta') r=saveConsulta(data);
  else if(action==='actualizarConsulta') r=actualizarConsulta(data);
  else if(action==='generarRespuestaIA') r=generarRespuestaIA(data);
  else if(action==='saveTarifas') r=saveTarifas(data);
  else if(action==='marcarRespondido') r=marcarRespondido(data);
  else if(action==='generarLinkMeet') r=generarLinkMeet(data);
  else if(action==='confirmarPagoMensaje') r=confirmarPagoMensaje(data);
  else if(action==='reservar_turno') r=reservarTurno(data);
  else if(action==='guardar_programacion_qx') r=guardarProgramacionQx(data);
  else if(action==='eliminar_programacion_qx') r=eliminarProgramacionQx(data);
  else if(action==='cancelar_turno') r=cancelarTurno(data);
  else if(action==='get_turno_publico') r=getTurnoPublico(data);
  else if(action==='subir_comprobante_sena_turno') r=subirComprobanteSenaTurno(data);
  else if(action==='verificar_sena_turno') r=verificarSenaTurno(data);
  else if(action==='get_senas_pendientes_verificar') r=getSenasPendientesVerificar();
  else if(action==='get_proximos_turnos') r=getProximosTurnos();
  else if(action==='agregar_horario_libre') r=agregarHorarioLibre(data);
  else if(action==='kine_pregunta_submit') r=guardarPreguntaKine(data);
  else if(action==='kine_pregunta_responder') r=responderPreguntaKine(data);
  else if(action==='kine_informe_enviar') r=enviarInformeKine(data);
  else if(action==='set_peaje_sistema_pct') r=setPeajeSistemaPct(data);
  else if(action==='registrar_aceptacion_ksm') r=registrarAceptacionKSM(data);
  else if(action==='registrar_facturacion_sistema') r=registrarFacturacionSistema(data);
  else if(action==='enviar_resumen_facturacion_sistema') r=enviarResumenFacturacionSistema(data);
  else if(action==='guardar_item_quirurgico') r=guardarItemQuirurgico(data);
  // ── NUEVO: permisos de especialistas + sincronización de módulos antes locales ──
  else if(action==='set_permisos_especialista') r=setPermisosEspecialista(data);
  else if(action==='sync_empresa_catastro') r=sincronizarEmpresaCatastro(data);
  else if(action==='guardar_cud_libro_discapacidad') r=guardarLibroDiscapacidadCUD(data);
  else if(action==='sync_plantilla') r=sincronizarPlantilla(data);
  else if(action==='sync_plantillas_lote') r=sincronizarPlantillasLote(data);
  else if(action==='guardar_pendiente_qx') r=guardarPendienteQx(data);
  else if(action==='sync_seguimiento_lote') r=sincronizarSeguimientoLote(data);
  // ── NUEVO: Mini-RIMP — igual que saveAutoevaluacionRMYT/evaluacion_submit,
  // sin token: lo completa un especialista externo que llega por un link
  // (QR de la receta de interconsulta), nunca va a tener el token del sistema ──
  else if(action==='guardar_minirimp') r=guardarRespuestaMiniRIMP(data);
  // ── NUEVO: Membresía de Seguimiento (mensual, por patología puntual) ──
  else if(action==='iniciar_renovar_membresia_seguimiento') r=iniciarORenovarMembresiaSeguimiento(data);
  else if(action==='desactivar_membresia_seguimiento') r=desactivarMembresiaSeguimiento(data);
  else if(action==='activar_cortesia_membresia') r=activarCortesiaMembresia(data);
  else if(action==='confirmar_renovacion_membresia') r=confirmarRenovacionMembresiaSeguimiento(data.DNI||data.dni||'');
  // ── NUEVO: integración con drmuratorisalud.ar (Clave A/B) ──
  else if(action==='registrarAccesoSalud') r=registrarAccesoSalud(data);
  else if(action==='habilitarClaveAccesoSalud') r=habilitarClaveAccesoSalud(data);
  else if(action==='extenderAccesoSalud') r=extenderAccesoSalud(data);
  else if(action==='listarAccesosSalud') r=listarAccesosSalud(data);
  else if(action==='obtenerTemasAccesoSalud') r=obtenerTemasAccesoSalud();
  else if(action==='alertasVencimientoAccesoSalud') r=alertasVencimientoAccesoSalud(data);
  else r={error:'Acción no reconocida'};
  return jsonOut(r);
}
function jsonOut(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}

// ── Funciones que ya existían, sin tocar ──
function getAllRows(sheetName){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName(sheetName);
  if(!sheet) return{error:'Hoja no encontrada: '+sheetName};
  const data=sheet.getDataRange().getValues();
  if(data.length<2) return{success:true,data:[]};
  const headers=data[0];
  return{success:true,data:data.slice(1).map(row=>{const obj={};headers.forEach((h,i)=>{obj[h]=row[i];});return obj;})};
}
function searchPatient(query){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName('Pacientes');
  if(!sheet) return{error:'Hoja Pacientes no encontrada'};
  const data=sheet.getDataRange().getValues(),headers=data[0],q=query.toLowerCase();
  const res=data.slice(1).filter(row=>row.some(c=>String(c).toLowerCase().includes(q))).map(row=>{const obj={};headers.forEach((h,i)=>{obj[h]=row[i];});return obj;});
  // También los pacientes que se inscribieron solo en kinesiología (QR del centro)
  try{
    const kh=ss.getSheetByName('KIN_HC');
    if(kh){
      const kd=kh.getDataRange().getValues(),kk=kd[0]||[],ci=n=>kk.indexOf(n),vistos={};
      res.forEach(o=>{vistos[String(o.DNI)]=1;});
      for(let j=kd.length-1;j>=1;j--){
        const dni=String(kd[j][ci('DNI')]);
        if(vistos[dni]) continue;
        if(['DNI','Apellido','Nombre'].some(n=>String(kd[j][ci(n)]).toLowerCase().includes(q))){
          vistos[dni]=1;
          res.push({ID:'',DNI:dni,Apellido:kd[j][ci('Apellido')],Nombre:kd[j][ci('Nombre')],Email:kd[j][ci('Email')],Telefono:kd[j][ci('Telefono')],Obra_Social:kd[j][ci('Obra_Social')],Diagnostico_Principal:String(kd[j][ci('Diagnostico')]||'')+' ('+String(kd[j][ci('Region')]||'')+')',_origen:'kinesiologia'});
        }
      }
    }
  }catch(e){}
  return{success:true,data:res};
}
// Agrega las columnas nuevas a la hoja Pacientes si todavía no existen
// (para sistemas que ya estaban en uso antes de agosto 2026)
function asegurarColumnasPacientes(sheet){
  let headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const faltantes = ['Localidad','Trabajos_JSON','Vivienda_JSON','Centro_Derivador','RIMP_Enviado','Antecedentes_Generales_JSON','Membresia'].filter(h=>headers.indexOf(h)===-1);
  faltantes.forEach(h=>{
    sheet.getRange(1, sheet.getLastColumn()+1).setValue(h).setBackground('#1A73E8').setFontColor('white').setFontWeight('bold');
  });
  return sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
}

// ── RIMP: link y envío del formulario de antecedentes (modo=clinico) ──
// Centralizado acá porque lo usan dos caminos distintos: el envío INMEDIATO
// al guardar un paciente marcado como "primera consulta" (addPaciente) y el
// recordatorio semanal a quienes todavía tienen campos vacíos (red de
// seguridad, por si el envío inmediato no se disparó o falló).
function getRimpUrl_(){
  // BUGFIX: tenía "/02_CELULAR/" de cuando se organizaban copias locales por
  // carpeta — pero el sitio publicado en GitHub Pages es plano (sin subcarpetas),
  // así que ese link SIEMPRE daba 404. Cualquier mail mandado antes de este
  // arreglo (RIMP inmediato al guardar "primera consulta", recordatorio semanal)
  // le llegó al paciente con un link roto.
  return baseSimm_() + 'RIDM_Muratori_v1.html'; // Subdominio de S.I.M-M (Cloudflare Pages)
}
// p.extenso=true → RIMP completo (infancia, vacunas, COVID, vivienda, jornada
// laboral) — se usa para trámites de discapacidad/CUD. Sin ese flag (default)
// se manda la versión breve (alergias, enfermedades, cirugías, medicación,
// deporte/federación, trabajo), pensada para una consulta habitual.
function construirLinkRimpAntecedentes_(p){
  // p.modo, si viene, manda por sobre p.extenso — lo usa
  // getPacientesCamposIncompletos() para elegir el modo exacto según qué le
  // falta a cada paciente (ver comentario ahí abajo).
  const modo = p.modo || (p.extenso ? 'clinico' : 'clinico_breve');
  return `${getRimpUrl_()}?modo=${modo}&apellido=${encodeURIComponent(p.apellido||'')}&nombre=${encodeURIComponent(p.nombre||'')}&dni=${encodeURIComponent(p.dni||'')}&tel=${encodeURIComponent(p.telefono||'')}`;
}
function enviarLinkRimpAPaciente_(p){
  if(!p || !p.email) throw new Error('Falta el mail del paciente');
  const link = construirLinkRimpAntecedentes_(p);
  const cuerpo = p.extenso ?
    (`Hola ${p.nombre||''} ${p.apellido||''},\n\n` +
    `Para completar tu ficha médica te pedimos que cargues algunos datos adicionales (antecedentes, datos laborales y de vivienda).\n\n` +
    `Por favor, entrá al siguiente link cuando puedas:\n${link}\n\n` +
    `Es rápido y nos ayuda a llevar un mejor seguimiento de tu salud.\n\n` +
    `Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943`)
    :
    (`Hola ${p.nombre||''} ${p.apellido||''},\n\n` +
    `Antes de tu consulta te pedimos que completes algunos datos (alergias, enfermedades, cirugías, medicación habitual y actividad física).\n\n` +
    `Por favor, entrá al siguiente link cuando puedas:\n${link}\n\n` +
    `Es rápido y nos ayuda a prepararnos mejor para tu consulta.\n\n` +
    `Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943`);
  MailApp.sendEmail(p.email, 'Completá tu ficha médica — RIMP', cuerpo, {
    replyTo:'drluismuratori@gmail.com', name:'Dr. Luis Alberto Muratori'
  });
}
// ── Botón "📨 Enviar RIMP" de OCM: manda el mail DIRECTO desde el servidor (sin abrir Gmail).
// El link ya viene armado por OCM con los datos de la ficha precargados (para que el
// paciente solo revise y complete). Se valida que apunte a nuestro RIDM y que el mail sea válido.
function enviarRimpPaciente(d){
  const mail = String(d.email||'').trim().toLowerCase();
  if(!/^[^@\s]+@[^@\s.]+(\.[^@\s.]+)+$/.test(mail)) return {success:false, error:'El mail no es válido: '+mail};
  const link = String(d.link||'');
  if(link.indexOf(getRimpUrl_()+'?')!==0) return {success:false, error:'Link de RIMP inválido'};
  const intro = d.extenso
    ? 'Para completar tu ficha médica te pedimos que cargues algunos datos adicionales (antecedentes, datos laborales y de vivienda).'
    : 'Antes de tu consulta te pedimos que completes algunos datos (alergias, enfermedades, cirugías, medicación habitual y actividad física).';
  const cuerpo = 'Hola ' + (d.nombre||'') + ' ' + (d.apellido||'') + ',\n\n' + intro + '\n\n' +
    'Por favor, entrá al siguiente link cuando puedas:\n' + link + '\n\n' +
    'La planilla ya te aparece con los datos que tenemos de tu ficha: revisalos, corregí lo que haga falta, completá lo que falte y apretá Enviar al final. Es rápido y nos ayuda a prepararnos mejor.\n\n' +
    'Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943';
  MailApp.sendEmail(mail, 'Completá tu ficha médica — RIMP', cuerpo, {replyTo:'drluismuratori@gmail.com', name:'Dr. Luis Alberto Muratori'});
  try{ // deja registrado en la hoja Pacientes que ya se le mandó
    const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pacientes');
    if(sheet && d.dni){
      const headers=asegurarColumnasPacientes(sheet), cDni=headers.indexOf('DNI'), cRimp=headers.indexOf('RIMP_Enviado');
      // Solo se lee la columna DNI (antes se leía TODA la hoja, con las columnas JSON, y tardaba tanto que
      // OCM llegaba a cortar la espera y no mostraba el cartel de "enviado" aunque el mail ya había salido).
      const nFilas=sheet.getLastRow()-1;
      if(nFilas>0){
        const dnis=sheet.getRange(2,cDni+1,nFilas,1).getValues();
        for(let i=0;i<dnis.length;i++){ if(String(dnis[i][0])===String(d.dni)){ sheet.getRange(i+2,cRimp+1).setValue(new Date().toLocaleString('es-AR')); break; } }
      }
    }
  }catch(_){}
  return {success:true, message:'RIMP enviado a '+mail};
}
// Marca en la hoja Pacientes que a este paciente (por ID) ya se le mandó el
// RIMP, para no volver a mandárselo (ni desde el guardado inmediato ni desde
// el recordatorio semanal) hasta que alguien lo resetee a mano si hiciera falta.
function marcarRimpEnviado_(idPaciente){
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sheet=ss.getSheetByName('Pacientes');
  if(!sheet) return;
  const headers = asegurarColumnasPacientes(sheet);
  const colId = headers.indexOf('ID'), colRimp = headers.indexOf('RIMP_Enviado');
  if(colId===-1 || colRimp===-1) return;
  const data = sheet.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][colId])===String(idPaciente)){
      sheet.getRange(i+1, colRimp+1).setValue(new Date().toLocaleString('es-AR'));
      return;
    }
  }
}

// ── Membresía: la da el médico, sin costo, en cualquier momento (no solo al
// crear la filiación) — por eso vive como acción propia además del campo que
// ya maneja addPaciente. La usa el botón "🤝 Dar membresía" de OCM y también
// el alta automática al primer control post-quirúrgico (ver
// guardarSeguimientoPosqx más abajo).
function setMembresiaPaciente(d){
  const dni = String(d.DNI || d.dni || '');
  if(!dni) return {success:false, error:'Falta DNI del paciente'};
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sheet=ss.getSheetByName('Pacientes');
  if(!sheet) return {success:false, error:'No existe la hoja Pacientes'};
  const headers = asegurarColumnasPacientes(sheet);
  const colDNI = headers.indexOf('DNI'), colMembresia = headers.indexOf('Membresia');
  if(colDNI===-1 || colMembresia===-1) return {success:false, error:'No se pudo preparar la columna Membresia'};
  const data = sheet.getDataRange().getValues();
  const membresia = (d.Membresia===true || d.Membresia==='true' || d.Membresia==='SI') ? 'SI' : 'NO';
  for(let i=1;i<data.length;i++){
    if(String(data[i][colDNI])===dni){
      sheet.getRange(i+1, colMembresia+1).setValue(membresia);
      return {success:true, message:'Membresía actualizada', membresia};
    }
  }
  return {success:false, error:'No se encontró un paciente con ese DNI'};
}

// Helper reutilizable: ¿este DNI tiene membresía activa? Lo usan las
// funciones que dan acceso a los 4 beneficios que solo son para paciente
// con membresía (RMYT, RSP con contexto clínico de mail, Agendas, RSPQ) —
// ver el comentario largo en addPaciente/Membresia más arriba. Devuelve
// false ante cualquier duda (sin hoja, sin columna, sin fila para ese DNI)
// para no otorgar por error algo que es decisión explícita del médico.
function tienePacienteMembresia_(dni){
  if(!dni) return false;
  // Membresía paga vigente (4 semanas) — la única que se otorga desde oct 2026
  if(membresiaSeguimientoVigente_(dni)) return true;
  // Membresías gratuitas anteriores (campo Membresia=SI de Pacientes): se siguen respetando
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sheet = ss.getSheetByName('Pacientes');
  if(!sheet) return false;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colDNI = headers.indexOf('DNI'), colMembresia = headers.indexOf('Membresia');
  if(colDNI===-1 || colMembresia===-1) return false;
  const dniStr = String(dni);
  for(let i=1;i<data.length;i++){
    if(String(data[i][colDNI])===dniStr){
      return String(data[i][colMembresia]||'').trim().toUpperCase()==='SI';
    }
  }
  return false;
}

function addPaciente(d){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName('Pacientes');
  const headers = asegurarColumnasPacientes(sheet);
  const data=sheet.getDataRange().getValues();
  const dniBuscado=String(d.DNI||'');
  const colDNI = headers.indexOf('DNI');
  const colRimpEnviado = headers.indexOf('RIMP_Enviado');

  const valoresNuevos = {
    Apellido:d.Apellido, Nombre:d.Nombre, DNI:dniBuscado, Fecha_Nac:d.Fecha_Nac,
    Obra_Social:d.Obra_Social, N_Afiliado:d.N_Afiliado, Plan:d.Plan, Telefono:d.Telefono,
    Email:d.Email, Domicilio:d.Domicilio, Lugar_Atencion:d.Lugar_Atencion,
    Diagnostico_Principal:d.Diagnostico_Principal, Estado_Tratamiento:d.Estado_Tratamiento||'Activo',
    Localidad:d.Localidad, Trabajos_JSON: d.Trabajos ? JSON.stringify(d.Trabajos) : undefined,
    Vivienda_JSON: d.Vivienda ? JSON.stringify(d.Vivienda) : undefined,
    Centro_Derivador: d.Centro_Derivador,
    // Antecedentes que carga el médico a mano la primera vez (motivo inicial,
    // enfermedades, cirugías, medicación, actividad laboral y deportiva) —
    // quedan en el paciente (no solo en la consulta del día) para que estén
    // disponibles siempre: en CAM, en certificados, en recetas, etc.
    Antecedentes_Generales_JSON: d.Antecedentes_Generales ? JSON.stringify(d.Antecedentes_Generales) : undefined,
    // Membresía: elección del Dr. Muratori, sin costo — la filiación (este
    // mismo alta) NO la da por sí sola, solo habilita lo transaccional
    // (recetas, pedidos de estudio, certificados). Membresía es lo que
    // habilita RMYT, respuesta de mail con contexto clínico (RSP), agendas
    // y seguimiento posquirúrgico (RSPQ) — se otorga acá si vino explícita
    // en el alta, o después con la acción 'set_membresia' / el botón
    // "🤝 Dar membresía" en OCM, o automático al primer control
    // posquirúrgico (ver guardarSeguimientoPosqx).
    Membresia: d.Membresia !== undefined ? (d.Membresia ? 'SI' : 'NO') : undefined
  };

  let idFinal = '', actualizado = false, yaTeniaRimp = false;

  if(dniBuscado){
    for(let i=1;i<data.length;i++){
      if(String(data[i][colDNI])===dniBuscado){
        const fila = data[i].slice();
        headers.forEach((h,idx)=>{
          if(valoresNuevos[h] !== undefined && valoresNuevos[h] !== ''){ fila[idx] = valoresNuevos[h]; }
        });
        sheet.getRange(i+1,1,1,fila.length).setValues([fila]);
        idFinal = data[i][headers.indexOf('ID')]; actualizado = true;
        yaTeniaRimp = colRimpEnviado>-1 ? !!data[i][colRimpEnviado] : false;
        break;
      }
    }
  }
  if(!actualizado){
    idFinal='PAC-'+String(sheet.getLastRow()).padStart(4,'0');
    const filaNueva = headers.map(h=> h==='ID' ? idFinal : (valoresNuevos[h] || ''));
    sheet.appendRow(filaNueva);
  }

  // ── "Primera consulta": si se marcó explícitamente y el paciente tiene
  // mail y todavía no se le había mandado el RIMP, se lo mandamos ahora
  // mismo — no hay que esperar al recordatorio semanal. Si ya se le había
  // mandado (a este mismo paciente, en un guardado anterior), no se repite,
  // así no se lo satura con el mismo link cada vez que se edita su ficha. ──
  let rimpEnviado = false;
  const mailDestino = d.Email || valoresNuevos.Email;
  if(d.Primera_Consulta && mailDestino && !yaTeniaRimp){
    try{
      enviarLinkRimpAPaciente_({
        apellido: d.Apellido||valoresNuevos.Apellido, nombre: d.Nombre||valoresNuevos.Nombre,
        dni: dniBuscado, telefono: d.Telefono||valoresNuevos.Telefono, email: mailDestino,
        // RIMP completo (vivienda, trabajos, infancia, vacunas, COVID) solo para
        // trámites de discapacidad/CUD — el resto recibe la versión breve.
        extenso: !!d.Discapacidad
      });
      rimpEnviado = true;
      marcarRimpEnviado_(idFinal);
    }catch(e){ /* si falla el envío del mail, no se interrumpe el guardado del paciente */ }
  }

  return{success:true, id:idFinal, actualizado:actualizado, rimpEnviado:rimpEnviado};
}

// ══════════════════════════════════════════════════════════════════
// FASE A (parte 2) — VISTA CONSOLIDADA POR DNI
// Cruza Pacientes + Historia_Clinica + Discapacidad_CUD + antecedentes
// de RIMP (RIDM_Pendientes) en un solo objeto por paciente. Es la base
// que va a usar el futuro Dashboard de estadísticas (Fase B).
// ══════════════════════════════════════════════════════════════════
function sheetToObjects(nombreHoja){
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sheet=ss.getSheetByName(nombreHoja);
  if(!sheet) return [];
  const data=sheet.getDataRange().getValues();
  if(data.length<2) return [];
  const headers=data[0];
  return data.slice(1).map(row=>{
    const obj={}; headers.forEach((h,i)=>{ obj[h]=row[i]; }); return obj;
  });
}

function construirVistaConsolidada(dni, cache){
  cache = cache || {
    pacientes: sheetToObjects('Pacientes'),
    historia: sheetToObjects('Historia_Clinica'),
    cud: sheetToObjects('Discapacidad_CUD'),
    ridm: sheetToObjects('RIDM_Pendientes')
  };
  const dniStr = String(dni||'');

  const pac = cache.pacientes.find(p=>String(p.DNI)===dniStr) || {};
  const historiaPac = cache.historia.filter(h=>String(h.DNI)===dniStr);
  const cudPac = cache.cud.filter(c=>String(c.DNI)===dniStr);
  // Último envío de RIMP para este DNI (donde viven los antecedentes detallados)
  const ridmPac = cache.ridm.filter(r=>String(r.DNI)===dniStr).sort((a,b)=>new Date(b.Timestamp)-new Date(a.Timestamp))[0];
  let antecedentes = {};
  try{ antecedentes = ridmPac ? JSON.parse(ridmPac.Datos_JSON||'{}') : {}; }catch(_){ antecedentes = {}; }

  let trabajos = [];
  try{ trabajos = pac.Trabajos_JSON ? JSON.parse(pac.Trabajos_JSON) : (antecedentes.trabajos||[]); }catch(_){ trabajos = antecedentes.trabajos||[]; }
  let vivienda = {};
  try{ vivienda = pac.Vivienda_JSON ? JSON.parse(pac.Vivienda_JSON) : {}; }catch(_){ vivienda = {}; }
  // Antecedentes que el médico cargó a mano en "+ Nuevo paciente" (motivo inicial,
  // enfermedades, cirugías, medicación, actividad laboral y deportiva) — el "pool"
  // que se guarda una sola vez y de acá en más pueden reusar CAM, certificados, etc.
  let antecedentesGenerales = {};
  try{ antecedentesGenerales = pac.Antecedentes_Generales_JSON ? JSON.parse(pac.Antecedentes_Generales_JSON) : {}; }catch(_){ antecedentesGenerales = {}; }

  return {
    dni: dniStr,
    apellido: pac.Apellido||antecedentes.apellido||'',
    nombre: pac.Nombre||antecedentes.nombre||'',
    fecha_nac: pac.Fecha_Nac||'',
    telefono: pac.Telefono||'',
    email: pac.Email||'',
    obra_social: pac.Obra_Social||'', plan: pac.Plan||'', nro_afiliado: pac.N_Afiliado||'',
    domicilio: pac.Domicilio||'',
    centro_derivador: pac.Centro_Derivador||'',
    localidad: pac.Localidad||antecedentes.localidad||'',
    lugar_atencion: pac.Lugar_Atencion||'',
    estado_tratamiento: pac.Estado_Tratamiento||'',
    trabajos: trabajos,                                    // [{rubro,empresa,puesto,exposicion}]
    vivienda: {
      agua: vivienda.agua || antecedentes.vivienda_agua || '',
      excretas: vivienda.excretas || antecedentes.vivienda_excretas || '',
      calefaccion: vivienda.calefaccion || antecedentes.vivienda_calefaccion || '',
      ventilacion: vivienda.ventilacion || antecedentes.vivienda_ventilacion || ''
    },
    antecedentes: {
      enfermedades: antecedentes.enfermedades || '',        // incluye psiquiátricos (Ansiedad, Depresión, TOC, Psicosis, Insomnio)
      otras_enf: antecedentes.otras_enf || '',
      infancia: antecedentes.infancia || '',
      vacunas: antecedentes.vacunas || '',
      covid: antecedentes.covid || '',
      covid_secuelas: antecedentes.covid_secuelas || '',
      cirugias: antecedentes.cirugias || '',
      medicacion: antecedentes.medicacion || '',
      alergia: antecedentes.alergia || '',                  // "Sí" / "No" / "" (no completado)
      alergia_det: antecedentes.alergia_det || '',
      deporte: antecedentes.deporte || '', deporte_det: antecedentes.deporte_det || '',
      federado: antecedentes.federado || '', federado_det: antecedentes.federado_det || ''
    },
    // Fuente distinta de "antecedentes" de arriba (esos vienen del RIMP que
    // completa el paciente en casa). Estos los carga el médico directamente.
    antecedentes_generales: {
      motivo_inicial: antecedentesGenerales.motivo_inicial || '',
      enfermedades: antecedentesGenerales.enfermedades || '',
      cirugias: antecedentesGenerales.cirugias || '',
      medicacion_habitual: antecedentesGenerales.medicacion_habitual || '',
      actividad_laboral: antecedentesGenerales.actividad_laboral || '',
      actividad_deportiva: antecedentesGenerales.actividad_deportiva || ''
    },
    diagnosticos: historiaPac.map(h=>h.Diagnostico).filter(Boolean),
    // Historia completa (no solo la lista de diagnósticos de arriba) — la usa
    // el Espacio de Consulta Virtual de Agenda de Mensajes para mostrar las
    // consultas anteriores del paciente durante la entrevista.
    historia: historiaPac.map(h=>({
      Fecha: h.Fecha||'', Diagnostico: h.Diagnostico||'', Observaciones: h.Observaciones||'',
      Medicamentos: h.Medicamentos||'', Proximo_Control: h.Proximo_Control||'',
      Clasificacion: h.Clasificacion||'', Terapeutica: h.Terapeutica||''
    })),
    discapacidad: cudPac.map(c=>({tipo:c.Tipo, subtipo:c.Subtipo, diagnostico:c.Diagnostico, codigos_cif:c.Codigos_CIF}))
  };
}

function getVistaConsolidada(dni){
  if(!dni) return{success:false, error:'Falta el DNI'};
  return{success:true, vista: construirVistaConsolidada(dni)};
}

function getVistaConsolidadaTodos(){
  const cache = {
    pacientes: sheetToObjects('Pacientes'),
    historia: sheetToObjects('Historia_Clinica'),
    cud: sheetToObjects('Discapacidad_CUD'),
    ridm: sheetToObjects('RIDM_Pendientes')
  };
  const dnis = [...new Set(cache.pacientes.map(p=>String(p.DNI)).filter(Boolean))];
  return{success:true, vistas: dnis.map(dni=>construirVistaConsolidada(dni, cache))};
}

// ── Recordatorio de campos esenciales incompletos ──
// Cubre dos grupos de campos, cada uno con su propia sección en el RIDM:
//  - "Básicos" (Teléfono, Domicilio, Fecha de nacimiento, Localidad) viven en
//    la sección 0 del RIDM (Datos personales), que SOLO se muestra en
//    modo=rapido o modo=completo. Son los que más faltan en los pacientes
//    cargados desde historias clínicas en papel (los 104 pendientes).
//  - "Extenso" (Datos laborales, Vivienda y saneamiento) viven en la sección
//    4, solo visible en modo=clinico o modo=completo.
// BUGFIX: antes esta función solo miraba Localidad/Trabajos/Vivienda y
// siempre mandaba modo=clinico o clinico_breve — pero esos modos NO incluyen
// la sección 0, así que un paciente al que le faltaba Teléfono/Domicilio/
// Fecha de nacimiento recibía un link que ni siquiera le mostraba esos
// campos (y aunque "Localidad" sí se controlaba, tampoco se podía cargar
// por ese link, por la misma razón). Ahora se detectan los dos grupos por
// separado y se elige el modo que realmente los muestra.
function getPacientesCamposIncompletos(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName('Pacientes');
  if(!sheet) return{success:true,pacientes:[]};
  const headers = asegurarColumnasPacientes(sheet);
  const data = sheet.getDataRange().getValues();
  const idx = (h)=>headers.indexOf(h);
  const pacientes = [];
  for(let i=1;i<data.length;i++){
    const row = data[i];
    const email = row[idx('Email')];
    if(!email) continue; // sin mail no se le puede mandar el recordatorio

    const faltaTelefono  = !row[idx('Telefono')];
    const faltaDomicilio = !row[idx('Domicilio')];
    const faltaFechaNac  = !row[idx('Fecha_Nac')];
    const faltaLocalidad = !row[idx('Localidad')];
    const faltaBasicos   = faltaTelefono || faltaDomicilio || faltaFechaNac || faltaLocalidad;

    const faltaTrabajos  = !row[idx('Trabajos_JSON')];
    const faltaVivienda  = !row[idx('Vivienda_JSON')];
    const faltaExtenso   = faltaTrabajos || faltaVivienda;

    if(!faltaBasicos && !faltaExtenso) continue;

    let modo;
    if(faltaBasicos && faltaExtenso) modo = 'completo';
    else if(faltaBasicos) modo = 'rapido';
    else modo = 'clinico';

    pacientes.push({
      id: row[idx('ID')], apellido: row[idx('Apellido')], nombre: row[idx('Nombre')],
      dni: row[idx('DNI')], telefono: row[idx('Telefono')], email: email,
      falta: [
        faltaTelefono?'Teléfono':null,
        faltaDomicilio?'Domicilio':null,
        faltaFechaNac?'Fecha de nacimiento':null,
        faltaLocalidad?'Localidad':null,
        faltaTrabajos?'Datos laborales':null,
        faltaVivienda?'Vivienda y saneamiento':null
      ].filter(Boolean),
      modo: modo
    });
  }
  return{success:true,pacientes};
}

function enviarRecordatorioCamposIncompletos(d){
  const soloIds = Array.isArray(d && d.ids) && d.ids.length ? d.ids : null;
  const resultado = getPacientesCamposIncompletos();
  let pacientes = resultado.pacientes;
  if(soloIds) pacientes = pacientes.filter(p=>soloIds.indexOf(p.id)>-1);

  let enviados = 0, errores = 0;
  pacientes.forEach(p=>{
    try{
      const link = construirLinkRimpAntecedentes_(p);
      const cuerpo =
        `Hola ${p.nombre||''} ${p.apellido||''},\n\n` +
        `Para completar tu ficha médica nos falta cargar: ${p.falta.join(', ')}.\n\n` +
        `Por favor, completá estos datos entrando al siguiente link (te va a llevar directo a la sección que falta):\n${link}\n\n` +
        `Es rápido y nos ayuda a llevar un mejor seguimiento de tu salud.\n\n` +
        `Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943`;
      MailApp.sendEmail(p.email, 'Faltan completar algunos datos de tu ficha médica', cuerpo, {
        replyTo:'drluismuratori@gmail.com', name:'Dr. Luis Alberto Muratori'
      });
      enviados++;
      marcarRimpEnviado_(p.id); // mismo marcador que usa el envío inmediato de "primera consulta"
    }catch(e){ errores++; }
  });
  return{success:true, enviados, errores, total: pacientes.length};
}

// Ejecutar UNA VEZ manualmente desde el editor de Apps Script (▶ Ejecutar) para
// activar el envío automático semanal. No se llama desde el sistema (doPost).
function crearTriggerRecordatorioSemanal(){
  ScriptApp.getProjectTriggers().forEach(t=>{
    if(t.getHandlerFunction()==='enviarRecordatorioCamposIncompletos') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('enviarRecordatorioCamposIncompletos')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();
}

// ══════════════════════════════════════════════════════════════════
// PANEL DE CONTROL DE HISTORIAS (Panel_HC.html) — oct 2026
// Solo LEE (Pacientes + RIDM_Pendientes + Historia_Clinica) y arma, por
// paciente: si el RIMP se envió y si el paciente lo completó, qué datos
// esenciales faltan, y sus consultas. No inventa nada: si un dato no está
// en las hojas, queda vacío y el panel lo muestra como "falta".
// ══════════════════════════════════════════════════════════════════
function fechaISO_(v){
  if(!v) return '';
  if(Object.prototype.toString.call(v)==='[object Date]' && !isNaN(v)) return Utilities.formatDate(v,'America/Argentina/Buenos_Aires','yyyy-MM-dd');
  const t=String(v).trim();
  let m=t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if(m) return m[3]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[1]).slice(-2);
  m=t.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(m) return m[1]+'-'+m[2]+'-'+m[3];
  return '';
}
function getPanelHC(){
  const pacs = sheetToObjects('Pacientes');
  const ridm = sheetToObjects('RIDM_Pendientes');
  const hcs  = sheetToObjects('Historia_Clinica');
  const porDniRidm = {}, porDniHc = {}, porIdHc = {};
  ridm.forEach(r=>{ const k=String(r.DNI||'').trim(); if(!k) return; (porDniRidm[k]=porDniRidm[k]||[]).push(r); });
  hcs.forEach(h=>{
    const k=String(h.DNI||'').trim(); const id=String(h.ID_Paciente||'').trim();
    if(k) (porDniHc[k]=porDniHc[k]||[]).push(h);
    else if(id) (porIdHc[id]=porIdHc[id]||[]).push(h);
  });
  const lleno = v => v!==undefined && v!==null && String(v).trim()!=='' && String(v).trim()!=='—';
  const salida = [];
  pacs.forEach(p=>{
    const dni = String(p.DNI||'').trim();
    if(!dni && !lleno(p.Apellido)) return;
    // Antecedentes: lo que cargó el médico en filiación + lo que respondió el paciente (RIDM)
    let ag = {}; try{ ag = JSON.parse(p.Antecedentes_Generales_JSON||'{}')||{}; }catch(_){ ag = {}; }
    const filasRidm = (porDniRidm[dni]||[]).slice().sort((a,b)=>new Date(a.Timestamp)-new Date(b.Timestamp));
    let rid = {}, ultimoRidm = '';
    filasRidm.forEach(r=>{
      let d={}; try{ d=JSON.parse(r.Datos_JSON||'{}')||{}; }catch(_){ d={}; }
      Object.keys(d).forEach(k=>{ if(lleno(d[k])) rid[k]=d[k]; });
      ultimoRidm = (Object.prototype.toString.call(r.Timestamp)==='[object Date]') ? r.Timestamp.toISOString() : String(r.Timestamp||'');
    });
    const ant = {
      enfermedades: lleno(rid.enfermedades)||lleno(rid.otras_enf) ? [rid.enfermedades,rid.otras_enf].filter(lleno).join(', ') : (ag.enfermedades||''),
      cirugias: lleno(rid.cirugias) ? rid.cirugias : (ag.cirugias||''),
      medicacion: lleno(rid.medicacion) ? rid.medicacion : (ag.medicacion_habitual||''),
      alergia: lleno(rid.alergia) ? (rid.alergia + (lleno(rid.alergia_det)?(': '+rid.alergia_det):'')) : '',
      deporte: lleno(rid.deporte_det) ? rid.deporte_det : (lleno(rid.deporte)?rid.deporte:(ag.actividad_deportiva||'')),
      trabajo: lleno(rid.profesion) ? rid.profesion : (ag.actividad_laboral||''),
      motivo: lleno(rid.motivo) ? rid.motivo : (ag.motivo_inicial||''),
      zonas: rid.zonas||'',
      infancia: rid.infancia||'', vacunas: rid.vacunas||''
    };
    const cons = ((porDniHc[dni]||[]).concat(porIdHc[String(p.ID||'').trim()]||[])).map(h=>({
      fecha: fechaISO_(h.Fecha), dx: String(h.Diagnostico||''), cie: String(h.Codigo_CIE10||''),
      seg: String(h.Segmento_Corporal||''), tipo: String(h.Tipo_Patologia||''), lat: String(h.Lateralidad||''),
      eva: h.EVA_Dolor===''||h.EVA_Dolor===undefined?'':String(h.EVA_Dolor),
      clas: String(h.Clasificacion||''), terap: String(h.Terapeutica||''),
      med: String(h.Medicamentos||h.Farmaco||''), ind: String(h.Indicaciones||'').slice(0,600),
      obs: String(h.Observaciones||'').slice(0,600), prox: String(h.Proximo_Control||''),
      lugar: String(h.Lugar_Atencion||''), act: String(h.Tipo_Actividad_Fisica||''),
      imc: String(h.IMC||'')
    })).sort((a,b)=>String(b.fecha).localeCompare(String(a.fecha))).slice(0,60);
    const falta = [];
    if(!lleno(p.Telefono)) falta.push('Teléfono');
    if(!lleno(p.Domicilio)) falta.push('Domicilio');
    if(!lleno(p.Fecha_Nac)) falta.push('Fecha de nacimiento');
    if(!lleno(p.Obra_Social)) falta.push('Obra social');
    if(!lleno(p.Email)) falta.push('Mail');
    if(!lleno(ant.enfermedades)) falta.push('Enfermedades');
    if(!lleno(ant.cirugias)) falta.push('Cirugías');
    if(!lleno(ant.medicacion)) falta.push('Medicación');
    if(!lleno(ant.alergia)) falta.push('Alergias');
    salida.push({
      id:String(p.ID||''), apellido:String(p.Apellido||''), nombre:String(p.Nombre||''), dni:dni,
      fnac:fechaISO_(p.Fecha_Nac), os:String(p.Obra_Social||''), tel:String(p.Telefono||''),
      mail:String(p.Email||''), domicilio:String(p.Domicilio||''), localidad:String(p.Localidad||''),
      dx:String(p.Diagnostico_Principal||''),
      membresia:String(p.Membresia||'').toUpperCase()==='SI',
      rimp_enviado:String(p.RIMP_Enviado||''),
      rimp_completado:ultimoRidm,
      rimp_filas:filasRidm.length,
      falta:falta, ant:ant, consultas:cons
    });
  });
  return {success:true, generado:new Date().toISOString(), pacientes:salida};
}
function addReceta(d){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName('Recetas');
  const nReceta='REC-'+new Date().getFullYear()+'-'+String(sheet.getLastRow()).padStart(4,'0');
  sheet.appendRow([nReceta,d.ID_Paciente||'',new Date().toLocaleDateString('es-AR'),d.Med1||'',d.Dosis1||'',d.Med2||'',d.Dosis2||'',d.Med3||'',d.Dosis3||'',d.Diagnostico||'',d.Obra_Social||'',d.N_Afiliado||'','Generada','No']);
  return{success:true,nReceta:nReceta};
}
function addHistoria(d){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName('Historia_Clinica');
  let headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  let dniCol=headers.indexOf('DNI');
  if(dniCol===-1){
    dniCol=headers.length;
    sheet.getRange(1,dniCol+1).setValue('DNI');
    headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  }
  // Columnas aditivas: "Clasificacion" y "Terapeutica" — para que la lista de
  // Antecedentes y consultas anteriores en OCM (arriba del recetario) pueda
  // mostrar, además del diagnóstico, la clasificación clínica (ej. grado de
  // escoliosis) y la terapéutica indicada, como palabras cortas junto a la
  // fecha. Si la hoja no las tiene todavía, se agregan solas (no rompe filas
  // viejas — quedan con esa celda vacía). Pedido de CAM Verbal, sep 2026.
  // Lugar_Atencion (sep 2026, pedido de Luis): el paciente se puede atender
  // en cualquier consultorio — lo que hay que registrar es DÓNDE fue CADA
  // consulta puntual, no un consultorio fijo del paciente (eso era el bug
  // en guardarPaciente()/autoCrearFiliacionDesdeReceta_ de OCM, ya
  // corregido — ver comentario ahí). Acá sí corresponde: cada fila de
  // Historia_Clinica es una consulta concreta, y OCM manda el centro que
  // tenía activo en ese momento puntual.
  ['Clasificacion','Terapeutica','Lugar_Atencion'].forEach(col=>{
    if(headers.indexOf(col)===-1){
      sheet.getRange(1,headers.length+1).setValue(col);
      headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    }
  });
  const id='HC-'+String(sheet.getLastRow()).padStart(4,'0');
  const fila=new Array(headers.length).fill('');
  const set=(nombreCol,valor)=>{const i=headers.indexOf(nombreCol);if(i>-1)fila[i]=valor;};
  set('ID_HC',id);
  set('ID_Paciente',d.ID_Paciente||'');
  set('Fecha',new Date().toLocaleDateString('es-AR'));
  set('Diagnostico',d.Diagnostico||'');
  set('Medicamentos',d.Medicamentos||'');
  set('Estudios_Realizados',d.Estudios_Realizados||'');
  set('Estudios_Pendientes',d.Estudios_Pendientes||'');
  set('Indicaciones',d.Indicaciones||'');
  set('Observaciones',d.Observaciones||'');
  set('Proximo_Control',d.Proximo_Control||'');
  set('DNI',d.DNI||'');
  set('Clasificacion',d.Clasificacion||'');
  set('Terapeutica',d.Terapeutica||'');
  set('Lugar_Atencion',d.Lugar_Atencion||'');
  sheet.appendRow(fila);
  return{success:true,id:id};
}

// ── Funciones NUEVAS ──
function getRidmPendientes(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName('RIDM_Pendientes');
  if(!sheet) return{success:true,pendientes:[]};
  const data=sheet.getDataRange().getValues();
  if(data.length<2) return{success:true,pendientes:[]};
  const headers=data[0];
  const pendientes=data.slice(1).map((row,i)=>{
    const obj={_row:i+2};
    headers.forEach((h,j)=>{obj[h.toLowerCase()]=row[j];});
    return obj;
  }).filter(r=>String(r.procesado).toUpperCase()!=='SI');
  return{success:true,pendientes:pendientes};
}
function marcarRidmProcesado(rowNum){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName('RIDM_Pendientes');
  if(!sheet||!rowNum) return{success:false,error:'Falta hoja o número de fila'};
  const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const col=headers.indexOf('Procesado')+1;
  sheet.getRange(parseInt(rowNum),col).setValue('SI');
  return{success:true};
}
// Completa Fecha_Nac/Telefono/Domicilio/Localidad/Email/Trabajos_JSON/
// Vivienda_JSON en la ficha de un paciente YA CARGADO (busca por DNI), pero
// SOLO los campos que vengan con datos Y que en la ficha estén vacíos —
// nunca pisa un valor que ya existía. Trabajos_JSON/Vivienda_JSON son datos
// laborales/de vivienda (no antecedentes médicos), por eso es seguro
// completarlos solos igual que el teléfono o el domicilio. Los antecedentes
// CLÍNICOS de verdad (enfermedades, cirugías, medicación, alergias, motivo
// de consulta) nunca se tocan acá — eso sigue yendo a RIDM_Pendientes para
// que el médico lo revise e importe a mano vía "📥 Importar" en OCM.
// Devuelve true si encontró al paciente por DNI (se haya modificado algo o
// no), false si el DNI no corresponde a ningún paciente cargado todavía
// (caso "primera consulta": ese sí necesita el alta manual de siempre).
function actualizarDatosBasicosPaciente_(dni, datos){
  const dniStr = String(dni||'').trim();
  if(!dniStr) return false;
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sheet=ss.getSheetByName('Pacientes');
  if(!sheet) return false;
  const headers = asegurarColumnasPacientes(sheet);
  const colDNI = headers.indexOf('DNI');
  if(colDNI===-1) return false;
  const data = sheet.getDataRange().getValues();
  const viv = datos.vivienda || {};
  const trabajosJSON = (Array.isArray(datos.trabajos) && datos.trabajos.length) ? JSON.stringify(datos.trabajos) : undefined;
  const vivJSON = (viv.agua || viv.excretas || viv.calefaccion || viv.ventilacion) ? JSON.stringify(viv) : undefined;
  for(let i=1;i<data.length;i++){
    if(String(data[i][colDNI]) !== dniStr) continue;
    const mapa = {
      Fecha_Nac: datos.fnac, Telefono: datos.tel, Domicilio: datos.domicilio,
      Localidad: datos.localidad, Email: datos.mail,
      Trabajos_JSON: trabajosJSON, Vivienda_JSON: vivJSON
    };
    Object.keys(mapa).forEach(campo=>{
      const valorNuevo = mapa[campo];
      if(valorNuevo===undefined || valorNuevo==='') return; // no hay nada que completar
      const col = headers.indexOf(campo);
      if(col===-1) return;
      if(data[i][col] !== '' && data[i][col] !== undefined && data[i][col] !== null) return; // ya tenía un valor — no se pisa
      sheet.getRange(i+1, col+1).setValue(valorNuevo);
    });
    return true;
  }
  return false;
}
// ¿Este envío de RIDM trae contenido clínico de verdad (motivo, enfermedades,
// cirugías, medicación, alergias, antecedentes de infancia/vacunas/COVID)?
// Si trae, tiene que quedar pendiente de revisión del médico en RIDM_Pendientes
// pase lo que pase con los datos básicos/laborales — nunca se auto-procesa
// un envío con contenido médico sin que el doctor lo vea primero.
function tieneContenidoClinicoRidm_(datos){
  datos = datos || {};
  const campos = ['motivo','enfermedades','infancia','vacunas','covid','covid_secuelas','otras_enf','cirugias','medicacion','alergia_det'];
  if(campos.some(c => String(datos[c]||'').trim())) return true;
  if(String(datos.alergia||'').trim()==='Sí') return true;
  return false;
}
function guardarRidmSubmit(d){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName('RIDM_Pendientes');
  if(!sheet){
    sheet=ss.insertSheet('RIDM_Pendientes');
    sheet.appendRow(['Timestamp','DNI','Apellido','Nombre','Telefono','Email','Datos_JSON','Procesado']);
  }
  // ── Auto-completar datos básicos si el DNI ya es un paciente conocido ──
  // Antes, CUALQUIER envío del RIDM (aunque solo fuera para completar
  // Teléfono/Domicilio/Fecha de nacimiento de un paciente que ya estaba
  // cargado) quedaba en RIDM_Pendientes esperando que alguien lo
  // "Importara" a mano desde OCM — así que "no quedan pendientes" no se
  // sostenía: el paciente completaba el formulario pero seguía apareciendo
  // en la lista de incompletos (getPacientesCamposIncompletos) hasta ese
  // clic extra. Como acá solo se tocan datos de contacto — nunca
  // antecedentes clínicos — es seguro completarlos solos.
  const datosEnvio = d.datos || {};
  const yaEraPaciente = actualizarDatosBasicosPaciente_(d.dni, {
    fnac: datosEnvio.fnac, tel: d.tel, domicilio: datosEnvio.domicilio,
    localidad: datosEnvio.localidad, mail: d.mail, trabajos: datosEnvio.trabajos,
    vivienda: { agua: datosEnvio.vivienda_agua, excretas: datosEnvio.vivienda_excretas,
                calefaccion: datosEnvio.vivienda_calefaccion, ventilacion: datosEnvio.vivienda_ventilacion }
  });
  // Si el envío trae contenido clínico real (esto pasa en modo=completo,
  // que además de los datos básicos también muestra antecedentes médicos y
  // motivo de consulta), la fila NO se marca como procesada aunque los
  // datos básicos/laborales ya se hayan volcado solos — el médico igual
  // tiene que revisar e importar esa parte médica a mano.
  const necesitaRevisionClinica = tieneContenidoClinicoRidm_(datosEnvio);
  const procesadoAuto = yaEraPaciente && !necesitaRevisionClinica;
  // 'SI' exacto (no 'SI (auto)') porque getRidmPendientes()/getRidmPaciente()
  // comparan con === 'SI' — un valor distinto igual aparecería como
  // pendiente de importar en el modal de OCM, aunque ya se haya completado solo.
  sheet.appendRow([new Date().toISOString(),d.dni||'',d.apellido||'',d.nombre||'',d.tel||'',d.mail||'',JSON.stringify(d.datos||{}), procesadoAuto?'SI':'NO']);
  // ── Filiación rápida: mandar automáticamente por MAIL el link para que el
  // paciente complete sus antecedentes (Parte 2) — antes esto dependía de que
  // el médico se lo copiara/mandara a mano por WhatsApp. Si no cargó mail en
  // Filiación, no se manda nada (no rompe el guardado; sigue disponible el
  // botón de copiar link en pantalla — WhatsApp sacado del sistema, sep 2026). ──
  // BUGFIX (sep 2026): antes el catch estaba vacío — si MailApp.sendEmail
  // fallaba (cuota de Gmail agotada, mail mal escrito, etc.) el error se
  // perdía por completo: no quedaba registrado en ningún lado y el mail
  // nunca salía, aunque el resto del guardado funcionara bien. Ahora se
  // deja registrado en Logger (Ejecuciones, en el editor de Apps Script)
  // y se devuelve el motivo real en mail_error para que el front-end lo
  // pueda mostrar en vez de solo decir "no se confirmó el envío".
  let mailEnviado = false, mailError = null;
  if(d.mail){
    try{
      enviarLinkRimpAPaciente_({ apellido:d.apellido||'', nombre:d.nombre||'', dni:d.dni||'', telefono:d.tel||'', email:d.mail, extenso:false });
      mailEnviado = true;
    }catch(e){
      mailError = String(e);
      Logger.log('guardarRidmSubmit: fallo el envio de mail a ' + d.mail + ' — ' + mailError);
    }
  }
  return{success:true,message:'RIDM guardado, en espera de importar a OCM',mail_enviado:mailEnviado,mail_error:mailError};
}
// ── Parte 2 del RIMP (antecedentes clínicos que el paciente completa en casa) ──
// Llega cuando el paciente completa el RIMP en modo=clinico.
// Se guarda igual que un RIDM_Pendientes para que OCM lo muestre con el badge
// y el médico pueda importarlo a la ficha del paciente con un toque.
// El campo datos.modo='clinico' distingue este envío de uno de primera vez.
function guardarRidmAntecedentes(d){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName('RIDM_Pendientes');
  if(!sheet){
    sheet=ss.insertSheet('RIDM_Pendientes');
    sheet.appendRow(['Timestamp','DNI','Apellido','Nombre','Telefono','Email','Datos_JSON','Procesado']);
  }
  // Datos laborales/de vivienda (Trabajos_JSON/Vivienda_JSON) NO son
  // antecedentes médicos — se completan solos en la ficha igual que en
  // guardarRidmSubmit, para que el paciente salga de la lista de datos
  // incompletos apenas los manda. Los antecedentes clínicos de verdad
  // (enfermedades, cirugías, medicación, alergias, motivo) siguen esperando
  // acá abajo a que el médico los revise e importe a mano.
  const datosEnvio = d.datos || {};
  actualizarDatosBasicosPaciente_(d.dni, {
    trabajos: datosEnvio.trabajos,
    vivienda: { agua: datosEnvio.vivienda_agua, excretas: datosEnvio.vivienda_excretas,
                calefaccion: datosEnvio.vivienda_calefaccion, ventilacion: datosEnvio.vivienda_ventilacion }
  });
  // Buscar si ya hay una fila no procesada para este DNI y actualizarla
  // (para no acumular múltiples envíos de antecedentes del mismo paciente)
  const data=sheet.getDataRange().getValues();
  const headers=data[0];
  const dniColIdx=headers.indexOf('DNI');
  const procColIdx=headers.indexOf('Procesado');
  const datosColIdx=headers.indexOf('Datos_JSON');
  for(let i=1;i<data.length;i++){
    if(String(data[i][dniColIdx])===String(d.dni||'') && String(data[i][procColIdx]).toUpperCase()!=='SI'){
      // Mezclar los antecedentes con los datos que ya están en esa fila
      let prevDatos={};
      try{ prevDatos=JSON.parse(data[i][datosColIdx]||'{}'); }catch(_){}
      const merged=Object.assign({},prevDatos,d.datos||{},{_tipo:'antecedentes+merge'});
      sheet.getRange(i+1,datosColIdx+1).setValue(JSON.stringify(merged));
      return{success:true,message:'Antecedentes agregados al RIDM pendiente existente'};
    }
  }
  // Si no hay fila previa, crear una nueva
  sheet.appendRow([new Date().toISOString(),d.dni||'',d.apellido||'',d.nombre||'',d.tel||'',d.mail||'',JSON.stringify(Object.assign({},d.datos||{},{_tipo:'antecedentes'})),'NO']);
  return{success:true,message:'Antecedentes guardados como nuevo RIDM pendiente'};
}
// ══════════════════════════════════════════════════════════════════
// Trae los antecedentes RIMP de UN paciente puntual, tal cual los mandó
// desde su casa (mismo formato que guardarRidmSubmit/guardarRidmAntecedentes
// escriben en Datos_JSON) — para que OCM se los pueda mostrar al médico
// durante la consulta sin tener que ir a buscarlos a mano a la planilla.
// Antes esto solo se leía de localStorage del propio navegador del médico,
// por eso nunca aparecía: el paciente lo completa en SU celular, no en la
// compu del consultorio.
// ══════════════════════════════════════════════════════════════════
// ══════════════════════════════════════════════════════════════════
// RESUMEN DE HISTORIA CLÍNICA CON FIRMA TÁCTIL — lo más simple posible,
// copiando el mismo patrón ya probado de RIDM_Pendientes (una hoja simple,
// guardar con appendRow/actualizar por DNI, leer filtrando por DNI):
//   1) OCM arma el texto del resumen y lo guarda acá (Firmado=NO).
//   2) Se manda al paciente el link de Autorizacion_HC_v1.html?dni=...
//   3) El paciente firma con el dedo (canvas) en esa página — eso llama
//      a guardarFirmaPaciente, que guarda la imagen y pone Firmado=SI.
//   4) OCM consulta getEstadoFirma para saber si ya firmó, y recién ahí
//      habilita imprimir/enviar.
// La firma táctil no existía en NINGÚN lugar del sistema (confirmado al
// buscar 'firma'+canvas en todo el código) — es la única pieza realmente
// nueva; el resto reutiliza el mecanismo de RIDM tal cual.
// ══════════════════════════════════════════════════════════════════
function hojaResumenFirma_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName('Resumenes_HC_Firma');
  if(!sheet){
    sheet=ss.insertSheet('Resumenes_HC_Firma');
    sheet.appendRow(['Timestamp','DNI','Apellido','Nombre','Resumen_Texto','Firmado','Firma_Base64','Fecha_Firma','Firma_Token']);
  } else {
    // MIGRACIÓN (sep 2026): agrega la columna Firma_Token si la hoja ya
    // existía de antes (no pisa ninguna fila existente).
    const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    if(headers.indexOf('Firma_Token')===-1){
      sheet.getRange(1, sheet.getLastColumn()+1).setValue('Firma_Token');
    }
  }
  return sheet;
}
// CORREGIDO (sep 2026): antes esto se protegía SOLO con el DNI (un número
// corto y adivinable, no un secreto) — cualquiera que probara DNIs al
// tanteo podía leer el resumen de historia clínica de otro paciente, o
// directamente firmar en su nombre. Ahora cada resumen tiene su propio
// Firma_Token (un UUID) que se genera acá, se manda en el link que recibe
// el paciente (?dni=...&t=<token>), y getResumenParaFirma/guardarFirmaPaciente
// lo exigen. El médico (OCM) sigue entrando igual con el token compartido
// de siempre (APP_TOKEN) — ver el doble chequeo más abajo.
function guardarResumenParaFirma(d){
  if(!d.dni) return {success:false, error:'Falta el DNI'};
  const token = Utilities.getUuid();
  const sheet=hojaResumenFirma_();
  const data=sheet.getDataRange().getValues();
  const headers=data[0];
  const dniCol=headers.indexOf('DNI');
  for(let i=1;i<data.length;i++){
    if(String(data[i][dniCol])===String(d.dni)){
      // Ya había un resumen para este DNI: se reemplaza (no se acumulan
      // filas viejas) y se vuelve a pedir la firma desde cero, porque el
      // texto cambió. Se genera un Firma_Token nuevo también, así un link
      // viejo ya mandado por mail deja de servir para el resumen anterior.
      sheet.getRange(i+1,1,1,9).setValues([[new Date().toISOString(), d.dni, d.apellido||'', d.nombre||'', d.resumen||'', 'NO', '', '', token]]);
      return {success:true, message:'Resumen actualizado, a la espera de la firma', token:token};
    }
  }
  sheet.appendRow([new Date().toISOString(), d.dni, d.apellido||'', d.nombre||'', d.resumen||'', 'NO', '', '', token]);
  return {success:true, message:'Resumen guardado, a la espera de la firma', token:token};
}
// dni: obligatorio. appToken: el token compartido de las apps (médico/OCM,
// vía parámetro "token"). firmaToken: el token propio de ESTE resumen (vía
// parámetro "t", el que recibe el paciente en el link). Alcanza con que
// UNO de los dos sea válido. Si el resumen es de antes de esta actualización
// (Firma_Token vacío, columna migrada recién), se sigue permitiendo el
// acceso solo por DNI como hasta ahora — así no se corta ningún link ya
// mandado — hasta que el médico genere el resumen de nuevo.
function getResumenParaFirma(dni, appToken, firmaToken){
  if(!dni) return {success:false, error:'Falta el DNI'};
  const sheet=hojaResumenFirma_();
  const data=sheet.getDataRange().getValues();
  if(data.length<2) return {success:true, encontrado:false};
  const headers=data[0];
  const filas=data.slice(1).map(row=>{const o={}; headers.forEach((h,j)=>o[h]=row[j]); return o;})
    .filter(r=>String(r.DNI)===String(dni));
  if(!filas.length) return {success:true, encontrado:false};
  const f=filas[filas.length-1];
  const tokenGuardado=String(f.Firma_Token||'');
  const okApp = tokenValido_(appToken);
  const okFirma = !tokenGuardado || String(firmaToken||'')===tokenGuardado;
  if(!okApp && !okFirma) return {success:false, error:'Token inválido o faltante'};
  return {success:true, encontrado:true, apellido:f.Apellido, nombre:f.Nombre, resumen:f.Resumen_Texto, firmado: String(f.Firmado).toUpperCase()==='SI'};
}
function guardarFirmaPaciente(d){
  if(!d.dni) return {success:false, error:'Falta el DNI'};
  if(!d.firma_base64) return {success:false, error:'Falta la firma'};
  const sheet=hojaResumenFirma_();
  const data=sheet.getDataRange().getValues();
  const headers=data[0];
  const dniCol=headers.indexOf('DNI'), firmadoCol=headers.indexOf('Firmado'), firmaCol=headers.indexOf('Firma_Base64'), fechaCol=headers.indexOf('Fecha_Firma'), tokenCol=headers.indexOf('Firma_Token');
  for(let i=1;i<data.length;i++){
    if(String(data[i][dniCol])===String(d.dni)){
      const tokenGuardado=String((tokenCol>-1?data[i][tokenCol]:'')||'');
      const okApp = tokenValido_(d.token);
      const okFirma = !tokenGuardado || String(d.t||'')===tokenGuardado;
      if(!okApp && !okFirma) return {success:false, error:'Token inválido o faltante'};
      sheet.getRange(i+1,firmadoCol+1).setValue('SI');
      sheet.getRange(i+1,firmaCol+1).setValue(d.firma_base64);
      sheet.getRange(i+1,fechaCol+1).setValue(new Date().toISOString());
      return {success:true, message:'Firma guardada correctamente'};
    }
  }
  return {success:false, error:'No se encontró un resumen pendiente para ese DNI — pedile al médico que lo genere de nuevo'};
}
function getEstadoFirma(dni, appToken, firmaToken){
  return getResumenParaFirma(dni, appToken, firmaToken);
}
// Envía por mail el resumen YA FIRMADO (a quien lo haya solicitado — no
// tiene por qué ser el mail del propio paciente, por eso pide destinatario
// aparte). Si todavía no firmó, no manda nada — se lo avisa a OCM.
function enviarResumenFirmadoPorMail(d){
  if(!d.dni) return {success:false, error:'Falta el DNI'};
  if(!d.destinatario) return {success:false, error:'Falta el mail de destino'};
  const r = getResumenParaFirma(d.dni);
  if(!r.success || !r.encontrado) return {success:false, error:'No hay un resumen guardado para ese DNI'};
  if(!r.firmado) return {success:false, error:'El paciente todavía no firmó — no se puede enviar'};
  const sheet=hojaResumenFirma_();
  const data=sheet.getDataRange().getValues();
  const headers=data[0];
  const dniCol=headers.indexOf('DNI'), firmaCol=headers.indexOf('Firma_Base64');
  let firmaBase64='';
  for(let i=1;i<data.length;i++){ if(String(data[i][dniCol])===String(d.dni)){ firmaBase64=data[i][firmaCol]; break; } }
  let inlineImages={};
  let firmaHtml='';
  try{
    if(firmaBase64){
      const b64=firmaBase64.replace(/^data:image\/\w+;base64,/,'');
      inlineImages.firmaPacienteImg = Utilities.newBlob(Utilities.base64Decode(b64), 'image/png', 'firma.png');
      firmaHtml = '<p>Firma del paciente/tutor:</p><img src="cid:firmaPacienteImg" style="max-width:260px">';
    }
  }catch(e){ firmaHtml=''; }
  const htmlBody = `
    <div style="font-family:Arial,sans-serif;font-size:13px;color:#0f172a">
      <h3>Resumen de Historia Clínica — ${r.apellido||''}, ${r.nombre||''} (DNI ${d.dni})</h3>
      <pre style="white-space:pre-wrap;font-family:inherit;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:12px">${(r.resumen||'').replace(/[<>&]/g, c=>({'<':'&lt;','>':'&gt;','&':'&amp;'}[c]))}</pre>
      ${firmaHtml}
      <p style="color:#64748b;font-size:11px;margin-top:16px">Autorizado por el paciente/tutor mediante firma táctil registrada en el sistema — Dr. Luis Alberto Muratori (M.N. 100.540 / M.P. 9943).</p>
    </div>`;
  try{
    MailApp.sendEmail({to:d.destinatario, subject:d.asunto||'Resumen de Historia Clínica — Dr. Muratori', htmlBody, inlineImages, name:'Dr. Luis Alberto Muratori'});
    return {success:true, message:'Enviado a '+d.destinatario};
  }catch(e){
    return {success:false, error:'No se pudo enviar el mail: '+e};
  }
}

// ══════════════════════════════════════════════════════════════════
// CONSENTIMIENTO INFORMADO QUIRÚRGICO — FIRMA TÁCTIL. Mismo mecanismo que
// Resumenes_HC_Firma (arriba), extendido a pedido del Dr. Muratori para que
// la firma forme parte del Libro Quirúrgico y no quede separada del resto
// de la documentación (sep 2026).
// ══════════════════════════════════════════════════════════════════
function hojaConsentimientoFirma_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName('Consentimientos_QX_Firma');
  if(!sheet){
    sheet=ss.insertSheet('Consentimientos_QX_Firma');
    sheet.appendRow(['Timestamp','DNI','Apellido','Nombre','Procedimiento','Texto_Consentimiento','Firmado','Firma_Base64','Fecha_Firma','Firma_Token']);
  } else {
    // MIGRACIÓN (sep 2026): mismo criterio que hojaResumenFirma_ — agrega
    // la columna si la hoja ya existía, sin tocar filas viejas.
    const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    if(headers.indexOf('Firma_Token')===-1){
      sheet.getRange(1, sheet.getLastColumn()+1).setValue('Firma_Token');
    }
  }
  // MIGRACIÓN (oct 2026): quién firmó — paciente o tutor/representante legal
  // (nombre, DNI y vínculo). Agrega las columnas si faltan, sin tocar filas viejas.
  const hdr2=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  ['Firmante_Tipo','Firmante_Nombre','Firmante_DNI','Firmante_Vinculo'].forEach(function(c){
    if(hdr2.indexOf(c)===-1){ sheet.getRange(1, sheet.getLastColumn()+1).setValue(c); hdr2.push(c); }
  });
  return sheet;
}
// Mismo mecanismo de Firma_Token que guardarResumenParaFirma — ver el
// comentario completo ahí arriba.
function guardarConsentimientoParaFirma(d){
  if(!d.dni) return {success:false, error:'Falta el DNI'};
  const token = Utilities.getUuid();
  const sheet=hojaConsentimientoFirma_();
  const data=sheet.getDataRange().getValues();
  const headers=data[0];
  const dniCol=headers.indexOf('DNI');
  for(let i=1;i<data.length;i++){
    if(String(data[i][dniCol])===String(d.dni)){
      // Igual que con el Resumen de HC: se reemplaza el pendiente anterior
      // (si había) y se vuelve a pedir la firma desde cero, con un
      // Firma_Token nuevo.
      sheet.getRange(i+1,1,1,10).setValues([[new Date().toISOString(), d.dni, d.apellido||'', d.nombre||'', d.procedimiento||'', d.texto||'', 'NO', '', '', token]]);
      // se pide la firma desde cero: se borra quién había firmado antes
      ['Firmante_Tipo','Firmante_Nombre','Firmante_DNI','Firmante_Vinculo'].forEach(function(c){
        const k=headers.indexOf(c); if(k>-1) sheet.getRange(i+1,k+1).setValue('');
      });
      return {success:true, message:'Consentimiento actualizado, a la espera de la firma', token:token};
    }
  }
  sheet.appendRow([new Date().toISOString(), d.dni, d.apellido||'', d.nombre||'', d.procedimiento||'', d.texto||'', 'NO', '', '', token]);
  return {success:true, message:'Consentimiento guardado, a la espera de la firma', token:token};
}
// Mismo doble chequeo (appToken del médico/OCM O firmaToken propio del
// documento) que getResumenParaFirma — ver el comentario completo ahí arriba.
function getConsentimientoParaFirma(dni, appToken, firmaToken){
  if(!dni) return {success:false, error:'Falta el DNI'};
  const sheet=hojaConsentimientoFirma_();
  const data=sheet.getDataRange().getValues();
  if(data.length<2) return {success:true, encontrado:false};
  const headers=data[0];
  const filas=data.slice(1).map(row=>{const o={}; headers.forEach((h,j)=>o[h]=row[j]); return o;})
    .filter(r=>String(r.DNI)===String(dni));
  if(!filas.length) return {success:true, encontrado:false};
  const f=filas[filas.length-1];
  const tokenGuardado=String(f.Firma_Token||'');
  const okApp = tokenValido_(appToken);
  const okFirma = !tokenGuardado || String(firmaToken||'')===tokenGuardado;
  if(!okApp && !okFirma) return {success:false, error:'Token inválido o faltante'};
  const firmado = String(f.Firmado).toUpperCase()==='SI';
  const out = {success:true, encontrado:true, apellido:f.Apellido, nombre:f.Nombre, procedimiento:f.Procedimiento, texto:f.Texto_Consentimiento, firmado: firmado,
    fecha_firma: firmado ? String(f.Fecha_Firma||'') : '',
    firmante_tipo: firmado ? String(f.Firmante_Tipo||'paciente') : '',
    firmante_nombre: firmado ? String(f.Firmante_Nombre||'') : '',
    firmante_dni: firmado ? String(f.Firmante_DNI||'') : '',
    firmante_vinculo: firmado ? String(f.Firmante_Vinculo||'') : ''};
  // La imagen de la firma SOLO se entrega a OCM (con el token de la app
  // configurado y correcto) — nunca al link que recibe el paciente.
  const realTok = getAppToken_();
  if(firmado && realTok && String(appToken||'')===String(realTok)) out.firma_base64 = String(f.Firma_Base64||'');
  return out;
}
function guardarFirmaConsentimiento(d){
  if(!d.dni) return {success:false, error:'Falta el DNI'};
  if(!d.firma_base64) return {success:false, error:'Falta la firma'};
  const sheet=hojaConsentimientoFirma_();
  const data=sheet.getDataRange().getValues();
  const headers=data[0];
  const dniCol=headers.indexOf('DNI'), firmadoCol=headers.indexOf('Firmado'), firmaCol=headers.indexOf('Firma_Base64'), fechaCol=headers.indexOf('Fecha_Firma'), tokenCol=headers.indexOf('Firma_Token');
  // Quien firma: paciente o tutor/representante legal (oct 2026)
  const tipoF = String(d.firmante_tipo||'paciente')==='tutor' ? 'tutor' : 'paciente';
  const nomF = String(d.firmante_nombre||'').trim().substring(0,120);
  const dniF = String(d.firmante_dni||'').replace(/\D/g,'').substring(0,15);
  const vinF = String(d.firmante_vinculo||'').trim().substring(0,60);
  if(tipoF==='tutor' && (nomF.length<5 || dniF.length<6 || !vinF)) return {success:false, error:'Faltan los datos del tutor o representante legal (nombre, DNI y vinculo)'};
  for(let i=1;i<data.length;i++){
    if(String(data[i][dniCol])===String(d.dni)){
      const tokenGuardado=String((tokenCol>-1?data[i][tokenCol]:'')||'');
      const okApp = tokenValido_(d.token);
      const okFirma = !tokenGuardado || String(d.t||'')===tokenGuardado;
      if(!okApp && !okFirma) return {success:false, error:'Token inválido o faltante'};
      sheet.getRange(i+1,firmadoCol+1).setValue('SI');
      sheet.getRange(i+1,firmaCol+1).setValue(d.firma_base64);
      sheet.getRange(i+1,fechaCol+1).setValue(new Date().toISOString());
      const setF=function(c,v){ const k=headers.indexOf(c); if(k>-1) sheet.getRange(i+1,k+1).setValue(v); };
      setF('Firmante_Tipo',tipoF); setF('Firmante_Nombre',tipoF==='tutor'?nomF:''); setF('Firmante_DNI',tipoF==='tutor'?dniF:''); setF('Firmante_Vinculo',tipoF==='tutor'?vinF:'');
      return {success:true, message:'Firma guardada correctamente'};
    }
  }
  return {success:false, error:'No se encontró un consentimiento pendiente para ese DNI — pedile al médico que lo genere de nuevo'};
}
function getEstadoFirmaConsentimiento(dni, appToken, firmaToken){
  return getConsentimientoParaFirma(dni, appToken, firmaToken);
}

function getRidmPaciente(dni){
  if(!dni) return {success:false, error:'Falta el DNI'};
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sheet=ss.getSheetByName('RIDM_Pendientes');
  if(!sheet) return {success:true, encontrado:false};
  const data=sheet.getDataRange().getValues();
  if(data.length<2) return {success:true, encontrado:false};
  const headers=data[0];
  const filas = data.slice(1).map((row,i)=>{
    const obj={_row:i+2};
    headers.forEach((h,j)=>{obj[h]=row[j];});
    return obj;
  }).filter(r=>String(r.DNI)===String(dni));
  if(!filas.length) return {success:true, encontrado:false};
  filas.sort((a,b)=>new Date(b.Timestamp)-new Date(a.Timestamp));
  // Preferimos la más reciente que todavía no se importó; si ya están todas
  // importadas, igual mostramos la más reciente (por si el médico la quiere
  // volver a consultar).
  const fila = filas.find(r=>String(r.Procesado).toUpperCase()!=='SI') || filas[0];
  let datos={};
  try{ datos=JSON.parse(fila.Datos_JSON||'{}'); }catch(_){}
  const merged = Object.assign({
    dni: fila.DNI, apellido: fila.Apellido, nombre: fila.Nombre,
    tel: fila.Telefono, mail: fila.Email,
    fecha_envio: (fila.Timestamp instanceof Date) ? fila.Timestamp.toLocaleString('es-AR') : fila.Timestamp
  }, datos, { _row: fila._row });
  return {success:true, encontrado:true, procesado: String(fila.Procesado).toUpperCase()==='SI', datos: merged};
}
function guardarBackup(d){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName('Backups');
  if(!sheet){
    sheet=ss.insertSheet('Backups');
    sheet.appendRow(['Timestamp','Resumen_JSON']);
  }
  sheet.appendRow([d.timestamp||new Date().toISOString(),d.data||'']);
  return{success:true,message:'Backup guardado'};
}
// ══════════════════════════════════════════════════════════════════
// NEURO-ORTOPEDIA / PARÁLISIS CEREBRAL — escalas de seguimiento pedidas por
// el Dr. Muratori (sep 2026), en base al nomenclador de SANeO:
//   GMFCS y MACS (función motora gruesa / habilidad manual, I-V),
//   Índice de Reimers (% de migración de cadera, cada lado),
//   Ashworth Modificada y Tardieu Modificada (tono/espasticidad, por grupo
//   muscular, R1/R2 para Tardieu), y Thomas-Stonell y Greenberg (babeo).
// Se guarda UNA FILA POR EVALUACIÓN (nunca se pisan entre sí) a propósito:
// lo clínico acá es la EVOLUCIÓN — el Índice de Reimers en particular hay
// que vigilarlo seriado para agarrar una subluxación de cadera antes de
// que duela, y Ashworth/Tardieu cambian con el tratamiento (toxina
// botulínica, yesos seriados). Se usa desde CAM — Herramientas Clínicas.
// ══════════════════════════════════════════════════════════════════
function hojaNeuroortopedia_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Neuroortopedia_PC');
  if(!sheet){
    sheet = ss.insertSheet('Neuroortopedia_PC');
    sheet.appendRow(['Timestamp','DNI','Paciente','Fecha_Evaluacion','Clasificacion_Anatomica',
      'GMFCS','MACS','Reimers_Cadera_D','Reimers_Cadera_I','Ashworth_JSON','Tardieu_JSON',
      'Babeo_Severidad','Babeo_Frecuencia','Observaciones']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
function guardarNeuroortopedia(d){
  if(!d || !String(d.dni||'').trim()) return {success:false, error:'Falta el DNI del paciente'};
  const sheet = hojaNeuroortopedia_();
  const numOVacio = v => (v!==undefined && v!==null && String(v).trim()!=='') ? Number(v) : '';
  sheet.appendRow([
    new Date().toISOString(), d.dni||'', d.paciente||'',
    d.fechaEvaluacion || new Date().toLocaleDateString('es-AR'),
    d.clasificacionAnatomica||'', d.gmfcs||'', d.macs||'',
    numOVacio(d.reimersD), numOVacio(d.reimersI),
    d.ashworth ? JSON.stringify(d.ashworth) : '',
    d.tardieu ? JSON.stringify(d.tardieu) : '',
    d.babeoSeveridad||'', d.babeoFrecuencia||'',
    d.observaciones||''
  ]);
  return {success:true, message:'Evaluación de neuro-ortopedia guardada'};
}
// Historial de evaluaciones de UN paciente puntual, más reciente primero —
// para que CAM le muestre al médico la evolución (ej. Reimers subiendo de
// una consulta a la siguiente) sin ir a buscarlo a mano a la planilla.
function getNeuroortopedia(dni){
  if(!dni) return {success:false, error:'Falta el DNI'};
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sheet = ss.getSheetByName('Neuroortopedia_PC');
  if(!sheet) return {success:true, evaluaciones:[]};
  const data = sheet.getDataRange().getValues();
  if(data.length<2) return {success:true, evaluaciones:[]};
  const headers = data[0];
  const evaluaciones = data.slice(1)
    .map(row=>{ const obj={}; headers.forEach((h,i)=>{obj[h]=row[i];}); return obj; })
    .filter(r=>String(r.DNI)===String(dni))
    .sort((a,b)=> new Date(b.Timestamp) - new Date(a.Timestamp));
  return {success:true, evaluaciones};
}
function guardarKineEvolucion(d){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName('Kinesiologia');
  if(!sheet){
    sheet=ss.insertSheet('Kinesiologia');
    // Vals_JSON en vez de Evolucion_JSON para que RSP y get_historial lo lean igual
    sheet.appendRow(['Timestamp','DNI','Paciente','Region','Vals_JSON']);
  }
  // Si la hoja ya existe con el header viejo, agregar Vals_JSON si falta
  const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  if(!headers.includes('DNI')){
    // Hoja vieja sin DNI — la reconvertimos (solo headers, no borra datos)
    sheet.getRange(1,1,1,5).setValues([['Timestamp','DNI','Paciente','Region','Vals_JSON']]);
  }
  sheet.appendRow([
    new Date().toISOString(),
    d.dni||'',
    d.paciente||'',
    d.region||'',
    JSON.stringify(d.vals||{})
  ]);
  return{success:true,message:'Evolución de kinesiología guardada'};
}
function guardarEvaluacionPaciente(d){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName('Evaluaciones_Pacientes');
  if(!sheet){
    sheet=ss.insertSheet('Evaluaciones_Pacientes');
    sheet.appendRow(['Timestamp','DNI','Paciente','Profesional','Region','Valores_JSON']);
  }
  sheet.appendRow([new Date().toISOString(),d.dni||'',d.paciente||'',d.profesional||'',d.region||'',JSON.stringify(d.vals||{})]);
  return{success:true,message:'Evaluación del paciente guardada'};
}
function getHistorialPaciente(dni){
  if(!dni) return{success:true,historial:[],evaluaciones:[],kinesiologia:[]};
  const ss=SpreadsheetApp.getActiveSpreadsheet();

  // ID del paciente, para encontrar también filas viejas que no tienen DNI propio
  let idPaciente=null;
  const pSheet=ss.getSheetByName('Pacientes');
  if(pSheet){
    const pData=pSheet.getDataRange().getValues();
    const pHeaders=pData[0];
    const dniCol=pHeaders.indexOf('DNI');
    const idCol=pHeaders.indexOf('ID');
    const paciente=pData.slice(1).find(row=>String(row[dniCol])===String(dni));
    if(paciente) idPaciente=paciente[idCol];
  }

  // 1) Historia clínica: por DNI directo (filas nuevas) o por ID_Paciente (filas viejas, sin DNI)
  let historial=[];
  const hSheet=ss.getSheetByName('Historia_Clinica');
  if(hSheet){
    const hData=hSheet.getDataRange().getValues();
    const hHeaders=hData[0];
    const dniColH=hHeaders.indexOf('DNI');
    const idPacCol=hHeaders.indexOf('ID_Paciente');
    historial=hData.slice(1).filter(row=>
      (dniColH>-1 && String(row[dniColH])===String(dni)) ||
      (idPaciente!==null && String(row[idPacCol])===String(idPaciente))
    ).map(row=>{const obj={};hHeaders.forEach((h,i)=>{obj[h]=row[i];});return obj;});
  }

  // 2) Evaluaciones que completa el paciente (médico o kinesiólogo) — vía DNI directo
  const evaluaciones = leerPorDni_('Evaluaciones_Pacientes', dni);

  // 3) Evoluciones que carga el kinesiólogo en KSM — vía DNI directo
  const kinesiologia = leerPorDni_('Kinesiologia', dni);

  return{success:true,historial:historial,evaluaciones:evaluaciones,kinesiologia:kinesiologia};
}
// Helper: lee una hoja que tiene columna DNI y devuelve las filas de ese paciente.
// Normaliza Evolucion_JSON → vals (para compatibilidad con filas guardadas antes del fix de header)
function leerPorDni_(nombreHoja, dni){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName(nombreHoja);
  if(!sheet) return[];
  const data=sheet.getDataRange().getValues();
  if(data.length<2) return[];
  const headers=data[0];
  const dniCol=headers.indexOf('DNI');
  if(dniCol===-1) return[];
  return data.slice(1).filter(row=>String(row[dniCol])===String(dni))
    .map(row=>{
      const obj={};
      headers.forEach((h,i)=>{obj[h]=row[i];});
      // Normalizar: si viene Evolucion_JSON pero no Vals_JSON, exponerlo como Valores_JSON también
      if(obj.Evolucion_JSON && !obj.Valores_JSON) obj.Valores_JSON=obj.Evolucion_JSON;
      if(obj.Vals_JSON && !obj.Valores_JSON) obj.Valores_JSON=obj.Vals_JSON;
      return obj;
    });
}

// ══════════════════════════════════════════════════════════════════
// PAGOS DE MERCADO PAGO — verificación real contra la API de MP
// (Apps Script no puede leer el header x-signature de la notificación,
// así que NO confiamos en el contenido del webhook: lo usamos solo como
// aviso de "revisá el pago X", y le preguntamos a Mercado Pago directo,
// con nuestro propio Access Token, si ese pago es real y está aprobado)
// ══════════════════════════════════════════════════════════════════
function confirmarPagoMP(paymentId){
  try{
    const token = PropertiesService.getScriptProperties().getProperty('MP_ACCESS_TOKEN');
    if(!token) return{success:false,error:'Falta configurar el Access Token de Mercado Pago — ejecutar configurarTokenMP() una vez desde el editor'};
    const resp = UrlFetchApp.fetch('https://api.mercadopago.com/v1/payments/'+encodeURIComponent(paymentId), {
      method: 'get',
      headers: {'Authorization': 'Bearer '+token},
      muteHttpExceptions: true
    });
    const pago = JSON.parse(resp.getContentText());
    if(!pago || !pago.id) return{success:false,error:'Mercado Pago no devolvió datos para ese pago'};
    const ref = parsearExternalReference_(pago.external_reference||'');
    guardarPagoConfirmado({
      payment_id: String(pago.id),
      dni: ref.dni,
      tipo: ref.tipo,
      monto: pago.transaction_amount||0,
      estado: pago.status||''
    });
    // NUEVO: si era una cuota de la Membresía de seguimiento y quedó
    // aprobada, activamos el mes automáticamente — así el paciente no
    // depende de que el médico entre a revisar pagos pendientes. Ver
    // activarMembresiaSeguimientoPorPago_ más abajo (sección Membresía
    // de Seguimiento).
    if(ref.tipo==='membresia_seguimiento' && pago.status==='approved'){
      activarMembresiaSeguimientoPorPago_(ref.dni, String(pago.id));
    }
    return{success:true, estado:pago.status, confirmado: pago.status==='approved'};
  }catch(e){
    return{success:false, error:'Error confirmando el pago: '+e.message};
  }
}
// Convención a futuro para RCM: el link de pago debe incluir external_reference
// con el formato "DNI|TIPO_DE_CONSULTA" (ej: "27444812|video15") para poder
// saber de quién es el pago y qué consulta corresponde. Hoy RCM todavía no
// lo manda — falta verlo y ajustarlo para que esto quede conectado de punta a punta.
// ══════════════════════════════════════════════════════════════════
// RCM — genera el link de pago real de Mercado Pago (Checkout Pro) para
// mandarle al paciente. Hasta ahora el sistema solo podía CONFIRMAR un
// pago ya hecho (confirmarPagoMP arriba) pero no existía nada que creara
// el link en sí — por eso a los pacientes nunca les aparecía la opción
// de pagar: el médico tenía que generarlo a mano afuera del sistema.
// external_reference queda como "DNI|tipo" para que, cuando MP avise el
// pago por webhook, parsearExternalReference_ lo pueda relacionar con el
// paciente correcto.
// ══════════════════════════════════════════════════════════════════
function crearLinkPagoMP(d){
  const token = (PropertiesService.getScriptProperties().getProperty('MP_ACCESS_TOKEN')||'').trim();
  if(!token) return {success:false, error:'Falta configurar el Access Token de Mercado Pago — ejecutar configurarTokenMP() una vez, o cargar MP_ACCESS_TOKEN en Propiedades del script'};
  const monto = parseFloat(d.monto);
  if(!monto || monto<=0) return {success:false, error:'Falta indicar el monto a cobrar'};
  const concepto = d.concepto || 'Consulta médica — Dr. Muratori';
  const urlBase = baseSimm_();
  const payload = {
    items: [{ title: concepto, quantity: 1, unit_price: monto, currency_id: 'ARS' }],
    back_urls: { success: urlBase, failure: urlBase, pending: urlBase },
    auto_return: 'approved',
    notification_url: ScriptApp.getService().getUrl(),
    external_reference: (d.dni||'') + '|' + (d.tipo||'teleconsulta'),
    statement_descriptor: 'DR MURATORI'
  };
  if(d.mail){ payload.payer = { email: d.mail }; }
  try{
    const resp = UrlFetchApp.fetch('https://api.mercadopago.com/checkout/preferences', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    const raw = resp.getContentText();
    if(!raw){
      return {success:false, error:'Mercado Pago respondió vacío (código HTTP '+code+'). Puede ser un problema con el Access Token guardado en Propiedades del script — revisar que esté completo (formato APP_USR-...), sin espacios ni saltos de línea de más.'};
    }
    let data;
    try{ data = JSON.parse(raw); }catch(eParse){
      return {success:false, error:'Mercado Pago devolvió una respuesta no reconocible (código HTTP '+code+'): ' + raw.substring(0,300)};
    }
    if(!data.init_point){
      return {success:false, error: 'Código HTTP '+code+' — ' + (data.message || 'Mercado Pago no devolvió el link de pago') + (data.cause ? ' — '+JSON.stringify(data.cause) : '') };
    }
    return {success:true, link: data.init_point, preference_id: data.id};
  }catch(err){
    return {success:false, error:'No se pudo conectar con Mercado Pago: ' + err};
  }
}
// ══════════════════════════════════════════════════════════════════
// DIAGNÓSTICO — para correr UNA VEZ manualmente desde el editor (▶ Ejecutar)
// cuando crearLinkPagoMP da error 403. Prueba el mismo Access Token contra
// un endpoint simple (GET /users/me) para aislar si el bloqueo es del
// token/cuenta en general, o específico de crear preferencias de pago.
// Manda el resultado por mail, igual que testGeminiDirecto().
// ══════════════════════════════════════════════════════════════════
// Versión ampliada (agosto 2026) — Mercado Pago Soporte (ticket WCS-48270)
// pidió puntualmente: la traza HTTP completa de una llamada a GET
// /users/me (headers, URL y método, con el token oculto), la respuesta
// completa incluyendo HEADERS (no solo el body), confirmación de que el
// token se manda en el header Authorization: Bearer (no por URL), y si hay
// una lista de IPs permitidas configurada en la app. Esta función arma esa
// traza para las dos llamadas y la manda por mail, lista para pegar en la
// respuesta del ticket. Lo único que Apps Script NO puede confirmar es la
// lista de IPs — eso se revisa a mano en developers.mercadopago.com, dentro
// de la app "S-I-M-M" → Configuración de seguridad.
function diagnosticarTokenMP(){
  const token = (PropertiesService.getScriptProperties().getProperty('MP_ACCESS_TOKEN')||'').trim();
  if(!token){
    MailApp.sendEmail('lmuratori@gmail.com', 'Test MP — ERROR', 'Falta configurar MP_ACCESS_TOKEN en Propiedades del script.');
    return;
  }
  const tokenOculto = token.substring(0,12) + '...' + token.substring(token.length-4);
  let cuerpo =
    'Access Token usado (primeros/últimos caracteres): ' + tokenOculto + '\n' +
    'Confirmado: el token se envía SIEMPRE en el header "Authorization: Bearer <token>", nunca como parámetro en la URL (ver el código de crearLinkPagoMP/doPost, sección Mercado Pago).\n\n';

  // ── GET /users/me ──
  try{
    const url = 'https://api.mercadopago.com/users/me';
    const headersEnviados = { 'Authorization': 'Bearer ' + tokenOculto };
    const resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    const codigo = resp.getResponseCode();
    const texto = resp.getContentText();
    let headersResp = {};
    try{ headersResp = resp.getAllHeaders(); }catch(eh){ headersResp = {error: String(eh)}; }
    cuerpo += '=== GET /users/me ===\n' +
      'cURL equivalente (token oculto):\n' +
      'curl -X GET "' + url + '" -H "Authorization: Bearer ' + tokenOculto + '"\n\n' +
      'Código HTTP: ' + codigo + '\n' +
      'Headers de la respuesta:\n' + JSON.stringify(headersResp, null, 2) + '\n' +
      'Body de la respuesta: ' + (texto || '(vacía)') + '\n\n';
  }catch(e){
    cuerpo += '=== GET /users/me ===\nNo se pudo ni llamar: ' + e + '\n\n';
  }

  // ── POST /checkout/preferences ──
  try{
    const url2 = 'https://api.mercadopago.com/checkout/preferences';
    const payload2 = JSON.stringify({ items: [{ title: 'Prueba diagnóstico', quantity: 1, unit_price: 100, currency_id: 'ARS' }] });
    const resp2 = UrlFetchApp.fetch(url2, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: payload2,
      muteHttpExceptions: true
    });
    const codigo2 = resp2.getResponseCode();
    const texto2 = resp2.getContentText();
    let headersResp2 = {};
    try{ headersResp2 = resp2.getAllHeaders(); }catch(eh){ headersResp2 = {error: String(eh)}; }
    cuerpo += '=== POST /checkout/preferences (prueba mínima) ===\n' +
      'cURL equivalente (token oculto):\n' +
      'curl -X POST "' + url2 + '" -H "Authorization: Bearer ' + tokenOculto + '" -H "Content-Type: application/json" -d \'' + payload2 + '\'\n\n' +
      'Código HTTP: ' + codigo2 + '\n' +
      'Headers de la respuesta:\n' + JSON.stringify(headersResp2, null, 2) + '\n' +
      'Body de la respuesta: ' + (texto2 || '(vacía)') + '\n\n';
  }catch(e){
    cuerpo += '=== POST /checkout/preferences ===\nNo se pudo ni llamar: ' + e + '\n\n';
  }

  cuerpo +=
    '── Pendiente de revisar A MANO (esto Apps Script no lo puede confirmar) ──\n' +
    'Mercado Pago Soporte también preguntó si hay una lista de IPs permitidas configurada en la aplicación. ' +
    'Revisar en developers.mercadopago.com → la app "S-I-M-M" → Configuración de seguridad. ' +
    'Si hay una lista cargada, la IP desde la que corre Google Apps Script cambia y no es fija — eso solo explicaría el 403 vacío en TODAS las llamadas, sin importar el endpoint, que es exactamente lo que está pasando.';

  MailApp.sendEmail('lmuratori@gmail.com', 'Test MP — diagnóstico Access Token (ampliado, para WCS-48270)', cuerpo);
}
function parsearExternalReference_(ref){
  const partes = String(ref).split('|');
  return{dni: partes[0]||'', tipo: partes[1]||''};
}
function guardarPagoConfirmado(info){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName('Pagos_Confirmados');
  if(!sheet){
    sheet=ss.insertSheet('Pagos_Confirmados');
    sheet.appendRow(['Timestamp','Payment_ID','DNI','Tipo_Consulta','Monto','Estado_MP','Atendido']);
  }
  // Si Mercado Pago reintenta la misma notificación, actualizamos en vez de duplicar
  const data=sheet.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][1])===String(info.payment_id)){
      sheet.getRange(i+1,6).setValue(info.estado);
      return;
    }
  }
  sheet.appendRow([new Date().toISOString(), info.payment_id, info.dni, info.tipo, info.monto, info.estado, 'NO']);
}
function getPagosPendientes(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName('Pagos_Confirmados');
  if(!sheet) return{success:true,pendientes:[]};
  const data=sheet.getDataRange().getValues();
  if(data.length<2) return{success:true,pendientes:[]};
  const headers=data[0];
  const pendientes=data.slice(1).map((row,i)=>{
    const obj={_row:i+2};
    headers.forEach((h,j)=>{obj[h.toLowerCase()]=row[j];});
    return obj;
  }).filter(r=>r.estado_mp==='approved' && String(r.atendido).toUpperCase()!=='SI');
  return{success:true,pendientes:pendientes};
}
function marcarPagoAtendido(paymentId){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName('Pagos_Confirmados');
  if(!sheet||!paymentId) return{success:false,error:'Falta hoja o Payment_ID'};
  const data=sheet.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][1])===String(paymentId)){
      sheet.getRange(i+1,7).setValue('SI');
      return{success:true};
    }
  }
  return{success:false,error:'No se encontró ese Payment_ID'};
}

// ══════════════════════════════════════════════════════════════════
// MEMBRESÍA DE SEGUIMIENTO — pensada para UNA patología puntual, con
// renovación mensual DE MUTUO ACUERDO entre médico y paciente (no es un
// cobro automático recurrente: cada mes se vuelve a decidir con el
// paciente si sigue). Incluye la receta/orden necesaria para ESE motivo
// puntual — nunca recetas a otras personas ni medicación fuera de la
// especialidad en tratamiento — y hasta 3 mensajes por semana (máx. 400
// palabras cada uno, siempre sobre el estado de la patología en
// seguimiento) para agilizar el control online. Los mensajes no usados
// en una semana NO se acumulan para la siguiente ("se usa o se pierde").
// Ver procesarBandejaEntrada() más abajo para el chequeo de tope/tema.
//
// Es un mecanismo TOTALMENTE SEPARADO del campo "Membresia" SI/NO de la
// hoja Pacientes (esa es gratuita, la da el médico a su criterio, y
// habilita RMYT/mail con contexto clínico/agendas/RSPQ) — no confundir
// ni fusionar ambos. Reutiliza la infraestructura de Mercado Pago que ya
// existía (crearLinkPagoMP/confirmarPagoMP) con tipo 'membresia_seguimiento'
// en vez de armar un circuito de cobro paralelo.
// ══════════════════════════════════════════════════════════════════
// ══════════════════════════════════════════════════════════════════
// MEMBRESÍA ÚNICA (oct 2026) — reglas definidas por el Dr. Muratori:
//  · Una sola membresía, paga: equivale al costo de UNA consulta particular.
//  · Dura 4 semanas (28 días). Al terminar NO se renueva sola: se continúa
//    solo si ambas partes están de acuerdo, y solo por esa patología.
//  · Control y seguimiento de UNA patología y de ESA persona — NO es
//    "médico en línea". Un mensaje por vez, con 8 hs mínimas entre mensajes
//    cubiertos; el médico responde dentro de 48 hs hábiles (lunes a viernes).
//  · Si el mensaje es por otra patología u otra persona, el sistema manda
//    un recordatorio automático (texto fijo) que deriva a Primera Consulta.
// ══════════════════════════════════════════════════════════════════
const MEMB_DURACION_DIAS = 28;
const MEMB_INTERVALO_MENSAJES_HS = 8;
const MEMB_RESPUESTA_HS = 48;
const MEMB_AVISO_ENFRIAMIENTO_HS = 12;      // máx. 1 aviso automático cada 12 hs por paciente
function urlPrimeraConsultaOnline_(){ return baseSimm_() + 'QR_Consulta_Virtual.html'; }

function sumarDiasMemb_(fecha, dias){ return new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate()+dias); }
// Normaliza lo que devuelve la planilla (Date o texto) a 'yyyy-MM-dd'
function normalizarFechaMemb_(v){
  if(!v) return '';
  if(v instanceof Date){
    const p = n => String(n).padStart(2,'0');
    return v.getFullYear()+'-'+p(v.getMonth()+1)+'-'+p(v.getDate());
  }
  return String(v).slice(0,10);
}
// Vencida = hoy es posterior al último día de vigencia. Sin fecha → no vence (membresías anteriores).
function membVencida_(fechaProx, hoy){
  const s = normalizarFechaMemb_(fechaProx);
  if(!s) return false;
  return normalizarFechaMemb_(hoy || new Date()) > s;
}
function membHorasDesde_(iso, ahora){
  if(!iso) return null;
  const t = new Date(iso).getTime();
  if(isNaN(t)) return null;
  return ((ahora||new Date()).getTime() - t) / 3600000;
}
// Fecha límite para responder: cuenta SOLO lunes a viernes. Si el mensaje llega
// sábado o domingo, el reloj arranca el lunes a las 00:00.
function responderAntesDe_(recibido, horas){
  let t = new Date(recibido.getTime());
  if(t.getDay()===6) t = new Date(t.getFullYear(), t.getMonth(), t.getDate()+2);
  else if(t.getDay()===0) t = new Date(t.getFullYear(), t.getMonth(), t.getDate()+1);
  let restante = (horas||MEMB_RESPUESTA_HS) * 3600000;
  while(restante > 0){
    const finDia = new Date(t.getFullYear(), t.getMonth(), t.getDate()+1);
    const disp = finDia.getTime() - t.getTime();
    if(restante <= disp){ t = new Date(t.getTime()+restante); restante = 0; }
    else {
      restante -= disp; t = finDia;
      while(t.getDay()===6 || t.getDay()===0) t = new Date(t.getFullYear(), t.getMonth(), t.getDate()+1);
    }
  }
  return t;
}
function textoAvisoFueraDeMembresia_(nombre, patologia, otraPersona){
  return 'Hola ' + (nombre||'') + ',\n\n' +
    'Recibimos tu mensaje. Te recordamos que tu membresía es para el control y seguimiento de UNA sola patología — «' + (patologia||'la acordada') + '» — y es exclusiva para vos. No es un servicio de médico en línea.\n\n' +
    (otraPersona
      ? 'Este mensaje parece referirse a OTRA persona, por lo que no queda cubierto por la membresía.\n\n'
      : 'Este mensaje parece referirse a otro motivo de consulta, por lo que no queda cubierto por la membresía.\n\n') +
    'Si necesitás consultar por otro motivo o por otra persona, te recomiendo usar la Primera Consulta Online (se activa con el QR, o desde este link: ' + urlPrimeraConsultaOnline_() + ') o la sección de Primera Consulta por mail: ambas son equivalentes.\n\n' +
    'Si es una urgencia, concurrí a una guardia médica.\n\n' +
    'Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943';
}
// Manda el recordatorio automático (texto fijo, sin IA) y deja registrada la hora
// para no repetirlo más de 1 vez cada MEMB_AVISO_ENFRIAMIENTO_HS por paciente.
function enviarAvisoFueraDeMembresia_(remitente, dni, nombre, patologia, otraPersona){
  try{
    const to = extraerEmailDe_(remitente);
    if(!to || to.indexOf('@')<1) return false;
    if(/no-?reply|donotreply|mailer-daemon|postmaster|notifications?@/i.test(to)) return false; // evita responderle a robots
    const sheet = hojaMembresiasSeguimiento_();
    const fila = filaMembresiaPorDni_(sheet, dni);
    const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    const colAviso = headers.indexOf('Ultimo_Aviso_Fuera_Tema')+1;
    if(fila!==-1 && colAviso>0){
      const prev = sheet.getRange(fila, colAviso).getValue();
      const h = membHorasDesde_(prev ? (prev instanceof Date ? prev.toISOString() : String(prev)) : '', new Date());
      if(h!==null && h < MEMB_AVISO_ENFRIAMIENTO_HS) return false;
    }
    MailApp.sendEmail(to, 'Seguimiento de tu patología — recordatorio', textoAvisoFueraDeMembresia_(nombre, patologia, otraPersona), {
      replyTo:'drluismuratori@gmail.com', name:'Dr. Luis Alberto Muratori'
    });
    if(fila!==-1 && colAviso>0) sheet.getRange(fila, colAviso).setValue(new Date().toISOString());
    return true;
  }catch(e){ Logger.log('enviarAvisoFueraDeMembresia_: ' + e); return false; }
}
// Lectura SIN efectos secundarios: ¿tiene una membresía paga vigente hoy?
function membresiaSeguimientoVigente_(dni){
  try{
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Membresias_Seguimiento');
    if(!sheet) return false;
    const data = sheet.getDataRange().getValues(), h = data[0];
    const cD=h.indexOf('DNI'), cE=h.indexOf('Estado'), cF=h.indexOf('Fecha_Proxima_Renovacion');
    for(let i=1;i<data.length;i++){
      if(String(data[i][cD])===String(dni)) return data[i][cE]==='activa' && !membVencida_(data[i][cF], new Date());
    }
  }catch(e){}
  return false;
}
function registrarMensajeCubiertoMembresia_(dni){
  const sheet = hojaMembresiasSeguimiento_();
  const fila = filaMembresiaPorDni_(sheet, dni);
  if(fila===-1) return false;
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const col = headers.indexOf('Ultimo_Mensaje_Cubierto')+1;
  if(col>0) sheet.getRange(fila, col).setValue(new Date().toISOString());
  return true;
}
// El médico activa la membresía SIN cobro (cortesía), por 4 semanas y por UNA patología.
function activarCortesiaMembresia(d){
  const dni = String(d.DNI || d.dni || '');
  if(!dni) return {success:false, error:'Falta DNI del paciente'};
  const patologia = String(d.Patologia || d.patologia || '').trim();
  if(!patologia) return {success:false, error:'Falta indicar la patología en seguimiento'};
  const sheet = hojaMembresiasSeguimiento_();
  let fila = filaMembresiaPorDni_(sheet, dni);
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  if(fila===-1){
    sheet.appendRow([dni, d.Apellido||'', d.Nombre||'', patologia, 'esperando_pago', 0, '', '', lunesDeLaSemana_(), 0, '']);
  } else {
    sheet.getRange(fila, headers.indexOf('Apellido')+1).setValue(d.Apellido||'');
    sheet.getRange(fila, headers.indexOf('Nombre')+1).setValue(d.Nombre||'');
    sheet.getRange(fila, headers.indexOf('Patologia')+1).setValue(patologia);
    sheet.getRange(fila, headers.indexOf('Precio_Mensual')+1).setValue(0);
  }
  activarMembresiaSeguimientoPorPago_(dni, '');
  const r = getMembresiaSeguimiento(dni);
  return {success:true, message:'Membresía activada sin cobro por '+MEMB_DURACION_DIAS+' días', vence: r && r.fecha_proxima_renovacion};
}

function hojaMembresiasSeguimiento_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Membresias_Seguimiento');
  if(!sheet){
    sheet = ss.insertSheet('Membresias_Seguimiento');
    const headers = ['DNI','Apellido','Nombre','Patologia','Estado','Precio_Mensual',
                      'Fecha_Activacion','Fecha_Proxima_Renovacion','Semana_Actual',
                      'Mensajes_Usados_Semana','Ultimo_Payment_ID_Aplicado',
                      'Ultimo_Mensaje_Cubierto','Ultimo_Aviso_Fuera_Tema'];
    sheet.getRange(1,1,1,headers.length).setValues([headers])
      .setBackground('#7C2D92').setFontColor('white').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  // Hojas creadas antes de oct 2026: agregar las columnas nuevas por nombre
  const hdrs = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  ['Ultimo_Mensaje_Cubierto','Ultimo_Aviso_Fuera_Tema'].forEach(n=>{
    if(hdrs.indexOf(n)===-1){ sheet.getRange(1, sheet.getLastColumn()+1).setValue(n).setBackground('#7C2D92').setFontColor('white').setFontWeight('bold'); }
  });
  return sheet;
}
// Fecha (AAAA-MM-DD) del lunes de la semana a la que pertenece "fecha"
// (o hoy, si no se pasa nada). Se usa para el tope "usalo o perdelo" de 3
// mensajes/semana — no hace falta un trigger aparte: cada lectura
// compara la semana guardada contra la de hoy y resetea sola si cambió.
function lunesDeLaSemana_(fecha){
  const d = fecha ? new Date(fecha) : new Date();
  const dia = d.getDay(); // 0=domingo..6=sábado
  const diff = (dia===0 ? -6 : 1-dia); // lleva cualquier día al lunes de esa semana
  const lunes = new Date(d.getFullYear(), d.getMonth(), d.getDate()+diff);
  return Utilities.formatDate(lunes, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function filaMembresiaPorDni_(sheet, dni){
  const data = sheet.getDataRange().getValues();
  const dniStr = String(dni);
  for(let i=1;i<data.length;i++){
    if(String(data[i][0])===dniStr) return i+1; // fila real de la planilla (1-based, con header)
  }
  return -1;
}
// Devuelve el estado de la membresía de seguimiento de un DNI, ya con el
// reset semanal aplicado si corresponde. Es lo que consulta el botón
// "📆 Membresía de seguimiento" de OCM para pintar el estado.
function getMembresiaSeguimiento(dni){
  if(!dni) return {success:false, error:'Falta DNI'};
  const sheet = hojaMembresiasSeguimiento_();
  const fila = filaMembresiaPorDni_(sheet, dni);
  if(fila===-1) return {success:true, existe:false};
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const row = sheet.getRange(fila,1,1,headers.length).getValues()[0];
  const obj = {}; headers.forEach((h,i)=>obj[h]=row[i]);
  const lunesHoy = lunesDeLaSemana_();
  if(String(obj.Semana_Actual)!==lunesHoy){
    sheet.getRange(fila, headers.indexOf('Semana_Actual')+1).setValue(lunesHoy);
    sheet.getRange(fila, headers.indexOf('Mensajes_Usados_Semana')+1).setValue(0);
    obj.Semana_Actual = lunesHoy; obj.Mensajes_Usados_Semana = 0;
  }
  const estadoEf = (obj.Estado==='activa' && membVencida_(obj.Fecha_Proxima_Renovacion, new Date())) ? 'vencida' : obj.Estado;
  return {
    success:true, existe:true,
    ultimo_mensaje_cubierto: obj.Ultimo_Mensaje_Cubierto ? (obj.Ultimo_Mensaje_Cubierto instanceof Date ? obj.Ultimo_Mensaje_Cubierto.toISOString() : String(obj.Ultimo_Mensaje_Cubierto)) : '',
    fecha_vencimiento: normalizarFechaMemb_(obj.Fecha_Proxima_Renovacion),
    duracion_dias: MEMB_DURACION_DIAS, intervalo_mensajes_hs: MEMB_INTERVALO_MENSAJES_HS, respuesta_hs: MEMB_RESPUESTA_HS,
    dni: obj.DNI, apellido: obj.Apellido, nombre: obj.Nombre,
    patologia: obj.Patologia, estado: estadoEf, precio_mensual: obj.Precio_Mensual,
    fecha_activacion: obj.Fecha_Activacion, fecha_proxima_renovacion: obj.Fecha_Proxima_Renovacion,
    mensajes_usados_semana: Number(obj.Mensajes_Usados_Semana)||0,
    mensajes_disponibles_semana: Math.max(0, 3-(Number(obj.Mensajes_Usados_Semana)||0))
  };
}
// Lista de TODAS las membresías de seguimiento (solo lectura, no modifica la hoja):
// la usa la pestaña "💳 Membresías" de "Pacientes en seguimiento". Estado efectivo
// (activa → vencida si ya pasó la fecha) y mensajes usados de la semana actual.
function getMembresiasLista(){
  const sheet = hojaMembresiasSeguimiento_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const col = function(n){ return headers.indexOf(n); };
  const lunesHoy = lunesDeLaSemana_();
  const hoyStr = normalizarFechaMemb_(new Date());
  const items = [];
  for(let i=1;i<data.length;i++){
    const r = data[i]; if(!r[0]) continue;
    const venc = normalizarFechaMemb_(r[col('Fecha_Proxima_Renovacion')]);
    const estado = String(r[col('Estado')]||'');
    const estadoEf = (estado==='activa' && membVencida_(r[col('Fecha_Proxima_Renovacion')], new Date())) ? 'vencida' : estado;
    const semana = normalizarFechaMemb_(r[col('Semana_Actual')]);
    const usados = (semana===lunesHoy) ? (Number(r[col('Mensajes_Usados_Semana')])||0) : 0;
    let dias = null;
    if(venc){ dias = Math.round((new Date(venc+'T12:00:00').getTime() - new Date(hoyStr+'T12:00:00').getTime())/86400000); }
    const ult = r[col('Ultimo_Mensaje_Cubierto')];
    items.push({ dni:String(r[0]), apellido:String(r[col('Apellido')]||''), nombre:String(r[col('Nombre')]||''), patologia:String(r[col('Patologia')]||''),
      estado:estadoEf, precio_mensual:r[col('Precio_Mensual')], fecha_activacion:normalizarFechaMemb_(r[col('Fecha_Activacion')]), vence:venc, dias_restantes:dias,
      mensajes_usados_semana:usados, mensajes_disponibles_semana:Math.max(0,3-usados),
      ultimo_mensaje_cubierto: ult ? (ult instanceof Date ? ult.toISOString() : String(ult)) : '' });
  }
  return {success:true, items:items};
}
// ══════════════════════════════════════════════════════════════════
// PENDIENTES PARA DAR FECHA QUIRÚRGICA (oct 2026, pedido de Luis) — por paciente
// del prequirúrgico, qué falta antes de asignar fecha: autorización de la obra
// social, implantes/material, estudios que certifican el prequirúrgico u otros.
// Una fila por paciente+ítem; estado: pendiente | solicitado | listo | eliminado.
// Se edita desde el celular o la PC (todo con token).
// ══════════════════════════════════════════════════════════════════
function hojaPendientesQx_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Pendientes_Qx');
  if(!sh){
    sh = ss.insertSheet('Pendientes_Qx');
    sh.appendRow(['DNI','Item','Estado','Nota','Actualizado']);
    sh.getRange(1,1,1,5).setBackground('#b45309').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function getPendientesQx(){
  const data = hojaPendientesQx_().getDataRange().getValues();
  const items = [];
  for(let i=1;i<data.length;i++){
    const r = data[i]; if(!r[0] || !r[1]) continue;
    if(String(r[2])==='eliminado') continue;
    items.push({dni:String(r[0]), item:String(r[1]), estado:String(r[2]||'pendiente'), nota:String(r[3]||''), actualizado:String(r[4]||'')});
  }
  return {success:true, items:items};
}
function guardarPendienteQx(d){
  if(!d || !String(d.dni||'').trim()) return {success:false, error:'Falta el DNI'};
  const item = String(d.item||'').trim().slice(0,120);
  if(!item) return {success:false, error:'Falta el nombre del pendiente'};
  const estado = String(d.estado||'pendiente');
  if(['pendiente','solicitado','listo','eliminado'].indexOf(estado)<0) return {success:false, error:'Estado inválido'};
  const lock = LockService.getScriptLock();
  try{
    lock.waitLock(10000);
    const sh = hojaPendientesQx_();
    const data = sh.getDataRange().getValues();
    const fila = [String(d.dni), item, estado, String(d.nota||'').slice(0,300), new Date().toISOString()];
    for(let i=1;i<data.length;i++){
      if(String(data[i][0])===fila[0] && String(data[i][1])===item){
        sh.getRange(i+1,1,1,5).setNumberFormat('@').setValues([fila]);
        return {success:true, actualizado:true};
      }
    }
    sh.getRange(sh.getLastRow()+1,1,1,5).setNumberFormat('@').setValues([fila]);
    return {success:true};
  }catch(e){
    return {success:false, error:'No se pudo guardar: '+e};
  }finally{
    try{ lock.releaseLock(); }catch(_){}
  }
}

// DASHBOARD DE RECOMENDACIONES TERAPÉUTICAS (RMYT/RSP) — solo lectura.
// Un renglón por paciente que recibió las Recomendaciones Médicas Terapéuticas
// y/o cargó autoevaluaciones semanales: última autoevaluación, alerta, tendencia y
// si tiene membresía activa (= respuesta constante); sin membresía el control se
// hace en la consulta médica.
function getRmytDashboard(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const num = function(v){ return (v===''||v===null||v===undefined||isNaN(v)) ? null : Number(v); };
  const iso = function(v){ return v instanceof Date ? v.toISOString() : String(v||''); };
  // Pacientes (nombre, obra social, contacto, diagnóstico principal)
  const pacs = {};
  const sp = ss.getSheetByName('Pacientes');
  if(sp){
    const d = sp.getDataRange().getValues(), h = d[0]||[];
    const c = function(n){ return h.indexOf(n); };
    for(let i=1;i<d.length;i++){
      const dni = String(d[i][c('DNI')]||''); if(!dni) continue;
      pacs[dni] = { apellido:String(d[i][c('Apellido')]||''), nombre:String(d[i][c('Nombre')]||''), os:String(d[i][c('Obra_Social')]||''),
        tel:String(d[i][c('Telefono')]||''), mail:String(d[i][c('Email')]||''), dx:String(d[i][c('Diagnostico_Principal')]||'') };
    }
  }
  // Membresías activas y vigentes
  const memb = {};
  const sm = ss.getSheetByName('Membresias_Seguimiento');
  if(sm){
    const d = sm.getDataRange().getValues(), h = d[0]||[];
    for(let i=1;i<d.length;i++){
      const dni = String(d[i][h.indexOf('DNI')]||''); if(!dni) continue;
      const venc = d[i][h.indexOf('Fecha_Proxima_Renovacion')];
      if(String(d[i][h.indexOf('Estado')])==='activa' && !membVencida_(venc, new Date())) memb[dni] = normalizarFechaMemb_(venc);
    }
  }
  // Recomendaciones enviadas
  const recs = {};
  const sd = ss.getSheetByName('IMDCYMT_Docs');
  if(sd){ const d = sd.getDataRange().getValues(); for(let i=1;i<d.length;i++){ if(d[i][1]) recs[String(d[i][1])] = iso(d[i][0]); } }
  // Autoevaluaciones semanales
  const evs = {};
  const sa = ss.getSheetByName('RMYT_Autoevaluaciones');
  if(sa){
    const d = sa.getDataRange().getValues(), h = d[0]||[];
    const c = function(n){ return h.indexOf(n); };
    for(let i=1;i<d.length;i++){
      const dni = String(d[i][c('DNI')]||''); if(!dni) continue;
      (evs[dni] = evs[dni] || []).push({ fecha:iso(d[i][c('Timestamp')]), segmento:String(d[i][c('Segmento')]||''), eva_reposo:num(d[i][c('EVA_Reposo')]), eva_movimiento:num(d[i][c('EVA_Movimiento')]),
        escala_pct_discapacidad:num(d[i][c('Escala_Pct_Discapacidad')]), incapacidad_pct:num(d[i][c('Incapacidad_Pct')]), recuperacion_pct:num(d[i][c('Recuperacion_Pct')]), cumplimiento_promedio:num(d[i][c('Cumplimiento_Prom')]),
        alerta:String(d[i][c('Alerta')]||'') });
    }
  }
  const dnis = {}; Object.keys(recs).forEach(function(k){ dnis[k]=1; }); Object.keys(evs).forEach(function(k){ dnis[k]=1; });
  const items = Object.keys(dnis).map(function(dni){
    const lista = (evs[dni]||[]).slice().sort(function(a,b){ return new Date(a.fecha)-new Date(b.fecha); }); // cronológico
    const ult = lista.length ? lista[lista.length-1] : null;
    let tend = null; try{ tend = lista.length ? calcularTendenciaRMYT_(lista) : null; }catch(_){ tend = null; }
    const p = pacs[dni] || {};
    return { dni:dni, paciente: ((p.apellido||'')+', '+(p.nombre||'')).replace(/^, |, $/g,'') || ('DNI '+dni), obra_social:p.os||'', telefono:p.tel||'', mail:p.mail||'', diagnostico:p.dx||'',
      recomendaciones_fecha: recs[dni]||'', n_evaluaciones: lista.length, ultima: ult,
      tendencia_estado: tend ? tend.estado : '', tendencia_mensaje: tend ? tend.mensaje : '',
      con_membresia: memb[dni]!==undefined, membresia_vence: memb[dni]||'' };
  });
  return {success:true, items:items};
}
// Suma 1 al contador semanal de mensajes cubiertos por la membresía — la
// llama procesarBandejaEntrada() cuando un mensaje entra dentro del cupo
// y del tema. Aplica el mismo reset semanal que getMembresiaSeguimiento
// por si es el primer mensaje de una semana nueva.
function incrementarMensajeSemanalMembresia_(dni){
  const sheet = hojaMembresiasSeguimiento_();
  const fila = filaMembresiaPorDni_(sheet, dni);
  if(fila===-1) return false;
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const colSemana = headers.indexOf('Semana_Actual'), colUsados = headers.indexOf('Mensajes_Usados_Semana');
  const lunesHoy = lunesDeLaSemana_();
  const semanaGuardada = sheet.getRange(fila, colSemana+1).getValue();
  let usados = 0;
  if(String(semanaGuardada)===lunesHoy){
    usados = Number(sheet.getRange(fila, colUsados+1).getValue())||0;
  } else {
    sheet.getRange(fila, colSemana+1).setValue(lunesHoy);
  }
  sheet.getRange(fila, colUsados+1).setValue(usados+1);
  return true;
}
// El médico, DE MUTUO ACUERDO con el paciente, decide iniciar o renovar
// por un mes más el seguimiento de UNA patología puntual — crea (o
// reutiliza) la fila y devuelve el link de pago de Mercado Pago para
// mandarle al paciente. Reutiliza crearLinkPagoMP con tipo
// 'membresia_seguimiento' para no armar un circuito de cobro paralelo.
// Queda en Estado 'esperando_pago' hasta que el pago se confirme (ver
// activarMembresiaSeguimientoPorPago_, disparado solo o a mano).
function iniciarORenovarMembresiaSeguimiento(d){
  const dni = String(d.DNI || d.dni || '');
  if(!dni) return {success:false, error:'Falta DNI del paciente'};
  const patologia = String(d.Patologia || d.patologia || '').trim();
  if(!patologia) return {success:false, error:'Falta indicar la patología en seguimiento'};
  const precio = parseFloat(d.Precio_Mensual || d.precio || 0);
  if(!precio || precio<=0) return {success:false, error:'Falta el precio mensual'};
  const sheet = hojaMembresiasSeguimiento_();
  let fila = filaMembresiaPorDni_(sheet, dni);
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const lunesHoy = lunesDeLaSemana_();
  if(fila===-1){
    sheet.appendRow([dni, d.Apellido||'', d.Nombre||'', patologia, 'esperando_pago', precio, '', '', lunesHoy, 0, '']);
  } else {
    sheet.getRange(fila, headers.indexOf('Apellido')+1).setValue(d.Apellido||'');
    sheet.getRange(fila, headers.indexOf('Nombre')+1).setValue(d.Nombre||'');
    sheet.getRange(fila, headers.indexOf('Patologia')+1).setValue(patologia);
    sheet.getRange(fila, headers.indexOf('Estado')+1).setValue('esperando_pago');
    sheet.getRange(fila, headers.indexOf('Precio_Mensual')+1).setValue(precio);
  }
  const linkPago = crearLinkPagoMP({
    dni: dni, monto: precio, tipo: 'membresia_seguimiento',
    concepto: 'Membresía de seguimiento — ' + patologia + ' — Dr. Muratori',
    mail: d.Email || d.mail || ''
  });
  if(!linkPago.success) return {success:false, error:'No se pudo generar el link de pago: '+linkPago.error};
  return {success:true, message:'Listo — mandale este link al paciente para confirmar el mes', link: linkPago.link};
}
// El médico corta el seguimiento (decisión propia, o porque el paciente
// no quiere seguir) — deja de contar como membresía activa, sin borrar
// el historial de la fila (patología, precios anteriores).
function desactivarMembresiaSeguimiento(d){
  const dni = String(d.DNI || d.dni || '');
  if(!dni) return {success:false, error:'Falta DNI'};
  const sheet = hojaMembresiasSeguimiento_();
  const fila = filaMembresiaPorDni_(sheet, dni);
  if(fila===-1) return {success:false, error:'Ese paciente no tiene una membresía de seguimiento registrada'};
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  sheet.getRange(fila, headers.indexOf('Estado')+1).setValue('inactiva');
  return {success:true, message:'Membresía de seguimiento desactivada'};
}
// Se llama SOLA desde confirmarPagoMP cuando el pago aprobado era de tipo
// 'membresia_seguimiento' (webhook de Mercado Pago) — y también la puede
// disparar a mano confirmarRenovacionMembresiaSeguimiento más abajo, por
// si el webhook no llegó. Activa la membresía por un mes desde hoy y
// arranca la semana de mensajes en 0. Si se le pasa paymentId y ya se
// había aplicado ESE mismo pago antes, no hace nada (evita reiniciar el
// mes si Mercado Pago reintenta la misma notificación).
function activarMembresiaSeguimientoPorPago_(dni, paymentId){
  if(!dni) return false;
  const sheet = hojaMembresiasSeguimiento_();
  const fila = filaMembresiaPorDni_(sheet, dni);
  if(fila===-1) return false; // pago de otro tipo con el mismo DNI, o fila borrada — no tocar nada
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const colPago = headers.indexOf('Ultimo_Payment_ID_Aplicado');
  if(paymentId && colPago>-1){
    const yaAplicado = String(sheet.getRange(fila, colPago+1).getValue()||'');
    if(yaAplicado === String(paymentId)) return false;
    sheet.getRange(fila, colPago+1).setValue(String(paymentId));
  }
  const hoy = new Date();
  const proximaRenovacion = sumarDiasMemb_(hoy, MEMB_DURACION_DIAS);
  const fmt = f => Utilities.formatDate(f, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sheet.getRange(fila, headers.indexOf('Estado')+1).setValue('activa');
  sheet.getRange(fila, headers.indexOf('Fecha_Activacion')+1).setValue(fmt(hoy));
  sheet.getRange(fila, headers.indexOf('Fecha_Proxima_Renovacion')+1).setValue(fmt(proximaRenovacion));
  sheet.getRange(fila, headers.indexOf('Semana_Actual')+1).setValue(lunesDeLaSemana_());
  sheet.getRange(fila, headers.indexOf('Mensajes_Usados_Semana')+1).setValue(0);
  const colUlt = headers.indexOf('Ultimo_Mensaje_Cubierto');
  if(colUlt>-1) sheet.getRange(fila, colUlt+1).setValue('');
  return true;
}
// El médico aprieta "revisar si pagó" — por si el webhook automático de
// Mercado Pago no llegó a activarla sola. Busca en Pagos_Confirmados un
// pago aprobado de este DNI, tipo 'membresia_seguimiento', que todavía no
// se haya aplicado a esta membresía, y si lo encuentra, activa el mes.
// Nunca confía en nada que no sea el propio registro ya verificado contra
// la API de Mercado Pago (ver confirmarPagoMP) — no mira monto ni fecha
// "a ojo", solo el Payment_ID ya confirmado.
function confirmarRenovacionMembresiaSeguimiento(dni){
  if(!dni) return {success:false, error:'Falta DNI'};
  const sheetMemb = hojaMembresiasSeguimiento_();
  const fila = filaMembresiaPorDni_(sheetMemb, dni);
  if(fila===-1) return {success:false, error:'Ese paciente no tiene una membresía de seguimiento registrada'};
  const headersMemb = sheetMemb.getRange(1,1,1,sheetMemb.getLastColumn()).getValues()[0];
  const ultimoAplicado = String(sheetMemb.getRange(fila, headersMemb.indexOf('Ultimo_Payment_ID_Aplicado')+1).getValue()||'');
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sheetPagos = ss.getSheetByName('Pagos_Confirmados');
  let paymentIdNuevo = '';
  if(sheetPagos){
    const data = sheetPagos.getDataRange().getValues();
    const hPagos = data[0];
    const colDni=hPagos.indexOf('DNI'), colTipo=hPagos.indexOf('Tipo_Consulta'),
          colEstado=hPagos.indexOf('Estado_MP'), colPid=hPagos.indexOf('Payment_ID');
    for(let i=1;i<data.length;i++){
      if(String(data[i][colDni])===String(dni) && data[i][colTipo]==='membresia_seguimiento' && data[i][colEstado]==='approved'){
        const pid = String(data[i][colPid]);
        if(pid !== ultimoAplicado) paymentIdNuevo = pid;
      }
    }
  }
  if(paymentIdNuevo){
    activarMembresiaSeguimientoPorPago_(dni, paymentIdNuevo);
    return {success:true, pagado:true, message:'Pago encontrado — membresía activada por este mes'};
  }
  return {success:true, pagado:false, message:'Todavía no se registró un pago nuevo este mes'};
}

// ══════════════════════════════════════════════════════════════════
// RECOMENDACIONES MEDICAS PACIENTE — Guarda y recupera el documento completo del
// paciente (para que funcione desde el celular, donde no hay localStorage de OCM)
// ══════════════════════════════════════════════════════════════════
function saveRecomendaciones(d){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName('IMDCYMT_Docs');
  if(!sheet){
    sheet=ss.insertSheet('IMDCYMT_Docs');
    sheet.appendRow(['Timestamp','DNI','Datos_JSON']);
    sheet.setFrozenRows(1);
  }
  const datos=d.datos||{};
  const dni=d.dni||datos.paciente&&datos.paciente.dni||'';
  // Actualizar si ya existe una fila para este DNI
  const rows=sheet.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][1])===String(dni)){
      sheet.getRange(i+1,1,1,3).setValues([[new Date().toISOString(),dni,JSON.stringify(datos)]]);
      return{success:true,message:'RECOMENDACIONES MEDICAS actualizado'};
    }
  }
  sheet.appendRow([new Date().toISOString(),dni,JSON.stringify(datos)]);
  return{success:true,message:'RECOMENDACIONES MEDICAS guardado'};
}
function getRecomendaciones(dni){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName('IMDCYMT_Docs');
  if(!sheet||!dni) return{success:false,datos:null};
  const rows=sheet.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][1])===String(dni)){
      try{
        const datos=JSON.parse(rows[i][2]||'{}');
        return{success:true,datos:datos};
      }catch(_){ return{success:false,datos:null}; }
    }
  }
  return{success:true,datos:null};
}
// Guarda el Access Token de Mercado Pago en un lugar seguro (las Propiedades
// del script) que nunca queda en este archivo ni se sube a la MyCloud.
// Pasos: 1) pegá tu token donde dice PEGÁ_AQUÍ  2) seleccioná esta función
// en el desplegable de arriba  3) tocá ▶ Ejecutar  4) borrá el token de aquí
// si querés (ya quedó guardado, no se vuelve a necesitar en el código).
function configurarTokenMP(){
  const token = 'PEGÁ_AQUÍ_TU_ACCESS_TOKEN_DE_MERCADO_PAGO';
  PropertiesService.getScriptProperties().setProperty('MP_ACCESS_TOKEN', token);
  Logger.log('Token guardado correctamente.');
}

// Chequeo del token SIN mostrarlo completo — para comparar contra lo que se ve
// en Mercado Pago (Credenciales de producción) sin necesidad de copiar/pegar
// el token entero en ningún lado. Correr después de configurarTokenMP(), y
// mirar el resultado en Ver → Registros de ejecución (o Ctrl+Enter).
function verificarTokenMPGuardado(){
  const token = PropertiesService.getScriptProperties().getProperty('MP_ACCESS_TOKEN') || '';
  if(!token){ Logger.log('No hay ningún token guardado todavía.'); return; }
  Logger.log('Longitud total: ' + token.length + ' caracteres');
  Logger.log('Empieza con: "' + token.substring(0,10) + '"  (debería ser "APP_USR-" para producción)');
  Logger.log('Termina con: "' + token.substring(token.length-6) + '"');
  Logger.log('¿Tiene espacio o salto de línea de más al principio/final?: ' + (token !== token.trim() ? 'SÍ — este es el problema, hay que sacarlo' : 'No, está limpio'));
}

// ══════════════════════════════════════════════════════════════════
// CONSULTAS ONLINE PAGAS — Cola de consultas + respuesta IA
// ══════════════════════════════════════════════════════════════════
const CONSULTAS_HEADERS = ['ID','DNI','Paciente','Tipo','Precio','Pregunta',
  'Estado','Fecha','Payment_ID','Respuesta_Draft','Respuesta_Final','Fecha_Respuesta'];

function sheetConsultas_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sh=ss.getSheetByName('Consultas_Online');
  if(!sh){
    sh=ss.insertSheet('Consultas_Online');
    sh.getRange(1,1,1,CONSULTAS_HEADERS.length).setValues([CONSULTAS_HEADERS])
      .setBackground('#7c3aed').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function saveConsulta(d){
  const sh=sheetConsultas_();
  const id='CONS-'+new Date().getTime();
  sh.appendRow([id,d.dni||'',d.paciente||'',d.tipo||'texto',d.precio||0,
    d.pregunta||'','pendiente_pago',new Date().toISOString(),'','','','']);
  return{success:true,id:id};
}

function getConsultas(estado){
  const sh=sheetConsultas_();
  const rows=sh.getDataRange().getValues();
  if(rows.length<2) return{success:true,consultas:[]};
  const headers=rows[0];
  let lista=rows.slice(1).map(row=>{
    const obj={};headers.forEach((h,i)=>{obj[h.toLowerCase()]=row[i];});return obj;
  });
  if(estado) lista=lista.filter(c=>c.estado===estado);
  return{success:true,consultas:lista.reverse()};
}

// ── REEMPLAZADA por el patch de peaje de sistema: ver más abajo el
//    bloque "PARCHE — Peaje de uso del sistema" para la versión final,
//    que agrega el registro automático del peaje al cerrar una consulta.

function generarRespuestaIA(d){
  const apiKey=PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if(!apiKey) return{success:false,error:'Falta configurar CLAUDE_API_KEY — ejecutar configurarClaudeAPI() una vez desde el editor'};
  const sistema=`Sos el Dr. Luis Alberto Muratori, especialista en Ortopedia, Traumatología y Cirugía de Columna (M.N. 100.540, M.P. 9943).
Respondé consultas médicas en español argentino, de forma clara, precisa y profesional.
Usá términos comprensibles para el paciente. Si la consulta requiere examen físico, aclaralo.
Si podés dar una respuesta útil sin verlo, hacelo.

Datos del paciente:
- Nombre: ${d.paciente||'No especificado'}
- Diagnóstico: ${d.diagnostico||'No especificado'}
- Medicación actual: ${d.medicacion||'No especificada'}
- Región tratada: ${d.region||'No especificada'}
- Historial relevante: ${d.historial||'Sin datos adicionales'}`;

  try{
    const resp=UrlFetchApp.fetch('https://api.anthropic.com/v1/messages',{
      method:'post',
      headers:{'Content-Type':'application/json','x-api-key':apiKey,'anthropic-version':'2023-06-01'},
      payload:JSON.stringify({model:'claude-sonnet-4-6',max_tokens:1000,
        system:sistema,messages:[{role:'user',content:d.pregunta}]}),
      muteHttpExceptions:true
    });
    const data=JSON.parse(resp.getContentText());
    if(data.content&&data.content[0]&&data.content[0].text){
      // Guardar el draft en la consulta
      if(d.id) actualizarConsulta({id:d.id,respuesta_draft:data.content[0].text,estado:'en_proceso'});
      return{success:true,respuesta:data.content[0].text};
    }
    return{success:false,error:'Respuesta inesperada de la API: '+JSON.stringify(data)};
  }catch(e){
    return{success:false,error:'Error de red: '+e.message};
  }
}

// ── TARIFAS ── Guardadas en una hoja de configuración
function getTarifas(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sh=ss.getSheetByName('Configuracion');
  if(!sh) return{success:true,tarifas:null}; // usa las default del frontend
  const rows=sh.getDataRange().getValues();
  for(const row of rows){
    if(row[0]==='tarifas_consulta'){
      try{ return{success:true,tarifas:JSON.parse(row[1]||'{}')}; }catch(_){}
    }
  }
  return{success:true,tarifas:null};
}

function saveTarifas(d){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sh=ss.getSheetByName('Configuracion');
  if(!sh){
    sh=ss.insertSheet('Configuracion');
    sh.appendRow(['clave','valor']);
    sh.setFrozenRows(1);
  }
  const rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(rows[i][0]==='tarifas_consulta'){
      sh.getRange(i+1,2).setValue(JSON.stringify(d.tarifas));
      return{success:true};
    }
  }
  sh.appendRow(['tarifas_consulta',JSON.stringify(d.tarifas)]);
  return{success:true};
}

// ── Configurar la clave de Claude (ejecutar una sola vez desde el editor) ──
function configurarClaudeAPI(){
  const key='PEGÁ_AQUÍ_TU_CLAUDE_API_KEY';
  PropertiesService.getScriptProperties().setProperty('CLAUDE_API_KEY',key);
  Logger.log('Claude API key guardada correctamente.');
}

// ═══════════════════════════════════════════════════════════════════
// NUEVO — Puente con CAM (Consulta Asistida Muratori, ex-Portada Excel)
// Agregado para integrar el razonamiento clinico (CIE-10, segmento,
// banderas rojas, EVA, etc.) sin romper nada de lo que ya funcionaba.
// ═══════════════════════════════════════════════════════════════════

// Guarda una entrada de Historia_Clinica con TODAS las columnas nuevas.
// Si una columna no existe todavia en la hoja, la crea sola (no hay que
// tocar la hoja a mano). Los campos viejos (Diagnostico, Medicamentos,
// etc.) siguen funcionando igual que con addHistoria().
function addHistoriaExtendida(d){
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sheet=ss.getSheetByName('Historia_Clinica');
  let headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];

  // Columnas nuevas que puede traer CAM. Si falta alguna en la hoja, se agrega sola.
  const COLUMNAS_NUEVAS = [
    'Centro_Atencion','Medico_Tratante','Derivado_De','Area_Derivacion','Motivo_Eleccion',
    'Consentimiento_Datos','Consentimiento_Fecha',
    'Circunstancia_Tipo','Circunstancia_Cobertura',
    'Codigo_CIE10','Segmento_Corporal','Tipo_Patologia','Lateralidad',
    'Antecedentes_JSON','Antecedentes_Quirurgicos','Examen_Fisico',
    'Banderas_Rojas','Alerta_Urgente',
    'Peso','Talla','IMC','TA_Sistolica','TA_Diastolica','EVA_Dolor',
    'Farmaco','Dosis','Farmaco_Magistral','Farmacia_Derivada',
    'Kinesico','Kine_Centro','Planilla_Seguimiento','Terapias_Especiales',
    'Autocuidado','Modalidad_Entrenamiento',
    'Estudios_Solicitados',
    'Cirugia_Programada','Tipo_Procedimiento','Fecha_Cirugia_Programada',
    'Kinesiologo_Responsable','Tipo_Actividad_Fisica','Gym_Instructor',
    'Indica_Ortesis','Tipo_Ortesis','Indica_Plantillas','Material_Plantilla',
    'Perfil_Correccion','Discrepancia_Longitud','Lado','Largo_Plantilla',
    'Suplemento_Viscoelastico',
    'Se_Entrega_Certificado','Tipo_Certificado','Reposo_Desde','Reposo_Hasta','Dias_Reposo',
    'Estado_Caso','N_Control','Redaccion_Historia_Clinica'
  ];
  let huboColumnaNueva=false;
  COLUMNAS_NUEVAS.forEach(col=>{
    if(headers.indexOf(col)===-1){
      sheet.getRange(1,headers.length+1).setValue(col);
      headers.push(col);
      huboColumnaNueva=true;
    }
  });
  if(huboColumnaNueva){
    sheet.getRange(1,1,1,headers.length).setBackground('#0F9D58').setFontColor('white').setFontWeight('bold');
  }

  const id='HC-'+String(sheet.getLastRow()).padStart(4,'0');
  const fila=new Array(headers.length).fill('');
  const set=(nombreCol,valor)=>{const i=headers.indexOf(nombreCol);if(i>-1)fila[i]=valor||'';};

  // Campos clasicos (compatibilidad con lo que ya lee OCM/RSP)
  set('ID_HC', id);
  set('ID_Paciente', d.ID_Paciente||'');
  set('Fecha', new Date().toLocaleDateString('es-AR'));
  set('Diagnostico', d.Diagnostico||'');
  set('Medicamentos', d.Medicamentos||d.Farmaco||'');
  set('Estudios_Realizados', d.Estudios_Realizados||'');
  set('Estudios_Pendientes', d.Estudios_Solicitados||'');
  set('Indicaciones', d.Indicaciones||'');
  set('Observaciones', d.Observaciones||'');
  set('Proximo_Control', d.Proximo_Control||'');
  set('DNI', d.DNI||'');

  // Campos nuevos (razonamiento clinico de CAM)
  set('Centro_Atencion', d.Centro_Atencion);
  set('Medico_Tratante', d.Medico_Tratante);
  set('Derivado_De', d.Derivado_De);
  set('Area_Derivacion', d.Area_Derivacion);
  set('Motivo_Eleccion', d.Motivo_Eleccion);
  set('Consentimiento_Datos', d.Consentimiento_Datos);
  set('Consentimiento_Fecha', d.Consentimiento_Fecha);
  set('Circunstancia_Tipo', d.Circunstancia_Tipo);
  set('Circunstancia_Cobertura', d.Circunstancia_Cobertura);
  set('Codigo_CIE10', d.Codigo_CIE10);
  set('Segmento_Corporal', d.Segmento_Corporal);
  set('Tipo_Patologia', d.Tipo_Patologia);
  set('Lateralidad', d.Lateralidad);
  set('Antecedentes_JSON', d.Antecedentes_JSON ? JSON.stringify(d.Antecedentes_JSON) : '');
  set('Antecedentes_Quirurgicos', d.Antecedentes_Quirurgicos);
  set('Examen_Fisico', d.Examen_Fisico);
  set('Banderas_Rojas', d.Banderas_Rojas);
  set('Alerta_Urgente', d.Alerta_Urgente);
  set('Peso', d.Peso);
  set('Talla', d.Talla);
  set('IMC', d.IMC);
  set('TA_Sistolica', d.TA_Sistolica);
  set('TA_Diastolica', d.TA_Diastolica);
  set('EVA_Dolor', d.EVA_Dolor);
  set('Farmaco', d.Farmaco);
  set('Dosis', d.Dosis);
  set('Farmaco_Magistral', d.Farmaco_Magistral);
  set('Farmacia_Derivada', d.Farmacia_Derivada);
  set('Kinesico', d.Kinesico);
  set('Kine_Centro', d.Kine_Centro);
  set('Planilla_Seguimiento', d.Planilla_Seguimiento);
  set('Terapias_Especiales', d.Terapias_Especiales);
  set('Autocuidado', d.Autocuidado);
  set('Modalidad_Entrenamiento', d.Modalidad_Entrenamiento);
  set('Estudios_Solicitados', d.Estudios_Solicitados);
  set('Cirugia_Programada', d.Cirugia_Programada);
  set('Tipo_Procedimiento', d.Tipo_Procedimiento);
  set('Fecha_Cirugia_Programada', d.Fecha_Cirugia_Programada);
  set('Kinesiologo_Responsable', d.Kinesiologo_Responsable);
  set('Tipo_Actividad_Fisica', d.Tipo_Actividad_Fisica);
  set('Gym_Instructor', d.Gym_Instructor);
  set('Indica_Ortesis', d.Indica_Ortesis);
  set('Tipo_Ortesis', d.Tipo_Ortesis);
  set('Indica_Plantillas', d.Indica_Plantillas);
  set('Material_Plantilla', d.Material_Plantilla);
  set('Perfil_Correccion', d.Perfil_Correccion);
  set('Discrepancia_Longitud', d.Discrepancia_Longitud);
  set('Lado', d.Lado);
  set('Largo_Plantilla', d.Largo_Plantilla);
  set('Suplemento_Viscoelastico', d.Suplemento_Viscoelastico);
  set('Se_Entrega_Certificado', d.Se_Entrega_Certificado);
  set('Tipo_Certificado', d.Tipo_Certificado);
  set('Reposo_Desde', d.Reposo_Desde);
  set('Reposo_Hasta', d.Reposo_Hasta);
  set('Dias_Reposo', d.Dias_Reposo);
  set('Estado_Caso', d.Estado_Caso||'Activo');
  set('N_Control', d.N_Control);
  set('Redaccion_Historia_Clinica', d.Redaccion_Historia_Clinica);

  sheet.appendRow(fila);
  return {success:true, id:id};
}

// Envia por MAIL (no WhatsApp, a proposito: el medico eligio mail para no
// recibir una respuesta inmediata del paciente durante la consulta) la
// receta/documentacion + un mini-resumen para que el paciente sepa que
// tiene y que debe hacer. Requiere que el paciente tenga Email cargado.
//
// d = {
//   Email: 'paciente@mail.com',
//   Nombre_Paciente: 'Juan Perez',
//   Asunto: 'Resumen de tu consulta - Dr. Muratori',
//   Mini_Resumen: 'texto en lenguaje simple para el paciente',
//   Receta_Texto: 'texto completo de la receta (o vacio si va como PDF adjunto)',
//   Pdf_Base64: 'opcional, base64 de un PDF ya generado en OCM',
//   Pdf_Nombre: 'opcional, nombre del archivo adjunto'
// }
function enviarRecetaPorMail(d){
  if(!d.Email) return {success:false, error:'El paciente no tiene mail cargado'};
  d.Email = String(d.Email).trim().toLowerCase();
  // Un mail mal escrito (ej. "gmail-com") NO da error en Apps Script: sale y rebota en silencio.
  if(!/^[^@\s]+@[^@\s.]+(\.[^@\s.]+)+$/.test(d.Email)) return {success:false, error:'El mail "'+d.Email+'" no es válido (revisá el punto: gmail.com)'};
  let cuota = null;
  try{ cuota = MailApp.getRemainingDailyQuota(); }catch(_){}
  if(cuota!==null && cuota<=0) return {success:false, error:'Se agotó el cupo diario de mails de Google (se renueva cada 24 hs). Imprimí la receta o entregala en mano.'};

  // Anti-duplicados: si llega EXACTAMENTE el mismo envío (mismo mail, asunto y contenido) en
  // menos de 2 minutos, no se vuelve a mandar (doble clic, reintento de red, etc.).
  try{
    const huella = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,
      d.Email+'|'+(d.Asunto||'')+'|'+(d.Mini_Resumen||'')+'|'+(d.Pdf_Nombre||'')+'|'+(d.Pdf_Base64?d.Pdf_Base64.length:0)+'|'+(d.Pdfs?d.Pdfs.length:0)));
    const cache = CacheService.getScriptCache();
    if(cache.get('mail_'+huella)) return {success:true, duplicado:true, enviado_a:d.Email, message:'Ya se había enviado hace instantes — no se repitió'};
    cache.put('mail_'+huella, '1', 120);
  }catch(_){}

  const cuerpo =
    'Hola ' + (d.Nombre_Paciente||'') + ',\n\n' +
    (d.Intro || 'Este es un resumen de tu consulta de hoy con el Dr. Muratori:') + '\n\n' +
    (d.Mini_Resumen||'') + '\n\n' +
    (d.Receta_Texto ? ('Indicaciones / Receta:\n' + d.Receta_Texto + '\n\n') : '') +
    'Ante cualquier duda, comunicate por los canales habituales del consultorio.\n\n' +
    'Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943';

  const opciones = { replyTo: 'drluismuratori@gmail.com', name: 'Dr. Luis Alberto Muratori' };
  // Soporta dos formatos: un solo PDF (Pdf_Base64/Pdf_Nombre, usado por la
  // Receta médica y la documentación quirúrgica) o varios PDF juntos en un
  // mismo mail (Pdfs: [{base64,nombre}, ...], usado por el envío de
  // Indicaciones Médicas tildadas). Antes solo se soportaba el primero, así
  // que el envío múltiple mandaba el mail SIN ningún PDF adjunto aunque
  // avisara "enviado" — corregido agosto 2026.
  if(d.Pdf_Base64){
    const bytes = Utilities.base64Decode(d.Pdf_Base64);
    const blob = Utilities.newBlob(bytes, 'application/pdf', d.Pdf_Nombre||'documentacion_consulta.pdf');
    opciones.attachments = [blob];
  } else if(d.Pdfs && d.Pdfs.length){
    opciones.attachments = d.Pdfs.map(function(p, i){
      const bytes = Utilities.base64Decode(p.base64);
      return Utilities.newBlob(bytes, 'application/pdf', p.nombre || ('documento_'+(i+1)+'.pdf'));
    });
  }

  try{
    MailApp.sendEmail(d.Email, d.Asunto||'Resumen de tu consulta - Dr. Muratori', cuerpo, opciones);
    return {success:true, enviado_a:d.Email, cuota_restante:(cuota===null?null:cuota-1)};
  }catch(e){
    return {success:false, error:String(e)};
  }
}

// ═══════════════════════════════════════════════════════════════════
// NUEVO — Lector CUD: guarda el registro de discapacidad, lo valida
// cruzado contra Pacientes (agenda familiar) y lo vuelca en Historia_Clinica.
// El Lector CUD manda el campo "accion" (no "action") y envuelve el
// registro dentro de {registro:{...}} - eso ya se resuelve en doPost.
// ═══════════════════════════════════════════════════════════════════
function guardarRegistroCUD(reg){
  if(!reg || !reg.dni) return {success:false, error:'Falta DNI en el registro del CUD'};
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) Guardar el registro completo en su propia hoja
  let dcSheet = ss.getSheetByName('Discapacidad_CUD');
  if(!dcSheet){
    dcSheet = ss.insertSheet('Discapacidad_CUD');
    dcSheet.appendRow(['Timestamp','DNI','Nombre','Telefono','Fecha_Nac','Diagnostico','Codigos_CIF','Tipo',
                        'Subtipo','Modalidad','Orientacion','Fecha_Emision','Fecha_Vencimiento','Junta',
                        'Firmantes_JSON','Efectores','Tipo_Pension','Trazabilidad']);
  }
  dcSheet.appendRow([
    new Date().toISOString(), reg.dni||'', reg.nombre||'', reg.telefono||'', reg.fnac||'',
    reg.diagnostico||'', reg.codigos||'', reg.tipo||'', reg.subtipo||'', reg.modalidad||'',
    reg.orientacion||'', reg.emision||'', reg.vencimiento||'', reg.junta||'',
    JSON.stringify(reg.firmantes||[]), reg.efectores||'', reg.tipoPension||'', reg.trazabilidad||''
  ]);

  // 2) Validacion cruzada con la Agenda familiar (Pacientes): si el DNI no existe
  //    todavia ahi, se crea un registro basico para que quede conectado; si ya
  //    existe, se actualiza el telefono si vino uno nuevo (sin pisar otros datos)
  const pSheet = ss.getSheetByName('Pacientes');
  let pacienteId = '';
  if(pSheet){
    const pData = pSheet.getDataRange().getValues();
    const pHeaders = pData[0];
    const dniCol = pHeaders.indexOf('DNI');
    const idCol = pHeaders.indexOf('ID');
    const telCol = pHeaders.indexOf('Telefono');
    let filaEncontrada = -1;
    for(let i=1;i<pData.length;i++){
      if(String(pData[i][dniCol])===String(reg.dni)){ filaEncontrada=i; break; }
    }
    if(filaEncontrada>-1){
      pacienteId = pData[filaEncontrada][idCol];
      if(reg.telefono && !pData[filaEncontrada][telCol] && telCol>-1){
        pSheet.getRange(filaEncontrada+1, telCol+1).setValue(reg.telefono);
      }
    } else {
      const nuevoId = addPaciente({
        Apellido:'', Nombre:reg.nombre||'', DNI:reg.dni, Fecha_Nac:reg.fnac||'',
        Telefono:reg.telefono||'', Diagnostico_Principal:reg.diagnostico||'',
        Estado_Tratamiento:'Discapacidad (CUD)'
      });
      pacienteId = nuevoId.id;
    }
  }

  // 3) Volcar tambien en Historia_Clinica, para que quede en la misma linea de
  //    tiempo que el resto de las consultas de ese paciente
  const hSheet = ss.getSheetByName('Historia_Clinica');
  if(hSheet){
    const idHc = 'HC-'+String(hSheet.getLastRow()).padStart(4,'0');
    const resumenCud = `Registro CUD: ${reg.tipo||''} ${reg.subtipo?'('+reg.subtipo+')':''}. `+
      `Vigencia: ${reg.emision||'?'} a ${reg.vencimiento||'?'}. Junta: ${reg.junta||'-'}.`;
    hSheet.appendRow([
      idHc, pacienteId||reg.dni, new Date().toLocaleDateString('es-AR'),
      reg.diagnostico||'', '', '', '', resumenCud, '', ''
    ]);
  }

  return {success:true, message:'Registro CUD guardado y conectado con Pacientes e Historia_Clinica', paciente_id: pacienteId};
}

// ══════════════════════════════════════════════════════════════════
// LECTOR CUD — Gemini Vision. El frontend (COL_Lector_CUD.html) manda
// la foto del certificado de discapacidad en base64; esto le pide a
// Gemini que extraiga los datos en el mismo formato que espera
// guardarRegistroCUD, para que el médico solo tenga que revisar y
// confirmar antes de guardar (nunca se guarda automáticamente).
// ══════════════════════════════════════════════════════════════════
function analizarFotoCUD(d){
  if(!d || !d.foto_base64) return {success:false, error:'Falta la foto'};
  const mime = d.mime_type || 'image/jpeg';
  const prompt = 'Sos un asistente que transcribe certificados únicos de discapacidad (CUD) argentinos a partir de una foto. ' +
    'Mirá la imagen y extraé EXACTAMENTE estos campos, en español, tal cual figuran impresos (si un campo no se ve o no figura, dejalo como cadena vacía "", no inventes datos):\n' +
    '- dni (solo números, sin puntos)\n' +
    '- nombre (Apellido, Nombre)\n' +
    '- telefono\n' +
    '- fnac (fecha de nacimiento, formato DD/MM/AAAA)\n' +
    '- diagnostico\n' +
    '- codigos (códigos CIF/CIE si figuran, separados por coma)\n' +
    '- tipo (ej: Motora, Visual, Auditiva, Mental, Visceral, Múltiple)\n' +
    '- subtipo\n' +
    '- modalidad (ej: Permanente, Transitorio)\n' +
    '- orientacion\n' +
    '- emision (fecha de emisión, DD/MM/AAAA)\n' +
    '- vencimiento (fecha de vencimiento, DD/MM/AAAA)\n' +
    '- junta (junta evaluadora / organismo emisor)\n' +
    '- efectores\n' +
    '- tipoPension (si figura pensión no contributiva u otra)\n' +
    '- trazabilidad (número de trámite/expediente si figura)\n\n' +
    'Devolvé SOLO un JSON válido, sin texto adicional ni markdown, con exactamente esas 15 claves (todas como texto).';

  try{
    const resp = UrlFetchApp.fetch(getGeminiUrl_(), {
      method:'post', contentType:'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [
          { text: prompt },
          { inline_data: { mime_type: mime, data: d.foto_base64 } }
        ]}],
        generationConfig: {
          thinkingConfig: { thinkingLevel: 'minimal' }
        }
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(resp.getContentText());
    registrarUsoTokens_('analizarFotoCUD', 'gemini-3.6-flash', data.usageMetadata);
    if(!data.candidates || !data.candidates.length){
      return {success:false, error:'Gemini no devolvió resultado — ' + (data.error ? data.error.message : 'respuesta vacía') };
    }
    let textoIA = data.candidates[0].content.parts[0].text;
    textoIA = textoIA.replace(/```json|```/g,'').trim();
    const registro = JSON.parse(textoIA);
    return {success:true, registro};
  }catch(err){
    return {success:false, error:'No se pudo leer la foto automáticamente: ' + err};
  }
}

// ═══════════════════════════════════════════════════════════════════
// NUEVO — Lector genérico de documentos de respaldo de discapacidad
// (estudios/informes que avalan el CUD, informes mensuales de
// prestadores). A diferencia de analizarFotoCUD (que espera ESPECÍFICAMENTE
// un CUD), este extrae solo lo básico que hace falta para cualquier
// papel: de qué es, cuándo y qué institución lo hizo — así se puede usar
// para fotografiar y cargar rápido cualquier informe de respaldo.
// ═══════════════════════════════════════════════════════════════════
function analizarFotoDocumentoDiscap(d){
  if(!d || !d.foto_base64) return {success:false, error:'Falta la foto'};
  const mime = d.mime_type || 'image/jpeg';
  const prompt = 'Sos un asistente que transcribe informes/estudios/certificados médicos argentinos (en el marco de un trámite ' +
    'de discapacidad — CUD, pensión no contributiva, etc.) a partir de una foto. ' +
    'Mirá la imagen y extraé EXACTAMENTE estos campos, en español (si un campo no se ve o no figura, dejalo como cadena vacía "", no inventes datos):\n' +
    '- tipo_documento (ej: "Informe kinésico", "Estudio de laboratorio", "Certificado neurológico", etc. — una frase corta describiendo qué es)\n' +
    '- fecha (fecha del documento, formato DD/MM/AAAA)\n' +
    '- institucion (institución, profesional o efector que lo emitió)\n' +
    '- resumen (resumen breve del contenido/hallazgo principal, 1-2 oraciones)\n\n' +
    'Devolvé SOLO un JSON válido, sin texto adicional ni markdown, con exactamente esas 4 claves (todas como texto).';

  try{
    const resp = UrlFetchApp.fetch(getGeminiUrl_(), {
      method:'post', contentType:'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [
          { text: prompt },
          { inline_data: { mime_type: mime, data: d.foto_base64 } }
        ]}],
        generationConfig: {
          thinkingConfig: { thinkingLevel: 'minimal' }
        }
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(resp.getContentText());
    registrarUsoTokens_('analizarFotoDocumentoDiscap', 'gemini-3.6-flash', data.usageMetadata);
    if(!data.candidates || !data.candidates.length){
      return {success:false, error:'Gemini no devolvió resultado — ' + (data.error ? data.error.message : 'respuesta vacía') };
    }
    let textoIA = data.candidates[0].content.parts[0].text;
    textoIA = textoIA.replace(/```json|```/g,'').trim();
    const registro = JSON.parse(textoIA);
    return {success:true, registro};
  }catch(err){
    return {success:false, error:'No se pudo leer la foto automáticamente: ' + err};
  }
}

// ═══════════════════════════════════════════════════════════════════
// NUEVO — Lector de Edad Ósea (GONIO_Columna). El frontend manda la foto
// de una radiografía de mano/muñeca YA TOMADA; esto le pide a Gemini una
// LECTURA BORRADOR por el método elegido (Greulich-Pyle o Tanner-Whitehouse),
// para que el médico la revise, corrija y recién ahí la guarde — igual
// criterio que analizarFotoCUD: nunca se guarda nada automáticamente, y
// el frontend deja bien claro que esto no es un dispositivo médico
// certificado ni tiene la validación de BoneXpert.
// ═══════════════════════════════════════════════════════════════════
function analizarFotoEdadOsea(d){
  if(!d || !d.foto_base64) return {success:false, error:'Falta la foto'};
  const mime = d.mime_type || 'image/jpeg';
  const metodo = (d.metodo && d.metodo.indexOf('Tanner') > -1) ? 'Tanner-Whitehouse (TW2/TW3, por puntaje de cada hueso)' : 'Greulich-Pyle (comparación con la lámina del atlas más parecida)';
  const prompt = 'Sos un asistente que ayuda a un médico a leer la edad ósea de una radiografía de mano y muñeca izquierda, por el método de ' + metodo + '. ' +
    'Esto es SOLO un borrador de apoyo — el médico va a revisar y corregir tu lectura antes de guardar nada, así que priorizá la honestidad sobre dar un número seguro.\n\n' +
    'Mirá la imagen y evaluá el grado de osificación/fusión de las epífisis del radio, cúbito, metacarpianos y falanges, y los núcleos de osificación del carpo. ' +
    'Si la imagen NO es una radiografía de mano/muñeca, está borrosa, mal encuadrada, o no podés ver con claridad las estructuras óseas relevantes, DECILO EXPLÍCITAMENTE en "advertencia" y poné "confianza":"Baja" — NO inventes un número igual.\n\n' +
    'Devolvé SOLO un JSON válido, sin texto adicional ni markdown, con exactamente estas claves:\n' +
    '- edad_osea_estimada (número, en años, tu mejor estimación puntual — o null si no se puede estimar)\n' +
    '- rango_estimado (texto, ej: "13.5 a 14.5 años" — un rango, no un número exacto, reconociendo la incertidumbre normal del método)\n' +
    '- referencia_comparada (texto, ej: "Comparable a la lámina de referencia de varón de 14 años del atlas de Greulich-Pyle" o la descripción equivalente en TW3)\n' +
    '- hallazgos_clave (texto breve: qué epífisis/núcleos usaste para la estimación y por qué)\n' +
    '- confianza ("Alta", "Media" o "Baja")\n' +
    '- advertencia (texto: cualquier limitación de la imagen o del método que el médico deba tener en cuenta; si no hay ninguna, poné cadena vacía "")';

  try{
    const resp = UrlFetchApp.fetch(getGeminiUrl_(), {
      method:'post', contentType:'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [
          { text: prompt },
          { inline_data: { mime_type: mime, data: d.foto_base64 } }
        ]}],
        generationConfig: {
          thinkingConfig: { thinkingLevel: 'minimal' }
        }
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(resp.getContentText());
    registrarUsoTokens_('analizarFotoEdadOsea', 'gemini-3.6-flash', data.usageMetadata);
    if(!data.candidates || !data.candidates.length){
      return {success:false, error:'Gemini no devolvió resultado — ' + (data.error ? data.error.message : 'respuesta vacía') };
    }
    let textoIA = data.candidates[0].content.parts[0].text;
    textoIA = textoIA.replace(/```json|```/g,'').trim();
    const registro = JSON.parse(textoIA);
    return {success:true, registro};
  }catch(err){
    return {success:false, error:'No se pudo leer la foto automáticamente: ' + err};
  }
}

// ═══════════════════════════════════════════════════════════════════
// NUEVO — Análisis por IA de clasificaciones visuales (GONIO_Columna):
// Risser (madurez ósea) y Nash-Moe (rotación vertebral). Mismo criterio
// que analizarFotoEdadOsea: borrador de apoyo, el médico siempre revisa
// y confirma el grado antes de guardar — nunca se guarda automáticamente.
// ═══════════════════════════════════════════════════════════════════
function analizarFotoClasificacionVisual(d){
  if(!d || !d.foto_base64) return {success:false, error:'Falta la foto'};
  if(!d.sistema) return {success:false, error:'Falta indicar el sistema de clasificación'};
  const mime = d.mime_type || 'image/jpeg';
  let prompt;
  if(d.sistema === 'Risser'){
    prompt = 'Sos un asistente que ayuda a un médico traumatólogo a estimar el signo de Risser (madurez ósea, osificación del núcleo apofisario ilíaco) a partir de una foto de una radiografía de pelvis/columna. ' +
      'Esto es SOLO un borrador de apoyo — el médico va a revisar y corregir tu lectura antes de guardar nada, así que priorizá la honestidad sobre dar una respuesta segura.\n\n' +
      'Evaluá el grado de osificación de la apófisis ilíaca: 0 = sin osificación; I = 25% osificado (de lateral a medial); II = 50%; III = 75%; IV = 100% osificado pero aún no fusionado; V = completamente osificado y fusionado a la cresta ilíaca.\n\n' +
      'Si la imagen no muestra con claridad la cresta ilíaca, o no es una radiografía de pelvis/columna, DECILO en "advertencia" y poné "confianza":"Baja" — NO inventes un grado igual.\n\n' +
      'Devolvé SOLO un JSON válido, sin texto adicional ni markdown, con exactamente estas claves:\n' +
      '- grado_estimado ("0","I","II","III","IV" o "V" — o null si no se puede estimar)\n' +
      '- hallazgos_clave (texto breve: qué viste en la cresta ilíaca)\n' +
      '- confianza ("Alta","Media" o "Baja")\n' +
      '- advertencia (texto con cualquier limitación; cadena vacía "" si no hay ninguna)';
  } else if(d.sistema === 'NashMoe'){
    prompt = 'Sos un asistente que ayuda a un médico traumatólogo a estimar el grado de rotación vertebral de Nash-Moe a partir de una foto de una radiografía de columna de frente (AP), sobre la vértebra ápex de una escoliosis. ' +
      'Esto es SOLO un borrador de apoyo — el médico va a revisar y corregir tu lectura antes de guardar nada, así que priorizá la honestidad sobre dar una respuesta segura.\n\n' +
      'Evaluá la posición del pedículo convexo respecto de los tercios del cuerpo vertebral: 0 = pedículos simétricos, cada uno en el tercio externo de su mitad (normal); I = pedículo convexo desplazado levemente hacia el tercio medio, el cóncavo se afina; II = pedículo convexo cerca del centro del cuerpo vertebral; III = pedículo convexo se acerca a la línea media; IV = pedículo convexo pasó la línea media, el cóncavo ya no es visible.\n\n' +
      'Si no podés identificar con claridad la vértebra ápex o los pedículos en la imagen, DECILO en "advertencia" y poné "confianza":"Baja" — NO inventes un grado igual.\n\n' +
      'Devolvé SOLO un JSON válido, sin texto adicional ni markdown, con exactamente estas claves:\n' +
      '- grado_estimado ("0","I","II","III" o "IV" — o null si no se puede estimar)\n' +
      '- hallazgos_clave (texto breve: qué viste en los pedículos de la vértebra ápex)\n' +
      '- confianza ("Alta","Media" o "Baja")\n' +
      '- advertencia (texto con cualquier limitación; cadena vacía "" si no hay ninguna)';
  } else {
    return {success:false, error:'Sistema de clasificación no reconocido: ' + d.sistema};
  }

  try{
    const resp = UrlFetchApp.fetch(getGeminiUrl_(), {
      method:'post', contentType:'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [
          { text: prompt },
          { inline_data: { mime_type: mime, data: d.foto_base64 } }
        ]}],
        generationConfig: {
          thinkingConfig: { thinkingLevel: 'minimal' }
        }
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(resp.getContentText());
    registrarUsoTokens_('analizarFotoClasificacionVisual_'+d.sistema, 'gemini-3.6-flash', data.usageMetadata);
    if(!data.candidates || !data.candidates.length){
      return {success:false, error:'Gemini no devolvió resultado — ' + (data.error ? data.error.message : 'respuesta vacía') };
    }
    let textoIA = data.candidates[0].content.parts[0].text;
    textoIA = textoIA.replace(/```json|```/g,'').trim();
    const registro = JSON.parse(textoIA);
    return {success:true, registro};
  }catch(err){
    return {success:false, error:'No se pudo leer la foto automáticamente: ' + err};
  }
}

// ═══════════════════════════════════════════════════════════════════
// NUEVO — Lector de Historias Clínicas viejas (Word/papel, texto pegado):
// el médico pega el texto tal cual estaba en el Word, la IA propone los
// datos de filiación + una lista de consultas con su fecha real, y recién
// después de revisar se guarda — mismo criterio que el Lector CUD, pero
// de texto en vez de foto (no hace falta Vision para esto).
// ═══════════════════════════════════════════════════════════════════
function analizarTextoPaciente(d){
  if(!d || !d.texto) return {success:false, error:'Falta el texto a analizar'};
  const prompt = 'Sos un asistente que transcribe historias clínicas argentinas viejas (escritas originalmente en Word o a mano y ' +
    'luego tipeadas) a partir del texto tal cual fue copiado del documento. El texto puede tener errores de tipeo, mayúsculas ' +
    'sueltas, o letras mal codificadas (por ejemplo "A±OS" en vez de "AÑOS", o "AG?ERO" en vez de "AGÜERO") — interpretalas ' +
    'igual, corrigiendo solo esos problemas de codificación, sin inventar ni completar datos que no estén.\n\n' +
    'Extraé EXACTAMENTE estos campos de FILIACIÓN (si un dato no está, dejalo como cadena vacía "", no inventes nada):\n' +
    '- apellido\n- nombre\n- dni (solo números)\n- edad (tal como figura, ej: "54 años")\n' +
    '- fecha_nac (SOLO si figura explícita en el texto, formato DD/MM/AAAA — NO calcules una fecha de nacimiento a partir de la edad, eso sería inventar un dato)\n' +
    '- obra_social (si el texto dice explícitamente que no tiene, particular, o no mutualizado, dejalo igual vacío — eso se marca aparte)\n' +
    '- plan\n- nro_afiliado\n- telefono\n' +
    '- email (si falta el signo @ antes del dominio, por ejemplo "nombreGMAIL.COM", corregilo asumiendo que se perdió el @, pero marcá "email_dudoso": true en ese caso)\n' +
    '- domicilio\n- localidad (si se puede inferir de la dirección)\n' +
    '- centro_derivador (institución, colega o médico que derivó al paciente, SOLO si se menciona explícitamente en el texto)\n\n' +
    'Devolvé SOLO un JSON válido, sin texto adicional ni markdown, con exactamente estas claves: apellido, nombre, dni, edad, ' +
    'fecha_nac, obra_social, plan, nro_afiliado, telefono, email, email_dudoso, domicilio, localidad, centro_derivador.\n\n' +
    'TEXTO A ANALIZAR:\n' + d.texto;

  try{
    const resp = UrlFetchApp.fetch(getGeminiUrl_(), {
      method:'post', contentType:'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [ { text: prompt } ]}],
        generationConfig: { thinkingConfig: { thinkingLevel: 'minimal' } }
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(resp.getContentText());
    registrarUsoTokens_('analizarTextoPaciente', 'gemini-3.6-flash', data.usageMetadata);
    if(!data.candidates || !data.candidates.length){
      return {success:false, error:'Gemini no devolvió resultado — ' + (data.error ? data.error.message : 'respuesta vacía') };
    }
    let textoIA = data.candidates[0].content.parts[0].text;
    textoIA = textoIA.replace(/```json|```/g,'').trim();
    const registro = JSON.parse(textoIA);
    return {success:true, registro};
  }catch(err){
    return {success:false, error:'No se pudo analizar el texto automáticamente: ' + err};
  }
}

// ══════════════════════════════════════════════════════════════════
// Asistente genérico de texto con IA — usado por la Consulta Guiada Beta
// (voz + IA): redactar en prosa clínica lo que el médico va dictando,
// sugerir maniobras de examen físico según el motivo/relato, y armar el
// resumen para el paciente. El prompt completo lo arma el front (cada
// tarea necesita instrucciones bien distintas), acá solo se lo pasamos
// a Gemini tal cual y devolvemos el texto de respuesta.
// ══════════════════════════════════════════════════════════════════
function asistenteConsultaIA(d){
  if(!d || !d.prompt) return {success:false, error:'Falta el texto a procesar'};
  try{
    const resp = UrlFetchApp.fetch(getGeminiUrl_(), {
      method:'post', contentType:'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: d.prompt }] }],
        generationConfig: { thinkingConfig: { thinkingLevel: 'minimal' } }
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(resp.getContentText());
    registrarUsoTokens_('asistenteConsultaIA', 'gemini-3.6-flash', data.usageMetadata);
    if(!data.candidates || !data.candidates.length){
      return {success:false, error:'Gemini no devolvió resultado — ' + (data.error ? data.error.message : 'respuesta vacía')};
    }
    const texto = data.candidates[0].content.parts[0].text || '';
    return {success:true, texto: texto.trim()};
  }catch(err){
    return {success:false, error:'No se pudo conectar con la IA: ' + err};
  }
}

// Misma idea que analizarTextoPaciente(), pero para cuando el médico tiene
// una FOTO de la ficha vieja (papel o Word impreso) en vez de poder pegar el
// texto — por ejemplo, sacada con la cámara web de la PC o el celular. Usa
// Vision (igual que analizarFotoCUD) y devuelve EXACTAMENTE el mismo esquema
// de campos que analizarTextoPaciente(), así el mismo formulario de revisión
// sirve para las dos vías de entrada.
function analizarFotoHistoria(d){
  if(!d || !d.foto_base64) return {success:false, error:'Falta la foto'};
  const mime = d.mime_type || 'image/jpeg';
  const prompt = 'Sos un asistente que transcribe historias clínicas argentinas viejas (fichas en papel o Word impreso) a partir ' +
    'de una foto. Puede haber letra manuscrita, mala calidad de escaneo o reflejos — hacé tu mejor lectura, pero si un dato no se ' +
    'llega a leer con confianza, dejalo como cadena vacía "" en vez de adivinar.\n\n' +
    'Extraé EXACTAMENTE estos campos de FILIACIÓN (si un dato no está o no se lee, dejalo como cadena vacía "", no inventes nada):\n' +
    '- apellido\n- nombre\n- dni (solo números)\n- edad (tal como figura, ej: "54 años")\n' +
    '- fecha_nac (SOLO si figura explícita en la foto, formato DD/MM/AAAA — NO calcules una fecha de nacimiento a partir de la edad, eso sería inventar un dato)\n' +
    '- obra_social (si el texto dice explícitamente que no tiene, particular, o no mutualizado, dejalo igual vacío — eso se marca aparte)\n' +
    '- plan\n- nro_afiliado\n- telefono\n' +
    '- email (si falta el signo @ antes del dominio, por ejemplo "nombreGMAIL.COM", corregilo asumiendo que se perdió el @, pero marcá "email_dudoso": true en ese caso)\n' +
    '- domicilio\n- localidad (si se puede inferir de la dirección)\n' +
    '- centro_derivador (institución, colega o médico que derivó al paciente, SOLO si se menciona explícitamente)\n\n' +
    'Devolvé SOLO un JSON válido, sin texto adicional ni markdown, con exactamente estas claves: apellido, nombre, dni, edad, ' +
    'fecha_nac, obra_social, plan, nro_afiliado, telefono, email, email_dudoso, domicilio, localidad, centro_derivador.';

  try{
    const resp = UrlFetchApp.fetch(getGeminiUrl_(), {
      method:'post', contentType:'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [
          { text: prompt },
          { inline_data: { mime_type: mime, data: d.foto_base64 } }
        ]}],
        generationConfig: { thinkingConfig: { thinkingLevel: 'minimal' } }
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(resp.getContentText());
    registrarUsoTokens_('analizarFotoHistoria', 'gemini-3.6-flash', data.usageMetadata);
    if(!data.candidates || !data.candidates.length){
      return {success:false, error:'Gemini no devolvió resultado — ' + (data.error ? data.error.message : 'respuesta vacía') };
    }
    let textoIA = data.candidates[0].content.parts[0].text;
    textoIA = textoIA.replace(/```json|```/g,'').trim();
    const registro = JSON.parse(textoIA);
    return {success:true, registro};
  }catch(err){
    return {success:false, error:'No se pudo leer la foto automáticamente: ' + err};
  }
}

// Guarda el paciente (upsert por DNI, reutilizando addPaciente) y, si vino
// una lista de historial, agrega una fila POR CADA entrada en Historia_Clinica
// con su FECHA REAL (a diferencia de addHistoria(), que siempre usa la fecha
// de hoy — acá el médico puede estar cargando una consulta de años atrás).
function guardarPacienteConHistorial(d){
  // Normaliza a las claves con mayúscula inicial que espera addPaciente(),
  // aceptando también las claves en minúscula que devuelve la IA (apellido,
  // nombre, dni, etc.) — así el front-end puede mandar cualquiera de las dos.
  const pick=(a,b)=> (d[a]!==undefined && d[a]!==null && d[a]!=='') ? d[a] : d[b];
  const dPaciente = {
    Apellido: pick('Apellido','apellido'), Nombre: pick('Nombre','nombre'), DNI: pick('DNI','dni'),
    Fecha_Nac: pick('Fecha_Nac','fecha_nac'), Obra_Social: pick('Obra_Social','obra_social'),
    N_Afiliado: pick('N_Afiliado','nro_afiliado'), Plan: pick('Plan','plan'),
    Telefono: pick('Telefono','telefono'), Email: pick('Email','email'),
    Domicilio: pick('Domicilio','domicilio'), Lugar_Atencion: pick('Lugar_Atencion','lugar_atencion'),
    Localidad: pick('Localidad','localidad'), Estado_Tratamiento: d.Estado_Tratamiento || 'Activo',
    Centro_Derivador: pick('Centro_Derivador','centro_derivador'),
    // Si se tilda "primera consulta" en el Lector de Fichas, addPaciente() manda
    // el RIMP de antecedentes ahora mismo (en vez de esperar al recordatorio semanal).
    Primera_Consulta: !!(d.Primera_Consulta || d.primera_consulta),
    Discapacidad: !!(d.Discapacidad || d.discapacidad)
  };
  const resultadoPaciente = addPaciente(dPaciente);
  let historiasGuardadas = 0;
  if(Array.isArray(d.historial) && d.historial.length){
    const ss=SpreadsheetApp.getActiveSpreadsheet(), sheet=ss.getSheetByName('Historia_Clinica');
    if(sheet){
      let headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
      let dniCol=headers.indexOf('DNI');
      if(dniCol===-1){
        dniCol=headers.length;
        sheet.getRange(1,dniCol+1).setValue('DNI');
        headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
      }
      d.historial.forEach(h=>{
        if(!h || !h.texto) return;
        const id='HC-'+String(sheet.getLastRow()).padStart(4,'0');
        const fila=new Array(headers.length).fill('');
        const set=(nombreCol,valor)=>{const i=headers.indexOf(nombreCol);if(i>-1)fila[i]=valor;};
        set('ID_HC',id);
        set('Fecha', h.fecha || new Date().toLocaleDateString('es-AR'));
        set('Observaciones', h.texto);
        set('DNI', d.DNI||d.dni||'');
        sheet.appendRow(fila);
        historiasGuardadas++;
      });
    }
  }
  return {success:true, paciente:resultadoPaciente, historiasGuardadas};
}

function probarPermisoMail(){
  MailApp.sendEmail('lmuratori@gmail.com', 'Prueba de permisos SMM', 'Si recibis esto, el permiso ya esta autorizado.');
}

// ══════════════════════════════════════════════════════════════════
// PARCHE — Autoevaluación Semanal RMYT
// A diferencia de RECOMENDACIONES MEDICAS (que guarda un solo documento por paciente y lo
// pisa cada vez), acá CADA autoevaluación semanal se agrega como una fila
// nueva — necesitamos el historial completo, no solo la última, para poder
// graficar la evolución semana a semana.
// ══════════════════════════════════════════════════════════════════
function saveAutoevaluacionRMYT(d){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('RMYT_Autoevaluaciones');
  if(!sheet){
    sheet = ss.insertSheet('RMYT_Autoevaluaciones');
    sheet.appendRow(['Timestamp','DNI','Segmento','EVA_Reposo','EVA_Movimiento',
      'Neurodinamico','Incapacidad_Pct','Recuperacion_Pct','Escala_Nombre',
      'Escala_Pct_Discapacidad','Cumplimiento_Prom','Alerta','Tipo_Dolor',
      'Medicacion_Tipo','Medicacion_Detalle','Efectos_Adversos','Intolerancias',
      'Gastritis','Datos_JSON']);
    sheet.setFrozenRows(1);
  }
  const datos = d.datos || {};
  const dni = d.dni || '';
  if(!dni) return {success:false, error:'Falta DNI del paciente'};
  // Beneficio exclusivo de paciente con membresía (ver comentario en
  // addPaciente/Membresia) — filiación sola no habilita el seguimiento
  // semanal RMYT. No se guarda nada de esta autoevaluación si el DNI no
  // tiene membresía activa hoy.
  if(!tienePacienteMembresia_(dni)){
    return {success:false, error:'Este paciente todavía no tiene membresía activa. El seguimiento semanal RMYT es un beneficio de membresía — otorgala desde OCM para habilitarlo.'};
  }

  const cump = datos.cumplimiento || {};
  const valoresCump = [cump.medicacion, cump.ejercicios, cump.kinesio, cump.actividad_guiada]
    .filter(v => v !== null && v !== undefined);
  const cumplimientoProm = valoresCump.length
    ? Math.round(valoresCump.reduce((a,b)=>a+b,0) / valoresCump.length * 10) / 10
    : null;

  // ── Semáforo de alerta: EVA sube ≥3 puntos vs la semana anterior, reporta
  // corriente eléctrica constante, efecto adverso o gastritis → bandera roja
  // para revisión (búsqueda por nombre de columna, no por índice fijo, para
  // que funcione igual en hojas viejas y nuevas — mismo criterio que se usó
  // para la cola de mensajes del QR).
  let alerta = '';
  const evaMov = datos.dolor ? datos.dolor.eva_movimiento : null;
  const filasPrevias = sheet.getDataRange().getValues();
  const headersPrevios = filasPrevias[0] || [];
  const colDNIPrevio = headersPrevios.indexOf('DNI');
  const colEvaMovPrevio = headersPrevios.indexOf('EVA_Movimiento');
  let evaAnterior = null;
  if(colDNIPrevio > -1 && colEvaMovPrevio > -1){
    for(let i = filasPrevias.length - 1; i >= 1; i--){
      if(String(filasPrevias[i][colDNIPrevio]) === String(dni)){
        evaAnterior = filasPrevias[i][colEvaMovPrevio];
        break;
      }
    }
  }
  if(evaAnterior !== null && evaAnterior !== '' && evaMov !== null && (evaMov - evaAnterior) >= 3) alerta = 'DOLOR_EN_AUMENTO';
  if(datos.neurodinamico === 2) alerta = alerta ? alerta + ' + NEURO_CONSTANTE' : 'NEURO_CONSTANTE';
  const efAdv = datos.efectos_adversos || {};
  if(efAdv.tiene) alerta = alerta ? alerta + ' + EFECTO_ADVERSO' : 'EFECTO_ADVERSO';
  if(datos.gastritis) alerta = alerta ? alerta + ' + GASTRITIS' : 'GASTRITIS';

  // Asegura que existan todas las columnas (compatibilidad con hojas creadas
  // antes de agregar tipo de dolor / medicación / efectos adversos).
  const cols = ['Timestamp','DNI','Segmento','EVA_Reposo','EVA_Movimiento','Neurodinamico',
    'Incapacidad_Pct','Recuperacion_Pct','Escala_Nombre','Escala_Pct_Discapacidad',
    'Cumplimiento_Prom','Alerta','Tipo_Dolor','Medicacion_Tipo','Medicacion_Detalle',
    'Efectos_Adversos','Intolerancias','Gastritis','Datos_JSON'];
  const colIdx = {};
  cols.forEach(c => { colIdx[c] = asegurarColumnaColaMensajes_(sheet, c) - 1; }); // 0-based

  const med = datos.medicacion || {};
  const medicacionDetalle = (med.tipo === 'farmacopea' && Array.isArray(med.farmacopea))
    ? med.farmacopea.map(f => (String(f.farmaco||'') + ' ' + String(f.dosis||'')).trim()).filter(Boolean).join(' | ')
    : (med.tipo === 'magistral' ? ('Magistral — ' + (med.magistral_farmacia || 'sin especificar')) : '');
  const intol = datos.intolerancias || {};

  const fila = new Array(cols.length).fill('');
  fila[colIdx['Timestamp']] = new Date().toISOString();
  fila[colIdx['DNI']] = dni;
  fila[colIdx['Segmento']] = (datos.segmento && datos.segmento.titulo) || '';
  fila[colIdx['EVA_Reposo']] = datos.dolor ? datos.dolor.eva_reposo : '';
  fila[colIdx['EVA_Movimiento']] = datos.dolor ? datos.dolor.eva_movimiento : '';
  fila[colIdx['Neurodinamico']] = datos.neurodinamico != null ? datos.neurodinamico : '';
  fila[colIdx['Incapacidad_Pct']] = datos.incapacidad_relativa_pct != null ? datos.incapacidad_relativa_pct : '';
  fila[colIdx['Recuperacion_Pct']] = datos.recuperacion_percibida_pct != null ? datos.recuperacion_percibida_pct : '';
  fila[colIdx['Escala_Nombre']] = (datos.segmento && datos.segmento.escala_nombre) || '';
  fila[colIdx['Escala_Pct_Discapacidad']] = (datos.segmento && datos.segmento.escala_pct_discapacidad != null) ? datos.segmento.escala_pct_discapacidad : '';
  fila[colIdx['Cumplimiento_Prom']] = cumplimientoProm != null ? cumplimientoProm : '';
  fila[colIdx['Alerta']] = alerta;
  fila[colIdx['Tipo_Dolor']] = (datos.dolor && Array.isArray(datos.dolor.tipo)) ? datos.dolor.tipo.join(', ') : '';
  fila[colIdx['Medicacion_Tipo']] = med.tipo || '';
  fila[colIdx['Medicacion_Detalle']] = medicacionDetalle;
  fila[colIdx['Efectos_Adversos']] = efAdv.tiene ? (efAdv.detalle || 'Sí, sin detalle') : '';
  fila[colIdx['Intolerancias']] = intol.tiene ? (intol.detalle || 'Sí, sin detalle') : '';
  fila[colIdx['Gastritis']] = datos.gastritis ? 'Sí' : '';
  fila[colIdx['Datos_JSON']] = JSON.stringify(datos);

  sheet.getRange(sheet.getLastRow()+1, 1, 1, cols.length).setValues([fila]);

  return {success:true, message:'Autoevaluación registrada', alerta: alerta || null};
}

function getAutoevaluacionesRMYT(dni){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('RMYT_Autoevaluaciones');
  if(!sheet || !dni) return {success:true, data:[], tendencia:null};
  const rows = sheet.getDataRange().getValues();
  if(rows.length < 2) return {success:true, data:[], tendencia:null};
  const headers = rows[0];
  const idx = (nombre) => headers.indexOf(nombre);
  const iDNI=idx('DNI'), iFecha=idx('Timestamp'), iSeg=idx('Segmento'), iEvaR=idx('EVA_Reposo'),
    iEvaM=idx('EVA_Movimiento'), iNeuro=idx('Neurodinamico'), iIncap=idx('Incapacidad_Pct'),
    iRecup=idx('Recuperacion_Pct'), iEscN=idx('Escala_Nombre'), iEscPct=idx('Escala_Pct_Discapacidad'),
    iCump=idx('Cumplimiento_Prom'), iAlerta=idx('Alerta'), iTipoDolor=idx('Tipo_Dolor'),
    iMedTipo=idx('Medicacion_Tipo'), iMedDet=idx('Medicacion_Detalle'), iEfAdv=idx('Efectos_Adversos'),
    iIntol=idx('Intolerancias'), iGastritis=idx('Gastritis');

  const resultado = [];
  for(let i = 1; i < rows.length; i++){
    if(String(rows[i][iDNI]) === String(dni)){
      resultado.push({
        fecha: rows[i][iFecha],
        segmento: rows[i][iSeg],
        eva_reposo: rows[i][iEvaR],
        eva_movimiento: rows[i][iEvaM],
        neurodinamico: rows[i][iNeuro],
        incapacidad_pct: rows[i][iIncap],
        recuperacion_pct: rows[i][iRecup],
        escala_nombre: rows[i][iEscN],
        escala_pct_discapacidad: rows[i][iEscPct],
        cumplimiento_promedio: rows[i][iCump],
        alerta: rows[i][iAlerta],
        tipo_dolor: iTipoDolor>-1 ? rows[i][iTipoDolor] : '',
        medicacion_tipo: iMedTipo>-1 ? rows[i][iMedTipo] : '',
        medicacion_detalle: iMedDet>-1 ? rows[i][iMedDet] : '',
        efectos_adversos: iEfAdv>-1 ? rows[i][iEfAdv] : '',
        intolerancias: iIntol>-1 ? rows[i][iIntol] : '',
        gastritis: iGastritis>-1 ? rows[i][iGastritis] : ''
      });
    }
  }
  // Más reciente primero para mostrarlo en pantalla…
  resultado.sort((a,b) => new Date(b.fecha) - new Date(a.fecha));
  // …pero la tendencia se calcula en orden cronológico ascendente.
  const cronologico = resultado.slice().reverse();
  const tendencia = calcularTendenciaRMYT_(cronologico);
  return {success:true, data: resultado, tendencia};
}

// ══════════════════════════════════════════════════════════════════
// Motor predictivo simple del RMYT
// ══════════════════════════════════════════════════════════════════
// Compara la evolución OBJETIVA (% de discapacidad de la escala del
// segmento, o el EVA de movimiento si todavía no hay escala) contra la
// evolución SUBJETIVA (recuperación percibida que informa el propio
// paciente) semana a semana, mediante una regresión lineal simple
// (pendiente por semana). El objetivo es exactamente el que se pidió:
// poder distinguir, con el tiempo, una mejoría real y sostenida de una
// simple apreciación personal. NO es un modelo clínico ni reemplaza el
// criterio médico — por eso el resultado siempre se devuelve marcado
// como estimación (es_estimacion:true) y con lenguaje prudente.
function calcularTendenciaRMYT_(registrosCronologicos){
  const regs = (registrosCronologicos || []).filter(r => r.fecha);
  if(regs.length < 2){
    return {
      estado: 'datos_insuficientes',
      n_registros: regs.length,
      es_estimacion: true,
      mensaje: regs.length === 0
        ? 'Todavía no hay autoevaluaciones cargadas para este paciente.'
        : 'Hace falta al menos una segunda autoevaluación semanal para poder estimar una tendencia.'
    };
  }

  const t0 = new Date(regs[0].fecha).getTime();
  const semanas = regs.map(r => (new Date(r.fecha).getTime() - t0) / (1000*60*60*24*7));

  function pendientePorSemana(valores){
    const pares = [];
    for(let i=0; i<valores.length; i++){
      const v = valores[i];
      if(v !== '' && v !== null && v !== undefined && !isNaN(v)) pares.push([semanas[i], Number(v)]);
    }
    if(pares.length < 2) return null;
    const n = pares.length;
    const sx = pares.reduce((a,p)=>a+p[0],0), sy = pares.reduce((a,p)=>a+p[1],0);
    const sxy = pares.reduce((a,p)=>a+p[0]*p[1],0), sxx = pares.reduce((a,p)=>a+p[0]*p[0],0);
    const denom = (n*sxx - sx*sx);
    if(Math.abs(denom) < 1e-9) return 0; // todas las mediciones cayeron en la misma semana
    return (n*sxy - sx*sy) / denom;
  }

  const pendEva = pendientePorSemana(regs.map(r=>r.eva_movimiento));
  const pendDiscapacidad = pendientePorSemana(regs.map(r=>r.escala_pct_discapacidad));
  const pendRecuperacionPercibida = pendientePorSemana(regs.map(r=>r.recuperacion_pct));

  // Indicador objetivo preferido: % de discapacidad de la escala del
  // segmento; si todavía no hay escala cargada, se usa el EVA de
  // movimiento (llevado a una escala 0-100 para que sea comparable).
  let pendObjetiva = null, fuenteObjetiva = null;
  if(pendDiscapacidad !== null){ pendObjetiva = pendDiscapacidad; fuenteObjetiva = 'discapacidad'; }
  else if(pendEva !== null){ pendObjetiva = pendEva * 10; fuenteObjetiva = 'eva'; }

  let estado, mensaje;
  if(pendObjetiva === null){
    estado = 'datos_insuficientes';
    mensaje = 'No hay suficientes valores numéricos cargados todavía para estimar una tendencia objetiva.';
  } else if(pendObjetiva <= -2){
    estado = 'mejoria_sostenida';
    mensaje = 'Los indicadores objetivos (' + (fuenteObjetiva==='discapacidad' ? '% de discapacidad de la escala' : 'EVA de movimiento') + ') muestran una mejoría sostenida semana a semana.';
  } else if(pendObjetiva < 2){
    estado = 'estancado';
    mensaje = 'Los indicadores objetivos se mantienen prácticamente sin cambios entre evaluaciones — conviene revisar el plan terapéutico.';
  } else {
    estado = 'empeorando';
    mensaje = 'Los indicadores objetivos muestran una tendencia al empeoramiento — se recomienda revisión clínica prioritaria.';
  }

  // Discordancia entre lo objetivo y lo percibido: esto es justamente lo
  // que permite distinguir "mejoría evidente" de "apreciación personal".
  let discordancia = false, mensajeDiscordancia = null;
  if(pendRecuperacionPercibida !== null && pendObjetiva !== null){
    const percibidaMejorando = pendRecuperacionPercibida > 2;
    const objetivaMejorando = pendObjetiva <= -2;
    const objetivaEstancadaOPeor = pendObjetiva > -2;
    if(percibidaMejorando && objetivaEstancadaOPeor){
      discordancia = true;
      mensajeDiscordancia = 'El paciente refiere sentirse mejor, pero eso todavía no se refleja en los indicadores objetivos — podría ser una mejoría más subjetiva/anímica que funcional por ahora.';
    } else if(!percibidaMejorando && objetivaMejorando){
      discordancia = true;
      mensajeDiscordancia = 'Los indicadores objetivos mejoran pero el paciente no lo percibe así — vale la pena indagar dolor residual, expectativas u otros factores.';
    }
  }

  // Proyección simple: solo si hay mejoría objetiva clara en la escala de
  // discapacidad, estima en cuántas semanas llegaría a ≤20% (umbral
  // habitual de "discapacidad leve" en escalas tipo Oswestry) por
  // extrapolación lineal. Se omite si da un plazo demasiado largo/incierto.
  let proyeccionSemanas = null;
  if(fuenteObjetiva === 'discapacidad' && pendDiscapacidad < -0.5){
    const ultimo = regs[regs.length-1];
    const valorActual = Number(ultimo.escala_pct_discapacidad);
    if(!isNaN(valorActual) && valorActual > 20){
      const semanasEstim = Math.ceil((valorActual - 20) / Math.abs(pendDiscapacidad));
      if(semanasEstim <= 52) proyeccionSemanas = semanasEstim;
    }
  }

  return {
    estado, mensaje, discordancia,
    mensaje_discordancia: mensajeDiscordancia,
    pendiente_eva_semana: pendEva !== null ? Math.round(pendEva*100)/100 : null,
    pendiente_discapacidad_semana: pendDiscapacidad !== null ? Math.round(pendDiscapacidad*100)/100 : null,
    pendiente_recuperacion_percibida_semana: pendRecuperacionPercibida !== null ? Math.round(pendRecuperacionPercibida*100)/100 : null,
    proyeccion_semanas_a_leve: proyeccionSemanas,
    n_registros: regs.length,
    es_estimacion: true
  };
}

// ── RMYT: link y envío del formulario semanal de autoevaluación ──────
// Mismo patrón que el RIMP (construirLinkRimpAntecedentes_ / enviarLinkRimpAPaciente_):
// arma un link personalizado por DNI + segmento y lo manda por mail. Queda
// lista para que OCM la dispare desde el selector de RMYT/RECOMENDACIONES MEDICAS — el
// cableado del botón en OCM es un paso aparte, todavía pendiente.
function getRmytUrl_(){
  // OJO: el archivo real que está publicado y que lee estos parámetros
  // (?dni= y ?seg=) es RMYT_Autoevaluacion_Semanal_v1.html — el nombre
  // viejo (RMYT_Autoevaluacion.html, sin "_Semanal") no está publicado.
  return baseSimm_() + 'RMYT_Autoevaluacion_Semanal_v1.html'; // Subdominio de S.I.M-M (Cloudflare Pages)
}
function construirLinkRMYTAutoeval_(p){
  // p.seg debe ser una de las claves de SEGMENTOS_RMYT del html (p.ej.
  // 'hombro','rodilla','cadera'...) — si no coincide ninguna, la página
  // igual abre con el formulario genérico, sin romperse.
  return getRmytUrl_() + '?dni=' + encodeURIComponent(p.dni||'') +
    (p.seg ? '&seg=' + encodeURIComponent(p.seg) : '');
}
function enviarLinkRMYTAPaciente_(p){
  if(!p || !p.email) throw new Error('Falta el mail del paciente');
  const link = construirLinkRMYTAutoeval_(p);
  const cuerpo =
    'Hola ' + (p.nombre||'') + ' ' + (p.apellido||'') + ',\n\n' +
    'Como parte de tu seguimiento' + (p.segmento ? (' (' + p.segmento + ')') : '') + ' te pedimos que completes esta breve autoevaluación semanal — te toma un par de minutos y nos ayuda a ver cómo vas evolucionando.\n\n' +
    'Por favor, entrá al siguiente link:\n' + link + '\n\n' +
    'Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943';
  MailApp.sendEmail(p.email, 'Autoevaluación semanal de seguimiento' + (p.segmento ? (' — ' + p.segmento) : ''), cuerpo, {
    replyTo:'drluismuratori@gmail.com', name:'Dr. Luis Alberto Muratori'
  });
}

// ══════════════════════════════════════════════════════════════════
// SEGUIMIENTO POSQUIRÚRGICO (SSQ) — 24 meses (sep 2026, unificado)
//
// Reemplaza el diseño anterior de esta sección (9 parámetros clínicos
// inventados por mí, puntaje que restaba). Este es el diseño UNIFICADO,
// alineado a dos fuentes reales que el Dr. Muratori ya tenía construidas
// y que yo no había visto todavía:
//   1) SSQ_Legajo_Integral_v1.html → pestaña "Seguimiento 24m": catálogo
//      real de eventos (SEG_CATALOGO_EVENTOS), si están o no relacionados
//      con la cirugía, y su propio "índice de evolución libre de eventos".
//      Hoy ese archivo es 100% standalone (guarda con descarga/mail de
//      JSON, nada llega a un backend) — acá lo conectamos a un backend real.
//   2) Revision_Integracion_SSQ_Posquirurgico.md (5-sep-2026) → ya había
//      diseñado esto mismo contra el código real: hoja `Seguimiento_Posquirurgico`,
//      dolor EVA con alerta DOLOR_EN_AUMENTO ≥3 puntos (mismo criterio que
//      RMYT), acción `ssq_guardar_seguimiento_posqx`. Se sigue ese plano
//      literal en vez de inventar nombres nuevos otra vez.
//
// Pedido explícito del Dr. Muratori: "que los factores de seguimiento
// sumen, no resten" → el Puntaje_Mes se arma GANANDO puntos por cada señal
// de evolución segura (no al revés, restando de una base de 100). Detalle
// de los pesos, ⚠️ orientativos, no clínicamente validados todavía:
//   +50  si no hubo ningún evento de salud ese mes ("Ninguno")
//   +35  si hubo evento pero se marcó explícitamente "no relacionado" con la cirugía
//   +25  si el dolor EVA no aumentó respecto del control anterior (o es el primer control)
//   +25  si se marcó evidencia de mejoría ese mes
// Máximo 100/mes. La puntuación de seguridad global de cada paciente
// (la que se refleja en el dashboard de Quirófano) es el promedio de los
// Puntaje_Mes de todos los controles ya cargados.
// ══════════════════════════════════════════════════════════════════
const SEG_CATALOGO_EVENTOS_ = [
  'Ninguno','Dolor o inflamación inusual','Secreción o supuración de la herida','Fiebre',
  'Dehiscencia de herida (se abrió la herida)','Reinternación','Nueva cirugía o reintervención',
  'Aflojamiento o falla de implante','Trombosis / embolia','Evento médico no relacionado (ej. gripe, otra enfermedad)','Otro'
];

function hojaSeguimientoPosquirurgico_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Seguimiento_Posquirurgico');
  if(!sheet){
    sheet = ss.insertSheet('Seguimiento_Posquirurgico');
    sheet.appendRow(['Timestamp','DNI','CodigoCaso','Mes_Control','Etapa_Postop','Fecha_Cirugia',
      'Evento','Relacion','Fecha_Evento','Dolor_EVA','Evidencia_Mejoria','Comentario',
      'Puntaje_Mes','Alerta','Datos_JSON']);
    sheet.getRange(1,1,1,15).setBackground('#0c4a6e').setFontColor('white').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Clasificación de la etapa post-quirúrgica según la bibliografía clásica
// de cuidados post-operatorios (enfermería/cirugía general): inmediato =
// primeras 24-48 h, mediato = hasta ~7 días (cicatrización inicial/alta),
// tardío = de ahí en adelante hasta la recuperación funcional completa.
// Ojo: distintas fuentes usan cortes ligeramente distintos (algunas
// extienden el inmediato a 72 h) — si el libro que le entregás al
// paciente usa otros números, decímelos y los cambio acá.
// Con la cadencia MENSUAL de este control, en la práctica casi todos los
// registros van a caer en "tardío" — el inmediato y el mediato ocurren en
// los primeros días, cuando el paciente todavía está en control directo
// (internación/alta) o dentro del primer mes de mensajes gratuitos que
// ya definiste, no en este formulario mensual.
function calcularEtapaPostQuirurgica_(fechaCirugiaStr, fechaControl){
  if(!fechaCirugiaStr) return '';
  const parsearFecha_ = f => {
    const s = String(f).trim();
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); // DD/MM/YYYY (es-AR)
    if(m) return new Date(Number(m[3]), Number(m[2])-1, Number(m[1]));
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  };
  const fCirugia = parsearFecha_(fechaCirugiaStr);
  if(!fCirugia) return '';
  const dias = Math.round((new Date(fechaControl) - fCirugia) / 86400000);
  if(dias < 0) return '';
  if(dias <= 1) return 'inmediato';
  if(dias <= 7) return 'mediato';
  return 'tardío';
}

function guardarSeguimientoPosqx(d){
  const dni = d.dni || '';
  if(!dni) return {success:false, error:'Falta DNI del paciente'};
  const mesControl = parseInt(d.mesControl, 10) || null;
  if(!mesControl || mesControl < 1 || mesControl > 24) return {success:false, error:'Mes de control inválido (debe ser 1 a 24)'};

  const evento = SEG_CATALOGO_EVENTOS_.indexOf(d.evento) > -1 ? d.evento : 'Ninguno';
  const relacion = evento === 'Ninguno' ? 'no_aplica' : (d.relacion || 'no_seguro');
  const dolorEva = d.dolorEva !== undefined && d.dolorEva !== '' ? Number(d.dolorEva) : null;
  const evidenciaMejoria = d.evidenciaMejoria === true || d.evidenciaMejoria === 'true';

  const sheet = hojaSeguimientoPosquirurgico_();

  // Dolor vs el control anterior de este mismo paciente — mismo criterio
  // (≥3 puntos) que ya usa el semáforo de RMYT y que pedía explícitamente
  // el punto 6 de Revision_Integracion_SSQ_Posquirurgico.md.
  const filasPrevias = sheet.getDataRange().getValues();
  const headersPrevios = filasPrevias[0] || [];
  const colDNIPrevio = headersPrevios.indexOf('DNI');
  const colDolorPrevio = headersPrevios.indexOf('Dolor_EVA');
  let dolorAnterior = null;
  let esPrimerControl = true;
  if(colDNIPrevio > -1){
    const delPaciente = filasPrevias.slice(1).filter(r => String(r[colDNIPrevio]) === String(dni));
    esPrimerControl = delPaciente.length === 0;
    if(delPaciente.length) dolorAnterior = delPaciente[delPaciente.length - 1][colDolorPrevio];
  }
  const dolorEnAumento = (dolorAnterior !== null && dolorAnterior !== '' && dolorEva !== null && (dolorEva - Number(dolorAnterior)) >= 3);

  // ── Puntaje del mes: SUMA de señales de evolución segura, nunca resta ──
  let puntaje = 0;
  if(evento === 'Ninguno') puntaje += 50;
  else if(relacion === 'no_relacionado') puntaje += 35;
  if(!dolorEnAumento) puntaje += 25; // incluye el caso "sin control anterior para comparar"
  if(evidenciaMejoria) puntaje += 25;

  // ── Alertas (texto informativo, no vuelven a tocar el puntaje) ──
  let alertas = [];
  if(evento !== 'Ninguno' && relacion === 'relacionado') alertas.push('EVENTO_RELACIONADO: ' + evento);
  if(dolorEnAumento) alertas.push('DOLOR_EN_AUMENTO');
  const alerta = alertas.join(' + ');

  const ahora = new Date();
  const etapa = calcularEtapaPostQuirurgica_(d.fechaCirugia||'', ahora);

  sheet.appendRow([
    ahora.toISOString(), dni, d.codigoCaso||'', mesControl, etapa, d.fechaCirugia||'',
    evento, relacion, d.fechaEvento||'', dolorEva ?? '', evidenciaMejoria ? 'Sí' : 'No',
    d.comentario||'', puntaje, alerta, JSON.stringify(d)
  ]);

  // Todo paciente operado queda "propio" automáticamente (decisión del Dr.
  // Muratori) — se detecta acá porque este es el primer control post-
  // quirúrgico real que se le carga, y eso solo pasa si hubo cirugía.
  let membresiaOtorgada = false;
  if(esPrimerControl){
    try{
      const resMemb = setMembresiaPaciente({DNI: dni, Membresia: true});
      membresiaOtorgada = !!(resMemb && resMemb.success);
    }catch(e){ /* si falla, no bloquea el guardado del control en sí */ }
  }

  return {success:true, message:'Control post-quirúrgico registrado', alerta: alerta || null, etapa, puntajeMes: puntaje, membresiaOtorgada};
}

function getSeguimientoPosqx(dni){
  if(!dni) return {success:true, data:[], puntuacionSeguridad:null};
  const sheet = hojaSeguimientoPosquirurgico_();
  const rows = sheet.getDataRange().getValues();
  if(rows.length < 2) return {success:true, data:[], puntuacionSeguridad:null};
  const headers = rows[0];
  const idx = (nombre) => headers.indexOf(nombre);
  const iDNI=idx('DNI'), iCaso=idx('CodigoCaso'), iFecha=idx('Timestamp'), iMes=idx('Mes_Control'), iEtapa=idx('Etapa_Postop'),
    iFechaCx=idx('Fecha_Cirugia'), iEvento=idx('Evento'), iRelacion=idx('Relacion'), iFechaEv=idx('Fecha_Evento'),
    iDolor=idx('Dolor_EVA'), iMejoria=idx('Evidencia_Mejoria'), iComentario=idx('Comentario'),
    iPuntaje=idx('Puntaje_Mes'), iAlerta=idx('Alerta');
  const data = [];
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][iDNI]) !== String(dni)) continue;
    data.push({
      fecha: rows[i][iFecha], codigoCaso: rows[i][iCaso], mesControl: rows[i][iMes], etapa: rows[i][iEtapa],
      fechaCirugia: rows[i][iFechaCx], evento: rows[i][iEvento], relacion: rows[i][iRelacion], fechaEvento: rows[i][iFechaEv],
      dolorEva: rows[i][iDolor], evidenciaMejoria: rows[i][iMejoria], comentario: rows[i][iComentario],
      puntajeMes: rows[i][iPuntaje], alerta: rows[i][iAlerta]
    });
  }
  data.sort((a,b) => Number(a.mesControl) - Number(b.mesControl));
  const puntuacionSeguridad = data.length ? Math.round(data.reduce((acc,r)=>acc+Number(r.puntajeMes||0),0) / data.length) : null;
  return {success:true, data, puntuacionSeguridad};
}

// ── Para el dashboard de Quirófano: la puntuación de seguridad promedio
// de CADA paciente que tiene al menos un control cargado, ordenada peor
// primero, para que el médico vea de entrada quién necesita atención. Mismo
// criterio de "lo más urgente arriba, sin importar la fecha" que ya usa
// getMensajesPendientes con el riesgo del QR.
function getResumenSeguridadQuirofano(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = hojaSeguimientoPosquirurgico_();
  const rows = sheet.getDataRange().getValues();
  if(rows.length < 2) return {success:true, data:[]};
  const headers = rows[0];
  const idx = (nombre) => headers.indexOf(nombre);
  const iDNI=idx('DNI'), iFecha=idx('Timestamp'), iMes=idx('Mes_Control'), iEtapa=idx('Etapa_Postop'),
    iAlerta=idx('Alerta'), iPuntaje=idx('Puntaje_Mes');

  // Agrupar por DNI: último control (para fecha/mes/etapa/alerta) +
  // promedio de Puntaje_Mes de todos sus controles (para la puntuación).
  const porDni = {};
  for(let i=1;i<rows.length;i++){
    const dni = String(rows[i][iDNI]);
    if(!porDni[dni]) porDni[dni] = {dni, puntajes:[], ultimo:null};
    porDni[dni].puntajes.push(Number(rows[i][iPuntaje]||0));
    porDni[dni].ultimo = {fecha: rows[i][iFecha], mesControl: rows[i][iMes], etapa: rows[i][iEtapa], alerta: rows[i][iAlerta]};
  }

  // Completar nombre del paciente desde la hoja Pacientes
  const pSheet = ss.getSheetByName('Pacientes');
  let pApellido={}, pNombre={};
  if(pSheet){
    const pData = pSheet.getDataRange().getValues();
    const pHeaders = pData[0];
    const dniCol=pHeaders.indexOf('DNI'), apCol=pHeaders.indexOf('Apellido'), nomCol=pHeaders.indexOf('Nombre');
    pData.slice(1).forEach(row=>{ pApellido[String(row[dniCol])] = row[apCol]; pNombre[String(row[dniCol])] = row[nomCol]; });
  }

  const data = Object.values(porDni).map(r => {
    const puntuacionSeguridad = Math.round(r.puntajes.reduce((a,b)=>a+b,0) / r.puntajes.length);
    const semaforo = puntuacionSeguridad>=85?'verde':(puntuacionSeguridad>=60?'amarillo':'rojo');
    return {
      dni: r.dni, apellido: pApellido[r.dni]||'', nombre: pNombre[r.dni]||'',
      fecha: r.ultimo.fecha, mesControl: r.ultimo.mesControl, etapa: r.ultimo.etapa, alerta: r.ultimo.alerta,
      puntuacionSeguridad, semaforo
    };
  });

  const rangoSemaforo_ = s => s==='rojo'?0:(s==='amarillo'?1:2);
  data.sort((a,b) => rangoSemaforo_(a.semaforo) - rangoSemaforo_(b.semaforo) || new Date(b.fecha) - new Date(a.fecha));
  return {success:true, data};
}

// Link + mail para que el paciente operado complete su control mensual —
// mismo patrón que enviarLinkRMYTAPaciente_.
function getSeguimientoPosqxUrl_(){
  return baseSimm_() + 'SSQ_Seguimiento_Posquirurgico.html'; // ⚠️ archivo todavía no construido
}
function construirLinkSeguimientoPosqx_(p){
  return getSeguimientoPosqxUrl_() + '?dni=' + encodeURIComponent(p.dni||'') +
    '&mes=' + encodeURIComponent(p.mesControl||'') +
    '&apellido=' + encodeURIComponent(p.apellido||'') +
    '&nombre=' + encodeURIComponent(p.nombre||'');
}
function enviarLinkSeguimientoPosqxAPaciente_(p){
  if(!p || !p.email) throw new Error('Falta el mail del paciente');
  const link = construirLinkSeguimientoPosqx_(p);
  const cuerpo =
    'Hola ' + (p.nombre||'') + ' ' + (p.apellido||'') + ',\n\n' +
    'Como parte de tu control post-quirúrgico (mes ' + (p.mesControl||'') + ' de seguimiento) te pedimos que completes este breve control — te toma un par de minutos.\n\n' +
    'Por favor, entrá al siguiente link:\n' + link + '\n\n' +
    'Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943';
  MailApp.sendEmail(p.email, 'Control post-quirúrgico — mes ' + (p.mesControl||''), cuerpo, {
    replyTo:'drluismuratori@gmail.com', name:'Dr. Luis Alberto Muratori'
  });
}

// ══════════════════════════════════════════════════════════════════
// PARCHE — Canal de comunicación de Kinesiología (KSM)
// ══════════════════════════════════════════════════════════════════
// Mail del consultorio a donde llegan los mensajes/informes del sistema.
const MEDICO_MAIL_KINESIOLOGIA = 'consultas.sism@gmail.com';

function sheetPreguntasKine_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Preguntas_Kinesiologia');
  if(!sheet){
    sheet = ss.insertSheet('Preguntas_Kinesiologia');
    sheet.appendRow(['Timestamp','DNI','Paciente','Pregunta','Respuesta','Fecha_Respuesta','Estado']);
    sheet.getRange(1,1,1,7).setBackground('#166534').setFontColor('white').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// El paciente pregunta algo desde KSM_Autoevaluacion_Paciente.html
function guardarPreguntaKine(d){
  if(!d.dni || !d.pregunta) return{success:false, error:'Falta DNI o pregunta'};
  const sheet = sheetPreguntasKine_();
  sheet.appendRow([new Date().toISOString(), d.dni, d.paciente||'', d.pregunta, '', '', 'pendiente']);
  return{success:true, message:'Pregunta enviada al kinesiólogo/a'};
}

// KSM_Kinesiologia.html pide las preguntas de un paciente (o todas si no
// se manda DNI) para mostrarlas en el panel del profesional.
function getPreguntasKine(dni){
  const sheet = sheetPreguntasKine_();
  const data = sheet.getDataRange().getValues();
  if(data.length<2) return{success:true, preguntas:[]};
  const headers = data[0];
  let preguntas = data.slice(1).map((row,i)=>{
    const obj = {_row:i+2};
    headers.forEach((h,j)=>{ obj[h]=row[j]; });
    return obj;
  });
  if(dni) preguntas = preguntas.filter(p=> String(p.DNI)===String(dni));
  return{success:true, preguntas:preguntas};
}

// El kinesiólogo responde desde KSM_Kinesiologia.html — se guarda la
// respuesta y se le manda un mail al paciente (nunca WhatsApp, regla de
// oro del sistema). El replyTo queda con el mail del kinesiólogo, para
// que si el paciente contesta, le llegue a él/ella directo.
function responderPreguntaKine(d){
  const fila = parseInt(d.fila);
  if(!fila || !d.respuesta) return{success:false, error:'Falta la fila o la respuesta'};
  const sheet = sheetPreguntasKine_();
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const filaValores = sheet.getRange(fila,1,1,headers.length).getValues()[0];
  const dni = filaValores[headers.indexOf('DNI')];
  const paciente = filaValores[headers.indexOf('Paciente')];
  const pregunta = filaValores[headers.indexOf('Pregunta')];

  sheet.getRange(fila, headers.indexOf('Respuesta')+1).setValue(d.respuesta);
  sheet.getRange(fila, headers.indexOf('Fecha_Respuesta')+1).setValue(new Date().toISOString());
  sheet.getRange(fila, headers.indexOf('Estado')+1).setValue('respondida');

  // Buscar el mail del paciente en Pacientes por DNI
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = ss.getSheetByName('Pacientes');
  let emailPaciente = '';
  if(pSheet){
    const pData = pSheet.getDataRange().getValues();
    const pHeaders = pData[0];
    const dniCol = pHeaders.indexOf('DNI'), emailCol = pHeaders.indexOf('Email');
    const fila2 = pData.find(row=> String(row[dniCol])===String(dni));
    if(fila2) emailPaciente = fila2[emailCol];
  }

  if(emailPaciente){
    const cuerpo = 'Hola ' + (paciente||'') + ',\n\n' +
      'Tu kinesiólogo/a respondió tu consulta:\n\n' +
      '"' + pregunta + '"\n\n' +
      'Respuesta:\n' + d.respuesta + '\n\n' +
      'Consultorio Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943';
    try{
      MailApp.sendEmail(emailPaciente, 'Respuesta de tu kinesiólogo/a', cuerpo, {
        replyTo: d.kinesiologoEmail || MEDICO_MAIL_KINESIOLOGIA,
        name: d.kinesiologoNombre ? (d.kinesiologoNombre + ' — Consultorio Dr. Muratori') : 'Consultorio Dr. Muratori'
      });
    }catch(e){
      return{success:true, message:'Respuesta guardada, pero no se pudo enviar el mail: ' + e};
    }
  }
  return{success:true, message:'Respuesta guardada y enviada'};
}

// El kinesiólogo manda a KSM el resumen de adherencia/progreso ya armado
// (KSM_Kinesiologia.html arma el texto con las evoluciones + autoevaluaciones
// que ya trae get_historial) y esta función lo hace llegar al médico.
function enviarInformeKine(d){
  if(!d.resumen || !d.kinesiologoEmail) return{success:false, error:'Falta el resumen o el mail del kinesiólogo'};
  try{
    MailApp.sendEmail(MEDICO_MAIL_KINESIOLOGIA,
      '[KSM] Informe de kinesiología — ' + (d.paciente||d.dni||''),
      d.resumen,
      { replyTo: d.kinesiologoEmail, name: (d.kinesiologoNombre||'Kinesiología') + ' — KSM' }
    );
    return{success:true, message:'Informe enviado al médico'};
  }catch(e){
    return{success:false, error:'No se pudo enviar el informe: ' + e};
  }
}

// ══════════════════════════════════════════════════════════════════
// PARCHE — Peaje de uso del sistema (IA + desarrollo continuo)
// ══════════════════════════════════════════════════════════════════
function getPeajeSistemaPct(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Configuracion');
  if(!sh) return{success:true, porcentaje:5};
  const rows = sh.getDataRange().getValues();
  for(const row of rows){
    if(row[0]==='peaje_sistema_pct'){
      const n = Number(row[1]);
      return{success:true, porcentaje: isNaN(n) ? 5 : n};
    }
  }
  return{success:true, porcentaje:5};
}

function setPeajeSistemaPct(d){
  const pct = Number(d.porcentaje);
  if(isNaN(pct) || pct<0 || pct>100) return{success:false, error:'Porcentaje inválido'};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Configuracion');
  if(!sh){ sh = ss.insertSheet('Configuracion'); sh.appendRow(['clave','valor']); sh.setFrozenRows(1); }
  const rows = sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(rows[i][0]==='peaje_sistema_pct'){ sh.getRange(i+1,2).setValue(pct); return{success:true}; }
  }
  sh.appendRow(['peaje_sistema_pct', pct]);
  return{success:true};
}

// ══════════════════════════════════════════════════════════════════
// PARCHE — Autorización y permisos por especialista (generalizado a
// TODOS los profesionales, no solo kinesiología: "colegas" incluidos).
//
// Dos pasos separados, a propósito:
//   1) El PROFESIONAL acepta el % de peaje (registrarAceptacionKSM, como
//      ya existía) — su consentimiento a la comisión del sistema.
//   2) El MÉDICO autoriza y elige, por especialista, qué módulos ve
//      (setPermisosEspecialista) — su decisión de qué se comparte con
//      quién. Nadie ve nada del punto 2 hasta que el médico lo habilita,
//      aunque ya haya aceptado el peaje del punto 1.
//
// Podoscopía queda afuera a propósito: es de uso EXCLUSIVO del médico y
// nunca se ofrece como módulo compartible, ni aunque se autorice todo lo
// demás (ver MODULOS_COMPARTIBLES más abajo).
// ══════════════════════════════════════════════════════════════════
const MODULOS_COMPARTIBLES = ['investigacion','cud','grafica','plantillas'];

function sheetEspecialistas_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Kinesiologos_KSM');
  if(!sh){
    sh = ss.insertSheet('Kinesiologos_KSM');
    sh.appendRow(['Timestamp','Nombre','Matricula','Email','Porcentaje_Aceptado']);
    sh.getRange(1,1,1,5).setBackground('#166534').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  asegurarColumnasEspecialistas_(sh);
  return sh;
}
// Agrega las columnas nuevas del sistema de permisos a la hoja ya
// existente (Kinesiologos_KSM) sin tocar los registros que ya había.
function asegurarColumnasEspecialistas_(sh){
  let headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const faltantes = ['Tipo','Autorizado','Modulos_JSON','Porcentaje_Personalizado','Fecha_Actualizacion_Permisos'].filter(h=>headers.indexOf(h)===-1);
  faltantes.forEach(h=>{
    sh.getRange(1, sh.getLastColumn()+1).setValue(h).setBackground('#166534').setFontColor('white').setFontWeight('bold');
  });
  return sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
}

// El profesional (kinesiólogo o colega) acepta el % de peaje vigente y
// queda registrado — "solo participa quien está de acuerdo". Si ya
// estaba registrado, actualiza sus datos sin pisar los permisos que el
// médico ya le haya otorgado.
function registrarAceptacionKSM(d){
  if(!d.email || !d.nombre) return{success:false, error:'Falta nombre o mail'};
  const sh = sheetEspecialistas_();
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  for(let i=1;i<data.length;i++){
    if(String(data[i][emailCol]).toLowerCase()===String(d.email).toLowerCase()){
      const fila = i+1;
      sh.getRange(fila, headers.indexOf('Timestamp')+1).setValue(new Date().toISOString());
      sh.getRange(fila, headers.indexOf('Nombre')+1).setValue(d.nombre);
      if(d.matricula) sh.getRange(fila, headers.indexOf('Matricula')+1).setValue(d.matricula);
      if(d.porcentaje) sh.getRange(fila, headers.indexOf('Porcentaje_Aceptado')+1).setValue(d.porcentaje);
      if(d.tipo) sh.getRange(fila, headers.indexOf('Tipo')+1).setValue(d.tipo);
      return{success:true, message:'Registro actualizado', yaExistia:true};
    }
  }
  const fila = new Array(headers.length).fill('');
  const set=(nombreCol,valor)=>{const i=headers.indexOf(nombreCol);if(i>-1)fila[i]=valor;};
  set('Timestamp', new Date().toISOString());
  set('Nombre', d.nombre);
  set('Matricula', d.matricula||'');
  set('Email', d.email);
  set('Porcentaje_Aceptado', d.porcentaje||'');
  set('Tipo', d.tipo||'kinesiologia');
  set('Autorizado', 'PENDIENTE');
  set('Modulos_JSON', '[]');
  sh.appendRow(fila);
  return{success:true, message:'Registro nuevo — queda pendiente de que el médico habilite sus módulos', yaExistia:false};
}

// Lo que consulta cada app (KSM, futuras apps de colegas) para saber qué
// puede mostrar de un profesional puntual.
function getPermisosEspecialista(email){
  if(!email) return{success:false, error:'Falta email'};
  const sh = sheetEspecialistas_();
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  for(let i=1;i<data.length;i++){
    if(String(data[i][emailCol]).toLowerCase()===String(email).toLowerCase()){
      let modulos = [];
      try{ modulos = JSON.parse(data[i][headers.indexOf('Modulos_JSON')]||'[]'); }catch(_){}
      const pctPersonal = data[i][headers.indexOf('Porcentaje_Personalizado')];
      return{
        success:true, registrado:true,
        autorizado: String(data[i][headers.indexOf('Autorizado')]).toUpperCase()==='SI',
        modulos: modulos,
        tipo: data[i][headers.indexOf('Tipo')]||'',
        porcentaje: (pctPersonal!=='' && pctPersonal!=null) ? Number(pctPersonal) : getPeajeSistemaPct().porcentaje
      };
    }
  }
  return{success:true, registrado:false, autorizado:false, modulos:[], tipo:'', porcentaje:getPeajeSistemaPct().porcentaje};
}

// Lo usa el MÉDICO (desde el panel de administración de especialistas)
// para autorizar y elegir módulos por profesional. 'podoscopia' se filtra
// siempre, aunque venga en la lista — no es negociable desde acá.
function setPermisosEspecialista(d){
  if(!d.email) return{success:false, error:'Falta email'};
  const sh = sheetEspecialistas_();
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  const modulos = (Array.isArray(d.modulos)?d.modulos:[]).filter(m=>MODULOS_COMPARTIBLES.indexOf(m)>-1);
  for(let i=1;i<data.length;i++){
    if(String(data[i][emailCol]).toLowerCase()===String(d.email).toLowerCase()){
      const fila = i+1;
      sh.getRange(fila, headers.indexOf('Autorizado')+1).setValue(d.autorizado?'SI':'NO');
      sh.getRange(fila, headers.indexOf('Modulos_JSON')+1).setValue(JSON.stringify(modulos));
      if(d.porcentajePersonalizado!==undefined && d.porcentajePersonalizado!==null && d.porcentajePersonalizado!==''){
        sh.getRange(fila, headers.indexOf('Porcentaje_Personalizado')+1).setValue(Number(d.porcentajePersonalizado));
      }
      if(d.tipo) sh.getRange(fila, headers.indexOf('Tipo')+1).setValue(d.tipo);
      sh.getRange(fila, headers.indexOf('Fecha_Actualizacion_Permisos')+1).setValue(new Date().toISOString());
      return{success:true};
    }
  }
  return{success:false, error:'No se encontró ningún especialista registrado con ese mail — primero tiene que entrar una vez a su app (KSM u otra) para registrarse y aceptar el peaje.'};
}

// Lista completa para el panel de administración del médico.
function getListaEspecialistas(){
  const sh = sheetEspecialistas_();
  const data = sh.getDataRange().getValues();
  if(data.length<2) return{success:true, especialistas:[]};
  const headers = data[0];
  const especialistas = data.slice(1).map(row=>{
    const obj={}; headers.forEach((h,i)=>{obj[h]=row[i];}); return obj;
  }).filter(e=>e.Email);
  return{success:true, especialistas};
}

// Compuerta única: todo endpoint que sirva un módulo compartible debe
// pasar por acá antes de devolver datos. 'podoscopia' siempre da false.
function tieneModuloAutorizado_(email, modulo){
  if(modulo==='podoscopia') return false;
  if(!email) return false;
  const p = getPermisosEspecialista(email);
  return !!(p.autorizado && p.modulos && p.modulos.indexOf(modulo)>-1);
}

function sheetFacturacionSistema_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Facturacion_Sistema');
  if(!sh){
    sh = ss.insertSheet('Facturacion_Sistema');
    sh.appendRow(['Timestamp','Origen','DNI','Paciente','Profesional_Nombre','Profesional_Email',
                  'Monto','Comision_Pct','Comision_Monto','Neto_Profesional','Mes_Referencia','Estado_Comision']);
    sh.getRange(1,1,1,12).setBackground('#0EA5E9').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// Función compartida — la usan tanto KSM (kinesiólogos) como Consultas
// Online (el propio médico) para registrar el peaje de una sesión/consulta
// paga. origen = 'kinesiologia' o 'consulta_medica'.
function registrarFacturacionSistema_(origen, dni, paciente, profesionalNombre, profesionalEmail, monto){
  if(!monto || monto<=0) return null;
  // Si el profesional tiene un % personalizado cargado (setPermisosEspecialista),
  // se usa ese; si no, el % global del sistema.
  let pct = getPeajeSistemaPct().porcentaje;
  if(profesionalEmail){
    const permisos = getPermisosEspecialista(profesionalEmail);
    if(permisos.registrado && !isNaN(Number(permisos.porcentaje))) pct = Number(permisos.porcentaje);
  }
  const comision = Math.round(monto * pct / 100 * 100) / 100;
  const neto = Math.round((monto - comision) * 100) / 100;
  const mesRef = new Date().toISOString().slice(0,7); // "2026-08"
  sheetFacturacionSistema_().appendRow([
    new Date().toISOString(), origen, dni||'', paciente||'',
    profesionalNombre||'', profesionalEmail||'',
    monto, pct, comision, neto, mesRef, 'pendiente'
  ]);
  return{comision, neto, porcentaje:pct};
}

// Endpoint que llama KSM_Kinesiologia.html cuando el kinesiólogo marca
// una sesión como paga.
function registrarFacturacionSistema(d){
  const r = registrarFacturacionSistema_('kinesiologia', d.dni, d.paciente, d.profesionalNombre||d.kinesiologoNombre, d.profesionalEmail||d.kinesiologoEmail, Number(d.monto));
  if(!r) return{success:false, error:'Monto inválido'};
  return Object.assign({success:true}, r);
}

// Resumen del mes — de un profesional puntual, o de todos si no se pasa
// mail (útil para vos: ver de un vistazo el peaje total del mes, sumando
// kinesiólogos + tus propias Consultas Online).
function getResumenFacturacionSistema(profesionalEmail, mes, origen){
  const sh = sheetFacturacionSistema_();
  const data = sh.getDataRange().getValues();
  if(data.length<2) return{success:true, filas:[], totalFacturado:0, totalComision:0, totalNeto:0};
  const headers = data[0];
  const mesRef = mes || new Date().toISOString().slice(0,7);
  let filas = data.slice(1).map(row=>{
    const obj={}; headers.forEach((h,i)=>{obj[h]=row[i];}); return obj;
  }).filter(f=> f.Mes_Referencia===mesRef);
  if(profesionalEmail) filas = filas.filter(f=> f.Profesional_Email===profesionalEmail);
  if(origen) filas = filas.filter(f=> f.Origen===origen);
  const totalFacturado = filas.reduce((s,f)=> s+Number(f.Monto||0), 0);
  const totalComision = filas.reduce((s,f)=> s+Number(f.Comision_Monto||0), 0);
  const totalNeto = filas.reduce((s,f)=> s+Number(f.Neto_Profesional||0), 0);
  return{success:true, filas:filas, mes:mesRef, totalFacturado:totalFacturado, totalComision:totalComision, totalNeto:totalNeto};
}

function enviarResumenFacturacionSistema(d){
  if(!d.profesionalEmail) return{success:false, error:'Falta el mail del profesional'};
  const resumen = getResumenFacturacionSistema(d.profesionalEmail, d.mes||'', '');
  const cuerpo =
    'Resumen de facturación del sistema — ' + (d.profesionalNombre||d.profesionalEmail) + '\n' +
    'Mes: ' + resumen.mes + '\n\n' +
    'Sesiones/consultas facturadas: ' + resumen.filas.length + '\n' +
    'Total facturado: $' + resumen.totalFacturado.toLocaleString('es-AR') + '\n' +
    'Peaje del sistema a transferir: $' + resumen.totalComision.toLocaleString('es-AR') + '\n' +
    'Neto del profesional: $' + resumen.totalNeto.toLocaleString('es-AR');
  try{
    MailApp.sendEmail(MEDICO_MAIL_KINESIOLOGIA,
      '[Sistema] Resumen de facturación — ' + (d.profesionalNombre||d.profesionalEmail) + ' — ' + resumen.mes,
      cuerpo,
      { replyTo: d.profesionalEmail, name: (d.profesionalNombre||'Sistema') + ' — SIM-M' }
    );
    return{success:true};
  }catch(e){
    return{success:false, error:'No se pudo enviar: ' + e};
  }
}

// ── actualizarConsulta (versión final — reemplaza la original) ──
// Único agregado respecto a la versión de siempre: cuando se marca una
// Consulta Online como respondida (respuesta_final), registra sola el
// peaje del sistema usando el precio que ya tenía cargado esa consulta.
function actualizarConsulta(d){
  const sh=sheetConsultas_();
  const rows=sh.getDataRange().getValues();
  const headers=rows[0];
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][0])===String(d.id)){
      if(d.estado!==undefined)   sh.getRange(i+1,headers.indexOf('Estado')+1).setValue(d.estado);
      if(d.payment_id)           sh.getRange(i+1,headers.indexOf('Payment_ID')+1).setValue(d.payment_id);
      if(d.respuesta_draft)      sh.getRange(i+1,headers.indexOf('Respuesta_Draft')+1).setValue(d.respuesta_draft);
      if(d.respuesta_final)      sh.getRange(i+1,headers.indexOf('Respuesta_Final')+1).setValue(d.respuesta_final);
      if(d.respuesta_final){
        sh.getRange(i+1,headers.indexOf('Fecha_Respuesta')+1).setValue(new Date().toISOString());
        try{
          registrarFacturacionSistema_(
            'consulta_medica',
            rows[i][headers.indexOf('DNI')],
            rows[i][headers.indexOf('Paciente')],
            'Dr. Luis Alberto Muratori',
            MEDICO_MAIL_KINESIOLOGIA,
            Number(rows[i][headers.indexOf('Precio')])
          );
        }catch(e){ Logger.log('No se pudo registrar el peaje de la consulta: ' + e); }
      }
      return{success:true};
    }
  }
  return{success:false,error:'Consulta no encontrada'};
}

// ══════════════════════════════════════════════════════════════════
// PARCHE — Procesador de Mensajes / Agenda de Correspondencia
// (versión optimizada en consumo de tokens — thinking al mínimo, cuerpo
// del mail limpio de citas antes de mandarlo a la IA, y log de tokens
// gastados en Log_Tokens_IA)
// ══════════════════════════════════════════════════════════════════
// SEGURIDAD (corregido en esta revisión): la clave de Gemini ya NO vive
// como texto plano acá — por eso había terminado expuesta antes. Ahora se
// guarda en las Propiedades del script, igual que ya se hacía con Mercado
// Pago y Claude. Para cargar la clave nueva (rotación):
//   1) Pegarla en configurarGeminiAPI() más abajo, donde dice PEGÁ_AQUÍ.
//   2) Elegir configurarGeminiAPI en el desplegable de arriba del editor.
//   3) Tocar ▶ Ejecutar UNA vez.
//   4) Borrar la clave de acá adentro si querés (ya quedó guardada, no hace falta de nuevo).
function configurarGeminiAPI(){
  const key = 'PEGÁ_AQUÍ_TU_GEMINI_API_KEY_NUEVA';
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
  Logger.log('Gemini API key guardada correctamente en las Propiedades del script.');
}
function getGeminiUrl_(){
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if(!key) throw new Error('Falta configurar GEMINI_API_KEY — ejecutar configurarGeminiAPI() una vez desde el editor.');
  return 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key=' + key;
}
const LABEL_PROCESADO = 'SIM-M/Procesado';
const LABEL_ERROR = 'SIM-M/Error-procesamiento';
const PALABRAS_SENSIBLES = ['discapacidad','cud','certificado único','pnc','invalidez','hospital','pensión no contributiva'];

// Saca las cadenas de mails anteriores citados (que solo suman tokens de
// entrada sin aportar información nueva) antes de mandar el texto a la IA.
// No modifica el cuerpo original que se guarda en la planilla.
function limpiarCuerpoMail_(texto){
  if(!texto) return '';
  const marcadoresDeCita = [
    /El .{0,60}escribi[oó]:/i,
    /-----\s*Mensaje original\s*-----/i,
    /-----\s*Original Message\s*-----/i,
    /^>.*$/m,
    /De:\s*.+\n\s*Enviado:/i,
    /From:\s*.+\n\s*Sent:/i
  ];
  let corte = texto.length;
  marcadoresDeCita.forEach(regex=>{
    const m = texto.match(regex);
    if(m && typeof m.index === 'number' && m.index < corte) corte = m.index;
  });
  return texto.slice(0, corte).trim();
}

// ══════════════════════════════════════════════════════════════════
// RSP — Resumen de Seguimiento del Paciente para la IA (sep 2026)
// No es una pantalla aparte: es el contexto clínico (parámetros de RMYT +
// resumen breve de Historia_Clínica) que se arma por DNI y se le pasa a
// analizarConIA_ para que el borrador de respuesta a un mail del paciente
// esté fundado en su caso real, no en el mail suelto sin contexto.
// ══════════════════════════════════════════════════════════════════

// Busca el DNI de un paciente a partir de su dirección de mail, en la
// hoja Pacientes (mismo patrón inverso que ya usa responderPreguntaKine,
// que busca el mail a partir del DNI).
function buscarDniPorEmail_(email){
  if(!email) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = ss.getSheetByName('Pacientes');
  if(!pSheet) return null;
  const pData = pSheet.getDataRange().getValues();
  const pHeaders = pData[0];
  const dniCol = pHeaders.indexOf('DNI'), emailCol = pHeaders.indexOf('Email');
  if(dniCol===-1 || emailCol===-1) return null;
  const emailNorm = String(email).trim().toLowerCase();
  const fila = pData.slice(1).find(row => String(row[emailCol]||'').trim().toLowerCase() === emailNorm);
  return fila ? String(fila[dniCol]) : null;
}

// Extrae la dirección de mail "pura" del From de Gmail (viene como
// 'Juan Pérez <juan@gmail.com>' o directamente 'juan@gmail.com').
function extraerEmailDe_(remitente){
  if(!remitente) return '';
  const m = String(remitente).match(/<([^<>]+)>/);
  return (m ? m[1] : remitente).trim().toLowerCase();
}

// Arma un resumen corto (para no gastar de más en tokens de Gemini) con
// lo último cargado de RMYT_Autoevaluaciones + lo último de Historia_Clínica,
// para ese DNI. Devuelve '' si no hay nada — analizarConIA_ simplemente no
// agrega la sección al prompt en ese caso (paciente sin historia previa,
// o mail que no matcheó con ningún DNI conocido).
function armarResumenClinicoParaIA_(dni){
  if(!dni) return '';
  const partes = [];
  try{
    const rmyt = getAutoevaluacionesRMYT(dni);
    // getAutoevaluacionesRMYT devuelve "data" ordenado MÁS RECIENTE PRIMERO
    // (ver el sort descendente al final de esa función). RSP es justamente
    // esto: un resumen SEMANA A SEMANA de lo que el paciente reporta que
    // está haciendo en su casa (adherencia a medicación/ejercicios/kinesio,
    // dolor, efectos adversos) a través del conector de parámetros del
    // RMYT — no alcanza con la última semana sola, hace falta la serie
    // reciente para que la IA vea si mejora, empeora o se estanca.
    if(rmyt && rmyt.success && rmyt.data && rmyt.data.length){
      const ULTIMAS_SEMANAS = 4;
      const recientes = rmyt.data.slice(0, ULTIMAS_SEMANAS).slice().reverse(); // cronológico ascendente para leer la evolución
      const lineas = recientes.map(r =>
        '  · ' + (r.fecha||'sin fecha') + ' — EVA reposo/movimiento ' + (r.eva_reposo ?? '—') + '/' + (r.eva_movimiento ?? '—') +
        ', recuperación percibida ' + (r.recuperacion_pct ?? '—') + '%, cumplimiento ' + (r.cumplimiento_promedio ?? '—') +
        (r.alerta ? (', ⚠️ ' + r.alerta) : '')
      );
      let bloqueTendencia = '';
      if(rmyt.tendencia && rmyt.tendencia.mensaje){
        bloqueTendencia = '- Tendencia estimada (no reemplaza el criterio médico): ' + rmyt.tendencia.mensaje + '\n';
      }
      partes.push(
        'Seguimiento semanal RMYT — últimas ' + recientes.length + ' semana(s) reportadas (segmento ' + (recientes[recientes.length-1].segmento||'—') + '), de más vieja a más nueva:\n' +
        lineas.join('\n') + '\n' +
        bloqueTendencia
      );
    }
  }catch(e){ Logger.log('armarResumenClinicoParaIA_: no se pudo leer RMYT — ' + e); }
  try{
    const hist = getHistorialPaciente(dni);
    if(hist && hist.success && hist.historial && hist.historial.length){
      const parseFecha = f => { const [d,m,y] = String(f||'').split('/'); return new Date(y+'-'+m+'-'+d); };
      const ultimo = hist.historial.slice().sort((a,b)=>parseFecha(b.Fecha)-parseFecha(a.Fecha))[0];
      partes.push(
        'Última consulta registrada en Historia Clínica (' + (ultimo.Fecha||'—') + '):\n' +
        '- Diagnóstico: ' + (ultimo.Diagnostico||'—') + '\n' +
        (ultimo.Indicaciones ? ('- Indicaciones: ' + ultimo.Indicaciones + '\n') : '') +
        (ultimo.Cirugia_Programada ? ('- Cirugía programada: ' + (ultimo.Tipo_Procedimiento||'') + ' (' + (ultimo.Fecha_Cirugia_Programada||'sin fecha') + ')\n') : '')
      );
    }
  }catch(e){ Logger.log('armarResumenClinicoParaIA_: no se pudo leer Historia_Clinica — ' + e); }
  return partes.join('\n');
}

// Registra cuántos tokens gastó cada llamada a Gemini (input, pensamiento,
// output, total), para poder ver el consumo real con el tiempo.
function registrarUsoTokens_(funcion, modelo, usage){
  if(!usage) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Log_Tokens_IA');
  if(!sheet){
    sheet = ss.insertSheet('Log_Tokens_IA');
    sheet.appendRow(['Timestamp','Funcion','Modelo','Tokens_Input','Tokens_Pensamiento','Tokens_Output','Tokens_Total']);
    sheet.getRange(1,1,1,7).setBackground('#0c4a6e').setFontColor('white').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    new Date().toISOString(), funcion, modelo,
    usage.promptTokenCount||0, usage.thoughtsTokenCount||0,
    usage.candidatesTokenCount||0, usage.totalTokenCount||0
  ]);
}

const LABEL_NO_MEDICO = 'SIM-M/No-medico';
// Remitentes/dominios que NUNCA son consultas médicas (redes, compras, publicidad,
// suscripciones). Se puede ampliar sin tocar el código con la propiedad de script
// MAILS_BLOQUEADOS (lista separada por comas). Y la propiedad MAILS_PERMITIDOS
// (comas) fuerza que un remitente/dominio SIEMPRE se analice (colegas, obras sociales).
const REMITENTES_NO_MEDICOS_ = ['facebook','instagram','whatsapp','linkedin','twitter','tiktok','pinterest','youtube','aliexpress','alibaba','temu','shein','amazon','spotify','netflix','ole.com.ar','orthobullets','paypal','booking','despegar','uber','rappi','pedidosya','mailchimp','sendgrid','news','newsletter','noreply','no-reply','donotreply','do-not-reply','mailer-daemon','postmaster','marketing','promo','accounts.google','google.com','github','cloudflare','canva','udemy','coursera','zoom.us','eventbrite','clubln','lanacion','clarin','clubdeinversores','cartafinanciera','coleccionables','suscribite','mercadolibre','ahumada','ypf','worldhistory','serviclub','imed.cl','centromedicopalmares','neuromed','centromedicopa','canonrecuperogastos','administracioncmp','cajasalud','comunicacionesinternas'];
function listaProp_(nombre){
  try{
    const v = PropertiesService.getScriptProperties().getProperty(nombre) || '';
    return v.split(',').map(function(x){return x.trim().toLowerCase();}).filter(Boolean);
  }catch(e){ return []; }
}
// true si el mail lo mandó el propio sistema/el dueño (copias, avisos, pruebas) o su
// asunto empieza con "SIM-M". Se evalúa SIEMPRE, aunque la dirección figure como
// paciente en la planilla (el médico puede estar cargado como paciente de prueba).
function esMailPropioOSistema_(remitente, asunto){
  if(/^\s*SIM-M\b/i.test(String(asunto||''))) return true;
  const r = String(remitente||'').toLowerCase();
  const permitidos = listaProp_('MAILS_PERMITIDOS');
  if(permitidos.some(function(x){ return r.indexOf(x) !== -1; })) return false;
  if(/luis\s+(alberto\s+)?muratori/.test(r)) return true;   // el propio médico desde otra casilla
  try{
    const yo = String(Session.getEffectiveUser().getEmail()||'').toLowerCase();
    if(yo && r.indexOf(yo) !== -1) return true;
  }catch(e){}
  if(r.indexOf(MAIL_CONSULTAS_) !== -1) return true;   // mails que manda el propio sistema desde consultas.sism
  return false;
}
// true si el remitente es claramente NO médico (por reglas, sin IA).
function esRemitenteNoMedico_(remitente, rawCabecera, asunto){
  const r = String(remitente||'').toLowerCase();
  if(/^\s*SIM-M\b/i.test(String(asunto||''))) return true;   // avisos automáticos del propio sistema (token, etc.)
  const permitidos = listaProp_('MAILS_PERMITIDOS');
  if(permitidos.some(function(x){ return r.indexOf(x) !== -1; })) return false;
  // Mails que el propio sistema se manda al dueño (copias, avisos, pruebas): no son consultas.
  try{
    const yo = String(Session.getEffectiveUser().getEmail()||'').toLowerCase();
    if(yo && r.indexOf(yo) !== -1) return true;
  }catch(e){}
  if(r.indexOf(MAIL_CONSULTAS_) !== -1) return true;
  const bloq = REMITENTES_NO_MEDICOS_.concat(listaProp_('MAILS_BLOQUEADOS'));
  // se compara también sin espacios/puntos/guiones: "Centro Medico Palmares" = centromedicopalmares
  const rn = r.normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[\s.\-_]/g,'');
  if(bloq.some(function(x){ return r.indexOf(x) !== -1 || rn.indexOf(x.replace(/[\s.\-_]/g,'')) !== -1; })) return true;
  // Asuntos administrativos (facturas, honorarios, circulares, recordatorios de turnos del centro)
  if(/\b(circular|factura|facturas|facturaci[oó]n|honorarios|retenci[oó]n|recupero de gastos|recordatorio|liquidaci[oó]n|expte|expediente)\b/i.test(String(asunto||''))) return true;
  if(rawCabecera && /^List-Unsubscribe:/mi.test(String(rawCabecera))) return true;   // mail masivo
  return false;
}

// Casilla de consultas del sistema: los pacientes escriben ACÁ (oct 2026).
const MAIL_CONSULTAS_ = 'consultas.sism@gmail.com';
// Si el proceso corre en la casilla personal del médico (lmuratori), solo se toman los
// mails dirigidos a consultas.sism (reenviados a esa casilla); el resto de su correo
// personal llega igual a su bandeja pero NO entra a la Agenda de Mensajes.
// Si corre dentro de la propia cuenta consultas.sism, se toma toda la bandeja.
// Se apaga con la propiedad PROCESAR_SOLO_CONSULTAS = no.
function filtroSoloConsultas_(yo){
  try{
    const modo = String(PropertiesService.getScriptProperties().getProperty('PROCESAR_SOLO_CONSULTAS') || 'si').toLowerCase();
    if(modo === 'no') return '';
    if(String(yo||'').toLowerCase() === MAIL_CONSULTAS_) return '';
  }catch(e){}
  return '(to:' + MAIL_CONSULTAS_ + ' OR cc:' + MAIL_CONSULTAS_ + ' OR deliveredto:' + MAIL_CONSULTAS_ + ') ';
}
function procesarBandejaEntrada(){
  // Solo corre con la cuenta del médico. Si otro usuario con acceso al proyecto creó
  // un activador propio, ese activador leería SU casilla y llenaría la Agenda con
  // mails ajenos: acá se ignora. Se puede cambiar con la propiedad MAIL_PROCESADOR.
  // Cuentas autorizadas (propiedad MAIL_PROCESADOR, separadas por coma). Por defecto:
  // la del médico y la casilla de consultas del sistema (consultas.sism@gmail.com).
  let yo = '';
  try{
    yo = String(Session.getEffectiveUser().getEmail()||'').toLowerCase();
    const permitidos = String(PropertiesService.getScriptProperties().getProperty('MAIL_PROCESADOR') || ('lmuratori@gmail.com,' + MAIL_CONSULTAS_)).toLowerCase().split(/[,;\s]+/).filter(Boolean);
    if(yo && permitidos.indexOf(yo) === -1){ Logger.log('procesarBandejaEntrada ignorada: corre con ' + yo + ' y solo debe correr con ' + permitidos.join(', ')); return; }
  }catch(e){}
  let label = GmailApp.getUserLabelByName(LABEL_PROCESADO);
  if(!label) label = GmailApp.createLabel(LABEL_PROCESADO);
  let labelError = GmailApp.getUserLabelByName(LABEL_ERROR);
  if(!labelError) labelError = GmailApp.createLabel(LABEL_ERROR);
  let labelNoMed = GmailApp.getUserLabelByName(LABEL_NO_MEDICO);
  if(!labelNoMed) labelNoMed = GmailApp.createLabel(LABEL_NO_MEDICO);

  const hilos = GmailApp.search(filtroSoloConsultas_(yo) + 'is:unread in:inbox -category:promotions -category:social -category:updates -category:forums -label:"' + LABEL_PROCESADO + '" -label:"' + LABEL_ERROR + '" -label:"' + LABEL_NO_MEDICO + '" -subject:"Plantillas por QR" newer_than:3d', 0, 20);
  const sheet = hojaColaMensajes_();

  hilos.forEach(hilo=>{
    try{
      const mensajes = hilo.getMessages();
      const ultimo = mensajes[mensajes.length-1];
      const remitente = ultimo.getFrom();
      const asunto = hilo.getFirstMessageSubject();
      const cuerpo = ultimo.getPlainBody().slice(0,3000);   // se guarda completo, como siempre
      const cuerpoParaIA = limpiarCuerpoMail_(cuerpo);       // versión recortada, solo para Gemini

      const textoCompleto = (asunto + ' ' + cuerpo).toLowerCase();
      const esSensible = PALABRAS_SENSIBLES.some(function(p){ return textoCompleto.includes(p); });

      // RSP — si el mail vino de un paciente que ya tenemos cargado (match
      // por su mail en la hoja Pacientes), le pasamos a la IA su resumen de
      // RMYT + última consulta, para que el borrador esté fundado en su
      // caso real y no sea una respuesta genérica.
      if(esMailPropioOSistema_(remitente, asunto)){
        hilo.addLabel(labelNoMed);
        return;
      }
      const dniDetectado = buscarDniPorEmail_(extraerEmailDe_(remitente));
      // Filtro "solo origen médico": un paciente conocido SIEMPRE pasa; si no, los
      // remitentes de publicidad/redes/compras/mail masivo se descartan sin gastar IA.
      // El mail NO se marca como leído ni se toca: queda en su casilla, solo con etiqueta.
      let rawCab = '';
      if(!dniDetectado){ try{ rawCab = String(ultimo.getRawContent()||'').slice(0,20000); }catch(e){} }
      if(!dniDetectado && esRemitenteNoMedico_(remitente, rawCab, asunto)){
        hilo.addLabel(labelNoMed);
        return;
      }
      // RSP con contexto clínico es beneficio de membresía (ver Membresia en
      // addPaciente): un paciente de solo filiación sigue recibiendo su
      // borrador de respuesta igual que cualquiera, pero sin el contexto
      // clínico enriquecido — se comporta como si no hubiera matcheado DNI.
      const contextoClinico = (dniDetectado && tienePacienteMembresia_(dniDetectado)) ? armarResumenClinicoParaIA_(dniDetectado) : '';

      // ── Membresía de Seguimiento (mensual, por UNA patología puntual) ──
      // Se resuelve ACÁ, antes de llamar a la IA, para poder pedirle a la
      // misma IA que además chequee si el mensaje realmente habla de esa
      // patología puntual — así no se convierte en un consultorio
      // sentimental/de psicología sin límite. Tope: 3 mensajes/semana (lo
      // no usado no se acumula) y 400 palabras por mensaje — ver
      // getMembresiaSeguimiento/incrementarMensajeSemanalMembresia_.
      const membresiaSeg = dniDetectado ? getMembresiaSeguimiento(dniDetectado) : null;
      const membresiaSegActiva = !!(membresiaSeg && membresiaSeg.success && membresiaSeg.existe && membresiaSeg.estado==='activa');
      const palabrasMensaje = (cuerpo.trim().match(/\S+/g)||[]).length;
      const dentroDeLargoSeg = palabrasMensaje <= 400;
      // Regla nueva: 1 mensaje cubierto cada MEMB_INTERVALO_MENSAJES_HS (8 hs), sin tope semanal
      const horasDesdeUltimoSeg = membresiaSegActiva ? membHorasDesde_(membresiaSeg.ultimo_mensaje_cubierto, new Date()) : null;
      const dentroDeCupoSeg = membresiaSegActiva && (horasDesdeUltimoSeg===null || horasDesdeUltimoSeg >= MEMB_INTERVALO_MENSAJES_HS);

      const analisis = analizarConIA_(asunto, cuerpoParaIA, esSensible, contextoClinico, membresiaSegActiva ? membresiaSeg.patologia : '', membresiaSegActiva ? ((membresiaSeg.apellido||'')+', '+(membresiaSeg.nombre||'')) : '');

      // La IA dice que no es médico y no es un paciente conocido: no entra a la Agenda.
      if(analisis.esMedico === false && !dniDetectado){
        hilo.addLabel(labelNoMed);
        return;
      }

      let cubiertoPorMembresia = false;
      let motivoFueraMembresia = '';
      if(membresiaSegActiva){
        if(!dentroDeCupoSeg){
          motivoFueraMembresia = 'Escribió antes de las '+MEMB_INTERVALO_MENSAJES_HS+' hs del último mensaje cubierto (es 1 mensaje cada '+MEMB_INTERVALO_MENSAJES_HS+' hs)';
        } else if(!dentroDeLargoSeg){
          motivoFueraMembresia = 'Supera las 400 palabras permitidas ('+palabrasMensaje+' palabras)';
        } else if(analisis.paraOtraPersona===true){
          motivoFueraMembresia = 'Parece una consulta por OTRA persona — la membresía es solo para el titular';
        } else if(!analisis.dentroDeSeguimiento){
          motivoFueraMembresia = 'No parece referirse al estado de "'+membresiaSeg.patologia+'" (la patología en seguimiento) — se trata como consulta común';
        } else {
          cubiertoPorMembresia = true;
        }
      }
      if(cubiertoPorMembresia){
        analisis.requierePago = false;
        registrarMensajeCubiertoMembresia_(dniDetectado);
      }
      // Recordatorio AUTOMÁTICO (texto fijo) si el mensaje es de otra patología u otra persona.
      // Solo si la IA devolvió el chequeo de tema de forma válida (nunca por una respuesta rota).
      let avisoEnviado = false;
      if(membresiaSegActiva && analisis.chequeoTemaValido===true && (analisis.paraOtraPersona===true || analisis.dentroDeSeguimiento===false)){
        avisoEnviado = enviarAvisoFueraDeMembresia_(remitente, dniDetectado, membresiaSeg.nombre, membresiaSeg.patologia, analisis.paraOtraPersona===true);
      }

      const id = Utilities.getUuid();
      const fila = [
        id, new Date().toISOString(), remitente, asunto, cuerpo,
        analisis.clasificacion, analisis.cantPreguntas,
        analisis.requierePago, analisis.requierePago ? 'pendiente' : 'no_aplica',
        analisis.borrador, 'pendiente_revision', '', hilo.getId()
      ];
      sheet.appendRow(fila);
      // DNI detectado — columna agregada aparte (por nombre, no rompe filas
      // viejas) para que el médico vea en Agenda de Mensajes a qué paciente
      // matcheó, sin tener que abrir la Historia Clínica.
      if(dniDetectado){
        const colDni = asegurarColumnaColaMensajes_(sheet, 'DNI_Detectado');
        sheet.getRange(sheet.getLastRow(), colDni).setValue(dniDetectado);
      }
      // NUEVO: Membresía de Seguimiento — para que Agenda de Mensajes pueda
      // mostrarle al médico si este mensaje quedó cubierto por la
      // membresía mensual del paciente (y si no, por qué no).
      const colCubierto = asegurarColumnaColaMensajes_(sheet, 'Cubierto_Membresia');
      sheet.getRange(sheet.getLastRow(), colCubierto).setValue(cubiertoPorMembresia ? 'SI' : 'NO');
      if(motivoFueraMembresia){
        const colMotivo = asegurarColumnaColaMensajes_(sheet, 'Motivo_Fuera_Membresia');
        sheet.getRange(sheet.getLastRow(), colMotivo).setValue(motivoFueraMembresia);
      }

      if(avisoEnviado){
        const colAv = asegurarColumnaColaMensajes_(sheet, 'Aviso_Automatico');
        sheet.getRange(sheet.getLastRow(), colAv).setValue('SI');
      }
      if(cubiertoPorMembresia){
        const colRe = asegurarColumnaColaMensajes_(sheet, 'Responder_Antes_De');
        sheet.getRange(sheet.getLastRow(), colRe).setValue(responderAntesDe_(new Date(), MEMB_RESPUESTA_HS).toISOString());
      }
      hilo.addLabel(label);
      hilo.markRead();
    }catch(err){
      hilo.addLabel(labelError);
      Logger.log('Error procesando hilo: ' + err);
    }
  });
}

function hojaColaMensajes_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Cola_Mensajes');
  if(!sheet){
    sheet = ss.insertSheet('Cola_Mensajes');
    sheet.appendRow(['id','fecha','remitente','asunto','cuerpo','clasificacion','cantPreguntas','requierePago','estadoPago','borrador','estado','fila_referencia','threadId']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function analizarConIA_(asunto, cuerpo, esSensible, contextoClinico, patologiaSeguimiento, nombrePaciente){
  const contextoSensible = esSensible ? 'El mensaje contiene palabras relacionadas a discapacidad/hospital público — probablemente corresponda al circuito gratuito y sensible.' : '';
  // RSP — si matcheamos el mail con un paciente conocido, este es su
  // resumen de RMYT + última consulta (ver armarResumenClinicoParaIA_).
  // Consultarlo SIEMPRE antes de armar el borrador: la respuesta tiene que
  // estar fundada en el caso real del paciente, no ser genérica.
  const bloqueContexto = contextoClinico
    ? ('CONTEXTO CLÍNICO DEL PACIENTE (consultalo antes de responder — la respuesta debe estar fundada en esto, no ser genérica):\n' + contextoClinico + '\n\n')
    : '';
  // Membresía de Seguimiento: si el paciente tiene una activa, le pedimos a
  // la IA que ADEMÁS diga si este mensaje puntual realmente habla del
  // estado/evolución de esa patología — para no dejar que la membresía se
  // convierta en un chat libre sin límite de tema (riesgo que señaló el
  // Dr. Muratori: "se puede tornar un consultorio sentimental o de
  // psicología").
  const bloqueSeguimiento = patologiaSeguimiento
    ? ('Este paciente tiene una MEMBRESÍA DE SEGUIMIENTO mensual activa, exclusiva para la patología: "' + patologiaSeguimiento + '". Además de clasificar el mensaje, indicá si este mensaje puntual se refiere al estado/evolución/síntomas de ESA patología (dentroDeSeguimiento=true) o si es otra cosa — otra dolencia, un tema administrativo no relacionado, o una conversación de tipo emocional/personal sin relación clínica con esa patología (dentroDeSeguimiento=false). Además indicá paraOtraPersona=true si quien escribe consulta por OTRA persona (hijo/a, pareja, madre, padre, amigo/a, etc.) y no por sí mismo/a; el titular de la membresía es: ' + (nombrePaciente||'(no informado)') + '. La membresía NO es un servicio de médico en línea: es solo control y seguimiento de esa patología y de ese titular. Un saludo o agradecimiento breve, o avisar que realizó/cargó un estudio de esa patología, cuenta como dentroDeSeguimiento=true.\n\n')
    : '';
  const prompt = 'Sos el asistente administrativo del Dr. Luis Alberto Muratori, médico traumatólogo (M.N. 100.540 / M.P. 9943, Mendoza, Argentina).\n' +
    'Te llegó este mail. Tu tarea es CLASIFICARLO y armar un BORRADOR de respuesta que el doctor va a revisar antes de enviar — nunca se envía sin su aprobación.\n\n' +
    contextoSensible + '\n\n' +
    bloqueContexto +
    bloqueSeguimiento +
    'ASUNTO: ' + asunto + '\n' +
    'CUERPO:\n' + cuerpo + '\n\n' +
    'Reglas:\n' +
    '- "sensible_publico": discapacidad, CUD, PNC, hospital público → gratuito, respuesta mínima y clara.\n' +
    '- "privado": consulta de consultorio privado (cualquier cantidad de preguntas) → requiere pago antes de enviarse.\n' +
    '- esMedico: true SOLO si es un mensaje de un paciente, de un familiar o de un colega sobre un problema de salud o un caso clínico: consulta, síntomas, estudios, informes o certificados de un paciente, pedido de turno o receta, plantillas ortopédicas, o respuesta a un mail del doctor. false si es administrativo o comercial: facturas, honorarios, retenciones, recupero de gastos, contratos, recordatorios de turnos de un centro médico, trámites de la clínica donde trabaja, cadenas o saludos masivos, publicidad, redes sociales, compras, suscripciones, webinars, newsletters, notificaciones automáticas, banca o spam.\n' +
    'Devolvé SOLO un JSON válido, sin texto adicional, con esta forma exacta:\n' +
    '{ "esMedico": true o false, "clasificacion": "sensible_publico o privado", "cantPreguntas": numero entero, "requierePago": true o false, "borrador": "texto completo de la respuesta sugerida, en tono profesional y calido, en español, firmada como Dr. Luis Alberto Muratori"' +
    (patologiaSeguimiento ? ', "dentroDeSeguimiento": true o false, "paraOtraPersona": true o false' : '') +
    ' }';

  try{
    const resp = UrlFetchApp.fetch(getGeminiUrl_(), {
      method:'post', contentType:'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: {
          thinkingConfig: { thinkingLevel: 'minimal' }
        }
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(resp.getContentText());
    registrarUsoTokens_('analizarConIA_', 'gemini-3.6-flash', data.usageMetadata);
    let textoIA = data.candidates[0].content.parts[0].text;
    textoIA = textoIA.replace(/```json|```/g,'').trim();
    const parsed = JSON.parse(textoIA);
    // Default seguro: si se esperaba el chequeo de tema y la IA no lo
    // devolvió (respuesta mal formada, campo faltante), nunca se cubre por
    // membresía "por las dudas" — queda como consulta común.
    parsed.chequeoTemaValido = !!(patologiaSeguimiento && typeof parsed.dentroDeSeguimiento === 'boolean' && typeof parsed.paraOtraPersona === 'boolean');
    if(patologiaSeguimiento && typeof parsed.dentroDeSeguimiento !== 'boolean') parsed.dentroDeSeguimiento = false;
    if(typeof parsed.paraOtraPersona !== 'boolean') parsed.paraOtraPersona = false;
    if(typeof parsed.esMedico !== 'boolean') parsed.esMedico = true;   // ante la duda, NO se pierde un mail médico
    return parsed;
  }catch(err){
    return {
      clasificacion: esSensible ? 'sensible_publico' : 'privado',
      cantPreguntas: (cuerpo.match(/\?/g)||[]).length || 1,
      requierePago: !esSensible,
      dentroDeSeguimiento: false,
      borrador: '[COMPLETAR] no se pudo generar el borrador automatico: ' + err
    };
  }
}

function getMensajesPendientes(){
  const sheet = hojaColaMensajes_();
  const rows = sheet.getDataRange().getValues();
  // A partir de acá se leen las columnas por NOMBRE de encabezado, no por
  // índice fijo — así los campos nuevos del formulario del QR (Segmento,
  // Cuando_Paso, etc., ver registrarSolicitudQR) se pueden seguir agregando
  // sin romper esta función ni depender del orden exacto en que se crearon.
  const headers = rows.length ? rows[0] : [];
  const idx = (nombreCol) => headers.indexOf(nombreCol);
  const val = (row, nombreCol) => { const i = idx(nombreCol); return i>-1 ? row[i] : ''; };
  const data = [];
  for(let i=1;i<rows.length;i++){
    if(rows[i][10] !== 'pendiente_revision') continue;
    data.push({
      id: rows[i][0], fecha: rows[i][1], remitente: rows[i][2], asunto: rows[i][3],
      cuerpo: rows[i][4], clasificacion: rows[i][5], cantPreguntas: rows[i][6],
      requierePago: rows[i][7], estadoPago: rows[i][8], borrador: rows[i][9],
      estado: rows[i][10], fila: i+1,
      // De dónde vino: el nombre de lugar que trae el QR (ver
      // registrarSolicitudQR/asegurarColumnaOrigenColaMensajes_), o
      // vacío/'Mail' para lo que sigue llegando por el triage de Gmail.
      origen: val(rows[i],'Origen') || 'Mail',
      // Nivel de riesgo que estimó la IA a partir del relato del paciente
      // en el formulario del QR (ver evaluarTriageIA_/registrarSolicitudQR).
      // Vacío en lo que llega por mail o en solicitudes de antes de este agregado.
      riesgo: val(rows[i],'Riesgo_IA') || '',
      // 'guardia' o 'entrevista' — qué recomendó la IA (o el valor por
      // defecto si la IA no respondió). Vacío en mensajes que no vinieron por QR.
      recomendacion: val(rows[i],'Recomendacion_IA') || '',
      segmento: val(rows[i],'Segmento') || '',
      cuando: val(rows[i],'Cuando_Paso') || '',
      vioMedico: val(rows[i],'Vio_Medico') || '',
      diagnosticoPrevio: val(rows[i],'Diagnostico_Previo') || '',
      medicacionRecibida: val(rows[i],'Medicacion_Recibida') || '',
      fotoUrl: val(rows[i],'Foto_URL') || '',
      comprobanteUrl: val(rows[i],'Comprobante_Pago_URL') || '',
      comprobanteDatos: val(rows[i],'Comprobante_Pago_Datos') || '',
      tipoSolicitud: val(rows[i],'Tipo_Solicitud') || '',
      // NUEVO: Membresía de Seguimiento — si este mensaje quedó cubierto
      // por la membresía mensual del paciente, y si no, por qué no.
      cubiertoMembresia: val(rows[i],'Cubierto_Membresia') || '',
      motivoFueraMembresia: val(rows[i],'Motivo_Fuera_Membresia') || '',
      avisoAutomatico: val(rows[i],'Aviso_Automatico') || '',
      responderAntesDe: val(rows[i],'Responder_Antes_De') || ''
    });
  }
  // Lo marcado como riesgo "alto" por la IA sube primero, sin importar la
  // fecha — para que el médico lo vea de entrada al abrir la lista.
  const rangoRiesgo_ = function(r){ return r==='alto' ? 0 : (r==='medio' ? 1 : 2); };
  data.sort(function(a,b){
    const diff = rangoRiesgo_(a.riesgo) - rangoRiesgo_(b.riesgo);
    if(diff !== 0) return diff;
    return new Date(b.fecha) - new Date(a.fecha);
  });
  return {success:true, data:data};
}

// ══════════════════════════════════════════════════════════════════
// QR — entrada pública de "consulta virtual de bajo costo", pensada para
// pacientes que llegan por un código QR (guardia, comercios, clubes, etc.)
// en vez de escribir por mail. Cae en la MISMA cola (Cola_Mensajes) que ya
// revisás en Agenda de Mensajes — no es un circuito aparte — sólo que acá
// no hace falta que la IA clasifique nada, porque el paciente ya cargó los
// datos en un formulario. Queda marcada como pago porque es la consulta
// de bajo costo (el cobro real se genera desde Agenda de Mensajes con el
// mismo botón de Mercado Pago que ya usa OCM).
// ══════════════════════════════════════════════════════════════════

// Agrega la columna 'Origen' a Cola_Mensajes si todavía no existe —
// retrocompatible con las filas viejas que llegaron por mail (van a quedar
// con esa celda vacía, y getMensajesPendientes las muestra como 'Mail').
function asegurarColumnaOrigenColaMensajes_(sheet){
  let headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  let col = headers.indexOf('Origen');
  if(col===-1){
    col = headers.length;
    sheet.getRange(1, col+1).setValue('Origen').setBackground('#0c4a6e').setFontColor('white').setFontWeight('bold');
  }
  return col+1; // 1-based, para getRange/appendRow
}

// Igual que la de Origen, pero para el nivel de riesgo que estima la IA
// (ver evaluarTriageIA_) a partir del relato que carga el paciente en el
// formulario del QR.
function asegurarColumnaRiesgoColaMensajes_(sheet){
  let headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  let col = headers.indexOf('Riesgo_IA');
  if(col===-1){
    col = headers.length;
    sheet.getRange(1, col+1).setValue('Riesgo_IA').setBackground('#0c4a6e').setFontColor('white').setFontWeight('bold');
  }
  return col+1;
}

// Agrega cualquier columna nueva de Cola_Mensajes por nombre — versión
// genérica de asegurarColumnaOrigenColaMensajes_/asegurarColumnaRiesgoColaMensajes_,
// para no repetir la misma función una y otra vez por cada campo nuevo del
// formulario del QR.
function asegurarColumnaColaMensajes_(sheet, nombreCol){
  let headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  let col = headers.indexOf(nombreCol);
  if(col===-1){
    col = headers.length;
    sheet.getRange(1, col+1).setValue(nombreCol).setBackground('#0c4a6e').setFontColor('white').setFontWeight('bold');
  }
  return col+1; // 1-based
}

// Guarda (si vino) la foto del estudio/informe que adjunta el paciente en el
// formulario del QR. Nunca bloquea el envío: si falla por lo que sea
// (imagen mal formada, sin permiso de Drive, etc.) devuelve '' y
// registrarSolicitudQR sigue como si no hubiese venido foto.
function guardarFotoEstudioQR_(fotoBase64, nombrePaciente){
  try{
    if(!fotoBase64) return '';
    const partes = String(fotoBase64).split(',');
    const meta = partes.length > 1 ? partes[0] : '';
    const b64 = partes.length > 1 ? partes[1] : partes[0];
    const mimeMatch = meta.match(/data:(.*?);base64/);
    const mime = mimeMatch ? mimeMatch[1] : 'image/jpeg';
    const bytes = Utilities.base64Decode(b64);
    const blob = Utilities.newBlob(bytes, mime, 'estudio_' + Date.now());
    const carpetas = DriveApp.getFoldersByName('SIMM - Fotos de Estudios (QR)');
    const carpeta = carpetas.hasNext() ? carpetas.next() : DriveApp.createFolder('SIMM - Fotos de Estudios (QR)');
    const archivo = carpeta.createFile(blob);
    archivo.setName((nombrePaciente || 'paciente') + ' — ' + archivo.getName());
    return archivo.getUrl();
  }catch(e){
    return '';
  }
}

// Pedido de Luis, sep 2026: mientras Mercado Pago no esté calibrado
// (falta MP_ACCESS_TOKEN — ver generarPreferenciaMP_ / configurarTokenMP),
// el paciente paga por transferencia a un alias y manda el comprobante
// (foto y/o número de operación) desde el mismo formulario del QR. Se
// guarda igual que guardarFotoEstudioQR_ pero en una carpeta propia, para
// no mezclarlo con las fotos de radiografías/informes. Nunca bloquea el
// envío de la solicitud si falla.
function guardarComprobantePagoQR_(comprobanteBase64, nombrePaciente){
  try{
    if(!comprobanteBase64) return '';
    const partes = String(comprobanteBase64).split(',');
    const meta = partes.length > 1 ? partes[0] : '';
    const b64 = partes.length > 1 ? partes[1] : partes[0];
    const mimeMatch = meta.match(/data:(.*?);base64/);
    const mime = mimeMatch ? mimeMatch[1] : 'image/jpeg';
    const bytes = Utilities.base64Decode(b64);
    const blob = Utilities.newBlob(bytes, mime, 'comprobante_' + Date.now());
    const carpetas = DriveApp.getFoldersByName('SIMM - Comprobantes de Pago (QR)');
    const carpeta = carpetas.hasNext() ? carpetas.next() : DriveApp.createFolder('SIMM - Comprobantes de Pago (QR)');
    const archivo = carpeta.createFile(blob);
    archivo.setName((nombrePaciente || 'paciente') + ' — ' + archivo.getName());
    return archivo.getUrl();
  }catch(e){
    return '';
  }
}

// ══════════════════════════════════════════════════════════════════
// GONIO — mediciones angulares y podoscópicas por foto/escáner/cámara.
// Los 4 frontends (GONIO_CostaBartani_v1.html, GONIO_ValgoVaro_v1.html,
// GONIO_AnguloFrente_v1.html, GONIO_Podoscopia_v1.html) ya estaban
// desarrollados y siempre mandaron el mismo contrato de datos
// (action:'guardar_gonio_medicion' con dni/tipo_medicion/pie/angulo/
// clasificacion/comentario/foto_base64/puntos/datos_extra); lo que faltaba
// era esta conexión al backend — por eso el botón "Guardar" nunca guardaba
// nada. Queda todo en una sola hoja (Gonio_Mediciones) distinguido por
// Tipo_Medicion, y la foto (venga de cámara del celular, escáner o webcam —
// al backend le llega igual, como base64) se sube a Drive igual que
// guardarFotoEstudioQR_, sin bloquear el guardado si Drive fallara.
// ══════════════════════════════════════════════════════════════════
function guardarGonioMedicion(d){
  if(!d || !d.dni) return {success:false, error:'Falta el DNI del paciente'};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Gonio_Mediciones');
  if(!sheet){
    sheet = ss.insertSheet('Gonio_Mediciones');
    sheet.appendRow(['Timestamp','DNI','Tipo_Medicion','Pie','Angulo','Clasificacion','Comentario','Foto_URL','Puntos_JSON','Datos_JSON']);
  }
  const fotoUrl = guardarFotoGonio_(d.foto_base64, d.dni, d.tipo_medicion);
  sheet.appendRow([
    new Date().toISOString(), d.dni||'', d.tipo_medicion||'', d.pie||'', d.angulo||'', d.clasificacion||'',
    d.comentario||'', fotoUrl, JSON.stringify(d.puntos||{}), JSON.stringify(d.datos_extra||{})
  ]);
  return {success:true, message:'Medición GONIO guardada en la carpeta del paciente', foto_url: fotoUrl};
}

// Mismo patrón que guardarFotoEstudioQR_: nunca bloquea el guardado de la
// medición si la foto falla por lo que sea (imagen mal formada, sin permiso
// de Drive, etc.) — devuelve '' y guardarGonioMedicion sigue igual.
function guardarFotoGonio_(fotoBase64, dni, tipoMedicion){
  try{
    if(!fotoBase64) return '';
    const partes = String(fotoBase64).split(',');
    const meta = partes.length > 1 ? partes[0] : '';
    const b64 = partes.length > 1 ? partes[1] : partes[0];
    const mimeMatch = meta.match(/data:(.*?);base64/);
    const mime = mimeMatch ? mimeMatch[1] : 'image/jpeg';
    const bytes = Utilities.base64Decode(b64);
    const blob = Utilities.newBlob(bytes, mime, (tipoMedicion||'gonio') + '_' + Date.now());
    const carpetas = DriveApp.getFoldersByName('SIMM - Fotos GONIO');
    const carpeta = carpetas.hasNext() ? carpetas.next() : DriveApp.createFolder('SIMM - Fotos GONIO');
    const archivo = carpeta.createFile(blob);
    archivo.setName('DNI ' + (dni||'?') + ' — ' + (tipoMedicion||'gonio') + ' — ' + archivo.getName());
    return archivo.getUrl();
  }catch(e){
    return '';
  }
}

// Usada por GONIO_Podoscopia_v1.html para traer la última radiografía de
// Costa-Bártani guardada de ese DNI/pie (proporción anatómica del vértice
// del arco) — devuelve TODAS las mediciones GONIO de ese paciente; cada
// frontend filtra por su propio Tipo_Medicion.
function getGonioMediciones(dni){
  if(!dni) return {success:false, error:'Falta el DNI'};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Gonio_Mediciones');
  if(!sheet) return {success:true, data:[]};
  const rows = sheet.getDataRange().getValues();
  if(rows.length<2) return {success:true, data:[]};
  const headers = rows[0];
  const dniCol = headers.indexOf('DNI');
  const data = [];
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][dniCol])===String(dni)){
      const obj = {};
      headers.forEach((h,idx)=>{ obj[h]=rows[i][idx]; });
      data.push(obj);
    }
  }
  return {success:true, data:data};
}

// Mensaje de reserva para cuando la IA no está disponible o falla — nunca
// dejamos al paciente sin ninguna indicación de qué sigue.
function mensajePacienteReserva_(){
  return 'Recibimos tu comprobante y tu mensaje. Mientras verificamos la transacción, tu consulta ya quedó registrada para que el doctor la revise y gestione tu asesoramiento online — te va a responder por escrito, punto por punto, a la brevedad.';
}

// Primera lectura que hace la IA sobre todo lo que cuenta el paciente en el
// formulario del QR (relato, segmento afectado, cuándo pasó, si ya vio un
// médico y qué le dijeron, qué medicación recibió) — pensada para dar, en
// TODOS los casos, una de dos recomendaciones claras: ir ya a una guardia
// traumatológica (si hay señales de alarma reales), o que corresponde
// coordinar la entrevista express de asesoramiento médico paga. NUNCA es un
// diagnóstico, y si la IA falla o no está configurada, no bloquea nada — la
// solicitud igual queda registrada para que el médico la revise a mano.
function evaluarTriageIA_(datos){
  try{
    const motivo = (datos && datos.motivo) || '';
    const prompt =
      'Sos un asistente de triage para el consultorio de traumatología/ortopedia del Dr. Luis Alberto Muratori. ' +
      'Un paciente completó este relato para pedir orientación por QR, en vez de ir directo a una guardia. ' +
      'Tu tarea es dar una PRIMERA LECTURA de riesgo — nunca un diagnóstico — para que el médico decida rápido cómo seguir, ' +
      'y para decirle al paciente, en TODOS los casos, cuál es el paso recomendado: ir a una guardia ahora, ' +
      'o que corresponde la entrevista express de asesoramiento médico (paga, por escrito — el doctor responde por texto, punto por punto, no es una videollamada).\n\n' +
      'Relato del paciente: """' + motivo + '"""\n' +
      'Segmento del cuerpo afectado: ' + ((datos && datos.segmento) || 'no especificado') + '\n' +
      'Cuándo pasó: ' + ((datos && datos.cuando) || 'no especificado') + '\n' +
      '¿Ya vio a un médico?: ' + ((datos && datos.vioMedico) || 'no especificado') +
      ((datos && datos.diagnosticoPrevio) ? (' — diagnóstico que le dieron: ' + datos.diagnosticoPrevio) : '') + '\n' +
      'Medicación que ya recibió: ' + ((datos && datos.medicacionRecibida) || 'ninguna mencionada') + '\n\n' +
      'Devolvé SOLO un JSON válido, sin texto adicional ni markdown, con esta forma exacta:\n' +
      '{ "nivel_riesgo": "bajo, medio o alto", ' +
      '"recomendacion": "guardia o entrevista", ' +
      '"razonamiento": "1-2 frases para el médico, en términos clínicos, explicando por qué asignaste ese nivel", ' +
      '"estudios_habituales": "estudios que suelen pedirse en casos así, en 1 frase, o vacío si no aplica", ' +
      '"mensaje_paciente": "SIEMPRE completo (nunca vacío): 2-3 frases en lenguaje simple, sin nombrar ninguna enfermedad ni usar jerga médica, ' +
      'explicando el paso recomendado — si es guardia, decirle que vaya ahora a la guardia traumatológica más cercana (sin mencionar pagos ni comprobantes, es una derivación de urgencia); ' +
      'si es entrevista, decirle primero que su comprobante quedó recibido y que se está verificando la transacción, y que su mensaje ya quedó registrado para contactar al médico y gestionar su asesoramiento online, ' +
      'y después explicarle que el doctor va a revisar su caso y le va a responder por escrito, punto por punto, con la orientación para su situación." }\n\n' +
      'Reglas: "guardia"/"alto" es solo para señales de alarma reales (déficit neurológico, deformidad evidente, imposibilidad total de apoyo o movimiento, signos de infección grave, dolor torácico, etc.) — no lo uses por precaución genérica. ' +
      'Si no hay señales de alarma claras, usá "entrevista" con nivel "bajo" o "medio". No inventes datos que el paciente no haya contado.';

    const resp = UrlFetchApp.fetch(getGeminiUrl_(), {
      method:'post', contentType:'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { thinkingConfig: { thinkingLevel: 'minimal' } }
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(resp.getContentText());
    registrarUsoTokens_('evaluarTriageIA_', 'gemini-3.6-flash', data.usageMetadata);
    if(!data.candidates || !data.candidates.length) return null;
    let texto = data.candidates[0].content.parts[0].text || '';
    texto = texto.replace(/```json|```/g,'').trim();
    const j = JSON.parse(texto);
    if(!j || ['bajo','medio','alto'].indexOf(j.nivel_riesgo)===-1) return null;
    if(['guardia','entrevista'].indexOf(j.recomendacion)===-1){
      j.recomendacion = j.nivel_riesgo==='alto' ? 'guardia' : 'entrevista';
    }
    if(!j.mensaje_paciente) j.mensaje_paciente = mensajePacienteReserva_();
    return j;
  }catch(e){
    return null;
  }
}

// d = {tipoSolicitud, nombre, dni, telefono, email, obraSocial, motivo,
//      segmento, cuando, vioMedico, diagnosticoPrevio, medicacionRecibida,
//      fotoBase64, comprobanteBase64, comprobanteDatos, origen}
function registrarSolicitudQR(d){
  if(!d || !String(d.nombre||'').trim()) return {success:false, error:'Falta el nombre'};
  if(!String(d.motivo||'').trim()) return {success:false, error:'Falta contar el motivo de consulta'};
  if(!String(d.telefono||'').trim() && !String(d.email||'').trim()){
    return {success:false, error:'Dejá al menos un teléfono o un mail de contacto'};
  }

  const sheet = hojaColaMensajes_();
  const colOrigen = asegurarColumnaOrigenColaMensajes_(sheet);
  const colRiesgo = asegurarColumnaRiesgoColaMensajes_(sheet);
  const colSegmento = asegurarColumnaColaMensajes_(sheet, 'Segmento');
  const colCuando = asegurarColumnaColaMensajes_(sheet, 'Cuando_Paso');
  const colVioMedico = asegurarColumnaColaMensajes_(sheet, 'Vio_Medico');
  const colDiagPrevio = asegurarColumnaColaMensajes_(sheet, 'Diagnostico_Previo');
  const colMedicacion = asegurarColumnaColaMensajes_(sheet, 'Medicacion_Recibida');
  const colFoto = asegurarColumnaColaMensajes_(sheet, 'Foto_URL');
  const colRecomendacion = asegurarColumnaColaMensajes_(sheet, 'Recomendacion_IA');
  const colComprobanteUrl = asegurarColumnaColaMensajes_(sheet, 'Comprobante_Pago_URL');
  const colComprobanteDatos = asegurarColumnaColaMensajes_(sheet, 'Comprobante_Pago_Datos');
  const colTipoSolicitud = asegurarColumnaColaMensajes_(sheet, 'Tipo_Solicitud');
  // AGREGADA (sept 2026, pedido de Luis): la entrevista express NO es una
  // teleconsulta por videollamada — es una orientación por escrito (relato
  // dictado o tipeado por el paciente, respondida por el doctor punto por
  // punto), así que no aplican ni la Resolución 3316/2023 (que rige
  // videollamada/audio) ni la grabación — solo queda un código simple que
  // identifica la orientación escrita en la historia clínica.
  const colCodigoOrient = asegurarColumnaColaMensajes_(sheet, 'Codigo_Orientacion_Escrita');
  const anchoFila = Math.max(colOrigen, colRiesgo, colSegmento, colCuando, colVioMedico, colDiagPrevio, colMedicacion, colFoto, colRecomendacion, colComprobanteUrl, colComprobanteDatos, colTipoSolicitud, colCodigoOrient);

  // La foto nunca bloquea el envío — si falla, sigue todo igual sin ella.
  const fotoUrl = d.fotoBase64 ? guardarFotoEstudioQR_(d.fotoBase64, d.nombre) : '';
  // Comprobante de pago (mientras Mercado Pago no esté calibrado, ver
  // guardarComprobantePagoQR_) — tampoco bloquea el envío si falla.
  const comprobanteUrl = d.comprobanteBase64 ? guardarComprobantePagoQR_(d.comprobanteBase64, d.nombre) : '';
  const comprobanteDatos = String(d.comprobanteDatos||'').trim();

  const triage = evaluarTriageIA_(d);
  const recomendacion = triage ? triage.recomendacion : 'entrevista'; // sin IA, por defecto se coordina la entrevista (nunca se asume urgencia sola)
  const mensajePaciente = triage ? triage.mensaje_paciente : mensajePacienteReserva_();

  const asunto = 'Solicitud por QR — ' + d.nombre + (d.tipoSolicitud ? (' — ' + d.tipoSolicitud) : '');
  let cuerpo =
    'Solicitud de orientación / entrevista express, ingresada por QR.\n\n' +
    'Qué necesita: ' + (d.tipoSolicitud||'—') + '\n' +
    'Nombre: ' + (d.nombre||'') + '\n' +
    'DNI: ' + (d.dni||'') + '\n' +
    'Teléfono: ' + (d.telefono||'') + '\n' +
    'Email: ' + (d.email||'') + '\n' +
    'Obra social / cobertura: ' + (d.obraSocial||'Particular') + '\n' +
    'Origen del QR: ' + (d.origen||'sin especificar') + '\n\n' +
    'Motivo de consulta:\n' + (d.motivo||'') + '\n\n' +
    'Segmento afectado: ' + (d.segmento||'—') + '\n' +
    'Cuándo pasó: ' + (d.cuando||'—') + '\n' +
    '¿Ya vio a un médico?: ' + (d.vioMedico||'—') + (d.diagnosticoPrevio ? (' — diagnóstico dado: ' + d.diagnosticoPrevio) : '') + '\n' +
    'Medicación/receta que le dieron: ' + (d.medicacionRecibida||'—') +
    (fotoUrl ? ('\nFoto de estudio/informe adjunta: ' + fotoUrl) : '') +
    (comprobanteUrl ? ('\nComprobante de pago (foto): ' + comprobanteUrl) : '') +
    (comprobanteDatos ? ('\nComprobante de pago (datos informados): ' + comprobanteDatos) : '');
  if(triage){
    cuerpo += '\n\n── Evaluación preliminar (IA — no reemplaza tu criterio clínico) ──\n' +
      'Nivel de riesgo estimado: ' + triage.nivel_riesgo + '\n' +
      'Recomendación: ' + (triage.recomendacion==='guardia' ? 'ir a guardia ahora' : 'coordinar entrevista express') + '\n' +
      'Razonamiento: ' + (triage.razonamiento||'—') +
      (triage.estudios_habituales ? ('\nEstudios habituales en casos así: ' + triage.estudios_habituales) : '');
  }
  const borrador =
    'Hola ' + (d.nombre||'') + ',\n\n' +
    (recomendacion==='guardia'
      ? 'Por lo que nos contaste, te recomendamos ir ahora a la guardia traumatológica más cercana para una evaluación en persona sin esperar.'
      : 'Recibimos tu consulta. La vamos a revisar y te vamos a responder por escrito, punto por punto, a la brevedad.') + '\n\n' +
    'Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943';

  // Código de orientación escrita: solo cuando el triage recomienda
  // entrevista (no cuando manda a guardia, que es derivación presencial
  // directa) y no es una simple renovación de receta (eso es administrativo,
  // no una orientación clínica).
  const esRenovacionReceta_ = /renovaci[oó]n de receta/i.test(String(d.tipoSolicitud||''));
  const esOrientacionEscrita_ = (recomendacion === 'entrevista') && !esRenovacionReceta_;

  const id = Utilities.getUuid();
  const codigoOrientacion = esOrientacionEscrita_ ? ('OE-' + id.split('-')[0].toUpperCase()) : '';
  const fila = new Array(anchoFila).fill('');
  fila[0] = id;
  fila[1] = new Date().toISOString();
  fila[2] = d.email || d.telefono || '';
  fila[3] = asunto;
  fila[4] = cuerpo;
  fila[5] = 'privado';
  fila[6] = 1;
  fila[7] = true;          // requierePago — la entrevista express es paga (salvo derivación directa a guardia)
  fila[8] = 'pendiente';   // estadoPago
  fila[9] = borrador;
  fila[10] = 'pendiente_revision';
  fila[11] = '';           // fila_referencia
  fila[12] = '';           // threadId — no viene de un hilo de Gmail
  fila[colOrigen-1] = d.origen || 'Sin especificar';
  fila[colRiesgo-1] = triage ? triage.nivel_riesgo : '';
  fila[colSegmento-1] = d.segmento || '';
  fila[colCuando-1] = d.cuando || '';
  fila[colVioMedico-1] = d.vioMedico || '';
  fila[colDiagPrevio-1] = d.diagnosticoPrevio || '';
  fila[colMedicacion-1] = d.medicacionRecibida || '';
  fila[colFoto-1] = fotoUrl;
  fila[colRecomendacion-1] = recomendacion;
  fila[colComprobanteUrl-1] = comprobanteUrl;
  fila[colComprobanteDatos-1] = comprobanteDatos;
  fila[colTipoSolicitud-1] = d.tipoSolicitud || '';
  fila[colCodigoOrient-1] = codigoOrientacion;
  sheet.appendRow(fila);

  return {
    success: true,
    id: id,
    codigo_orientacion: codigoOrientacion,
    recomendacion: recomendacion,
    alerta_riesgo: recomendacion==='guardia',
    mensaje_paciente: mensajePaciente
  };
}

function confirmarPagoMensaje(d){
  const sheet = hojaColaMensajes_();
  sheet.getRange(d.fila, 9).setValue('confirmado');
  return {success:true};
}

// Si en la Gmail de lmuratori agregaste consultas.sism@gmail.com como dirección
// "Enviar como" (Configuración → Cuentas e importación), las respuestas salen
// "de" consultas.sism y las respuestas del paciente vuelven ahí. Si todavía no
// está agregada, sigue saliendo como antes (sin romper nada).
function aliasConsultasSism_(){
  try{ return GmailApp.getAliases().indexOf(MEDICO_MAIL_KINESIOLOGIA) > -1 ? MEDICO_MAIL_KINESIOLOGIA : ''; }
  catch(e){ return ''; }
}
function marcarRespondido(d){
  const sheet = hojaColaMensajes_();
  const threadId = sheet.getRange(d.fila, 13).getValue();
  const alias = aliasConsultasSism_();
  const firma = { name: 'Dr. Luis Alberto Muratori' };
  if(alias && d.destinatario){
    GmailApp.sendEmail(d.destinatario, 'Re: ' + (d.asunto||'Su consulta'), d.respuestaFinal, Object.assign({ from: alias, replyTo: alias }, firma));
  }else{
    try{
      const thread = GmailApp.getThreadById(threadId);
      thread.reply(d.respuestaFinal, firma);
    }catch(e){
      GmailApp.sendEmail(d.destinatario, 'Re: ' + (d.asunto||'Su consulta'), d.respuestaFinal, firma);
    }
  }
  sheet.getRange(d.fila, 11).setValue('respondido');
  return {success:true};
}

// Genera una reunión de Google Meet REAL (no solo un link a un evento de
// Calendar sin videollamada adjunta — eso era lo que hacía la versión
// anterior de esta función con CalendarApp.createEvent, que nunca pide
// una conferencia y por eso el "link" resultante muchas veces no tenía
// ninguna sala de Meet detrás).
//
// Para que esto funcione hace falta UNA sola vez: en el editor de Apps
// Script, ir a Servicios (el + al lado de "Servicios" en el panel
// izquierdo) → agregar "Google Calendar API" (identificador "Calendar").
// Si ese servicio no está agregado, esta función va a devolver un error
// explicando exactamente eso, en vez de fallar en silencio.
function generarLinkMeet(d){
  const inicio = new Date(d.fechaHoraISO);
  const fin = new Date(inicio.getTime() + (d.duracionMin||5)*60000);
  const eventoAEnviar = {
    summary: d.titulo || 'Consulta virtual con el Dr. Muratori',
    description: d.descripcion || '',
    start: { dateTime: inicio.toISOString() },
    end: { dateTime: fin.toISOString() },
    attendees: d.mailPaciente ? [{ email: d.mailPaciente }] : [],
    conferenceData: {
      createRequest: {
        requestId: Utilities.getUuid(),
        conferenceSolutionKey: { type: 'hangoutsMeet' }
      }
    }
  };
  try{
    const creado = Calendar.Events.insert(eventoAEnviar, 'primary', {
      conferenceDataVersion: 1,
      sendUpdates: d.mailPaciente ? 'all' : 'none'
    });
    let linkMeet = '';
    if(creado.conferenceData && creado.conferenceData.entryPoints){
      const entradaVideo = creado.conferenceData.entryPoints.filter(function(ep){ return ep.entryPointType === 'video'; })[0];
      if(entradaVideo) linkMeet = entradaVideo.uri;
    }
    if(!linkMeet) linkMeet = creado.hangoutLink || '';
    return {
      success: true,
      link: linkMeet || (creado.htmlLink || 'Evento creado, pero no se pudo confirmar el link de Meet — revisá el evento en Google Calendar.'),
      link_es_meet_real: !!linkMeet,
      evento_id: creado.id,
      link_calendario: creado.htmlLink || ''
    };
  }catch(e){
    return {
      error: 'No se pudo crear la reunión de Meet: ' + e.message + '. Si el error menciona "Calendar is not defined" o similar, falta activar el servicio avanzado "Google Calendar API" en Servicios del editor de Apps Script (una sola vez).'
    };
  }
}

// ══════════════════════════════════════════════════════════════════
// AGENDA DE TURNOS — fuente única de verdad para los turnos presenciales,
// compartida por dos puntos de entrada distintos que antes no se cruzaban:
//   1) El médico, al cerrar el Espacio de Consulta Virtual (después de la
//      videollamada), eligiendo un horario real para el control presencial.
//   2) Las secretarias de cada consultorio, desde Turnos_Secretaria.html
//      (una página por sede, reachable por su propio QR impreso).
// Como las dos escriben en la MISMA hoja (Turnos_Agenda) y reservarTurno_
// vuelve a chequear el horario elegido antes de guardarlo, un turno que ya
// tomó la secretaria de Palmares no puede quedar pisado por uno que el
// médico ofrezca por la misma vía de la videollamada, ni al revés.
//
// Pedido de Luis A. Muratori, sep 2026: horarios reales cargados a mano
// (no hay una fuente automática de "disponibilidad" — son los horarios de
// atención habituales de cada consultorio). Si cambian, alcanza con editar
// CONSULTORIOS_HORARIOS_ acá abajo.
// ══════════════════════════════════════════════════════════════════
// DURACION_TURNO_MIN_ queda como el valor por defecto (lo sigue usando el
// Hospital Sícoli, que es agenda libre — ver más abajo). Pedido de Luis,
// sep 2026: "Centro Médico Palmares, Rehabilitarte y Consultorios Privados
// del Norte son cada 15 minutos, Hospital Sícoli cada 20 minutos" — como
// cada consultorio puede tener su propia duración, se agregó
// CONSULTORIOS_DURACION_MIN_ (si un consultorio no figura ahí, se usa
// DURACION_TURNO_MIN_ como antes).
const DURACION_TURNO_MIN_ = 20;
const CONSULTORIOS_DURACION_MIN_ = {
  'Centro Médico Palmares': 15,
  'Rehabilitarte Instituto Médico': 15,
  'Consultorios Privados del Norte': 15
  // 'Hospital Dr. Domingo Sícoli' no hace falta listarlo: usa el default (20)
};
function duracionTurnoConsultorio_(nombreConsultorio){
  return CONSULTORIOS_DURACION_MIN_[nombreConsultorio] || DURACION_TURNO_MIN_;
}
const SEMANAS_AGENDA_ = 4;
// dia: 0=domingo … 6=sábado (igual que Date.getDay())
const CONSULTORIOS_HORARIOS_ = {
  // oct 2026, pedido de Luis: Hospital Sícoli con horario fijo, turnos cada 20 min.
  // Martes 6:00–12:00 · miércoles y jueves 6:00–12:20 · viernes 6:00–13:00 (almuerzo 13–14) y 14:00–16:40.
  'Hospital Dr. Domingo Sícoli': [
    { dia: 2, desde: '06:00', hasta: '12:00' },
    { dia: 3, desde: '06:00', hasta: '12:20' },
    { dia: 4, desde: '06:00', hasta: '12:20' },
    { dia: 5, desde: '06:00', hasta: '13:00' },
    { dia: 5, desde: '14:00', hasta: '16:40' }
  ],
  'Centro Médico Palmares': [
    { dia: 2, desde: '13:30', hasta: '16:45' }, // martes, cada 15 min
    { dia: 4, desde: '13:30', hasta: '20:00' }  // jueves, cada 15 min
  ],
  'Rehabilitarte Instituto Médico': [
    { dia: 5, desde: '17:00', hasta: '18:00' }  // viernes, cada 15 min
  ]
  // 'Consultorios Privados del Norte' NO va acá: se define mes a mes (agenda libre, ver abajo).
};

// Pedido de Luis, sep 2026: el Hospital Dr. Domingo Sícoli es un hospital
// público — no tiene un horario semanal fijo como los consultorios
// privados de arriba, así que no le sirve un patrón fijo en
// CONSULTORIOS_HORARIOS_. En cambio, tiene AGENDA LIBRE: la secretaria del
// hospital va cargando ella misma, día a día, los horarios puntuales que
// van quedando disponibles (ver agregarHorarioLibre/hojaHorariosLibres_
// más abajo), y esos horarios entran a la misma bolsa de disponibilidad
// que ve el paciente/la app, con el mismo control de choque contra
// Turnos_Agenda que ya usa reservarTurno.
// Consultorios del Norte: los viernes se definen mes a mes (18:00–20:30 o 19:00–21:00, cada 15 min) — se cargan con agregarHorarioLibre.
const CONSULTORIOS_AGENDA_LIBRE_ = ['Consultorios Privados del Norte'];

// ══════════════════════════════════════════════════════════════
// SEÑA DE RESERVA + RECORDATORIOS DE TURNO (oct 2026, pedido de Luis) — los
// consultorios privados de abajo piden una seña por transferencia para
// asegurar el turno; el Hospital Dr. Domingo Sícoli (agenda libre) NO
// lleva seña, solo recordatorio del día anterior para confirmar asistencia.
// Para agregar o sacar un consultorio de la seña, alcanza con editar este
// array — no hace falta tocar nada más.
const CONSULTORIOS_CON_SENA_ = ['Centro Médico Palmares', 'Rehabilitarte Instituto Médico', 'Consultorios Privados del Norte'];

// Mismo alias que ya usa QR_Consulta_Virtual.html para la entrevista
// express, para no repartir dos alias distintos. Si cambia, hay que
// actualizarlo en los dos lugares — uno es este backend y el otro es un
// archivo estático en Cloudflare Pages sin backend propio, no comparten
// configuración.
const ALIAS_TRANSFERENCIA_SENA_ = 'BUCLE.PODIAS.BALCON';

// Monto de la seña — configurable sin tocar código: Configuración del
// proyecto → Propiedades del script → MONTO_SENA_TURNO (ej: 10000). Si no
// se configuró nada, usa este valor por defecto.
function montoSenaTurno_(){
  const prop = PropertiesService.getScriptProperties().getProperty('MONTO_SENA_TURNO');
  const monto = parseFloat(prop);
  return (monto && monto>0) ? monto : 10000;
}

// Base pública de SIM-M para armar links dentro de los mails que manda
// este backend (Apps Script no sabe su propia URL de sitio — son archivos
// estáticos aparte, en Cloudflare Pages). Cuando se haga la migración a
// sistema.drmuratorisalud.ar (ver GUIA_DE_INSTALACION.md, Paso 1),
// actualizar esta única línea.
// DOMINIOS (oct 2026): drmuratorisalud.ar pasa a ser el sitio público de
// salud (artículos + Clave A/B); S.I.M-M móvil vive en el subdominio
// sistema.drmuratorisalud.ar. TODOS los links que este backend le manda a
// pacientes (RIMP, RMYT, QR de consulta virtual, gestión de turno, retorno
// de Mercado Pago, seguimiento posqx) salen de esta función. Para cambiarla
// sin tocar código: Propiedades del script → SIMM_BASE_URL (con la barra final).
function baseSimm_(){
  var p = '';
  try{ p = (PropertiesService.getScriptProperties().getProperty('SIMM_BASE_URL')||'').trim(); }catch(e){}
  if(!p) p = 'https://sistema.drmuratorisalud.ar/';
  return p.replace(/\/*$/, '/');
}

function hojaTurnosAgenda_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Turnos_Agenda');
  if(!sheet){
    sheet = ss.insertSheet('Turnos_Agenda');
    sheet.appendRow(['ID_Turno','Timestamp_Reserva','Consultorio','Fecha','Hora','Duracion_Min',
      'DNI','Nombre','Telefono','Email','Origen','Estado','Notas']);
    sheet.getRange(1,1,1,13).setBackground('#0c4a6e').setFontColor('white').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Hoja donde la secretaria del Hospital Sícoli (o cualquier otro
// consultorio de agenda libre que se sume en el futuro) va cargando los
// horarios puntuales que tiene disponibles — ver CONSULTORIOS_AGENDA_LIBRE_.
function hojaHorariosLibres_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Horarios_Libres');
  if(!sheet){
    sheet = ss.insertSheet('Horarios_Libres');
    sheet.appendRow(['ID','Timestamp_Carga','Consultorio','Fecha','Hora','Cargado_Por','Estado']);
    sheet.getRange(1,1,1,7).setBackground('#0c4a6e').setFontColor('white').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// d = {consultorio, fecha:'yyyy-MM-dd', horaDesde:'HH:mm', horaHasta:'HH:mm',
//      cargadoPor}
// La secretaria del hospital carga un RANGO horario (ej. "de 6 a 8") y acá
// se parte solo en turnos de DURACION_TURNO_MIN_ (20 minutos — pedido de
// Luis, sep 2026: "los horarios del Sícoli son cada 20 minutos", mismo
// intervalo que ya usan los demás consultorios) — así la secretaria no
// tiene que cargar cada turno individual a mano. No reemplaza nada de
// CONSULTORIOS_HORARIOS_, es exclusivo de los consultorios listados en
// CONSULTORIOS_AGENDA_LIBRE_.
function agregarHorarioLibre(d){
  if(!d || !String(d.consultorio||'').trim()) return {success:false, error:'Falta el consultorio'};
  if(CONSULTORIOS_AGENDA_LIBRE_.indexOf(d.consultorio)===-1){
    return {success:false, error:'Este consultorio no usa agenda libre'};
  }
  if(!String(d.fecha||'').trim() || !String(d.horaDesde||'').trim() || !String(d.horaHasta||'').trim()){
    return {success:false, error:'Falta la fecha o el rango horario (desde/hasta)'};
  }
  const inicioMin = minutosDesdeHHMM_(d.horaDesde);
  const finMin = minutosDesdeHHMM_(d.horaHasta);
  if(finMin <= inicioMin) return {success:false, error:'El horario "hasta" tiene que ser posterior al "desde"'};

  const sheet = hojaHorariosLibres_();
  const filas = sheet.getDataRange().getValues();
  const yaCargados = {}; // clave "Fecha|Hora" -> true, para no duplicar
  for(let i=1;i<filas.length;i++){
    const estado = String(filas[i][6]||'').toLowerCase();
    if(estado==='eliminado') continue;
    if(filas[i][2]===d.consultorio) yaCargados[filas[i][3]+'|'+filas[i][4]] = true;
  }

  const filasNuevas = [];
  const durLibre = duracionTurnoConsultorio_(d.consultorio);
  for(let m=inicioMin; m+durLibre<=finMin; m+=durLibre){
    const hh = String(Math.floor(m/60)).padStart(2,'0');
    const mm = String(m%60).padStart(2,'0');
    const horaStr = hh+':'+mm;
    if(yaCargados[d.fecha+'|'+horaStr]) continue; // ya estaba cargado, no duplicar
    filasNuevas.push([Utilities.getUuid(), new Date().toISOString(), d.consultorio, d.fecha, horaStr, d.cargadoPor||'', 'disponible']);
  }
  if(!filasNuevas.length) return {success:false, error:'No se cargó ningún horario nuevo — puede que ya estuvieran todos cargados, o que el rango sea menor a la duración de un turno'};

  sheet.getRange(sheet.getLastRow()+1, 1, filasNuevas.length, filasNuevas[0].length).setValues(filasNuevas);
  return {success:true, cantidad:filasNuevas.length, horarios:filasNuevas.map(f=>f[4])};
}

function minutosDesdeHHMM_(hhmm){
  const partes = String(hhmm||'0:0').split(':');
  return (parseInt(partes[0],10)||0)*60 + (parseInt(partes[1],10)||0);
}
function fechaYYYYMMDD_(d){
  return Utilities.formatDate(d, 'America/Argentina/Mendoza', 'yyyy-MM-dd');
}

// Genera, para un consultorio (o para todos si no se especifica), la lista
// de horarios de las próximas SEMANAS_AGENDA_ semanas, ya descontando los
// que figuran ocupados en Turnos_Agenda (cualquier estado salvo
// 'cancelado') y los que ya pasaron. Se agrupa por fecha para que el
// frontend los pueda mostrar como una lista de días con sus horarios.
function getTurnosDisponibles(params){
  params = params || {};
  const consultorioFiltro = params.consultorio || '';
  const todosLosConsultorios = Object.keys(CONSULTORIOS_HORARIOS_).concat(CONSULTORIOS_AGENDA_LIBRE_);
  const nombresConsultorios = consultorioFiltro ? [consultorioFiltro] : todosLosConsultorios;

  const sheet = hojaTurnosAgenda_();
  const filas = sheet.getDataRange().getValues();
  const ocupados = {}; // clave "Consultorio|Fecha|Hora" -> true
  for(let i=1;i<filas.length;i++){
    const estado = String(filas[i][11]||'').toLowerCase();
    if(estado === 'cancelado') continue;
    const clave = filas[i][2] + '|' + filas[i][3] + '|' + filas[i][4];
    ocupados[clave] = true;
  }

  const ahora = new Date();
  const resultado = {}; // Consultorio -> [{fecha, dia_semana, horarios:[hora,...]}]

  nombresConsultorios.forEach(function(nombreConsultorio){
    if(CONSULTORIOS_AGENDA_LIBRE_.indexOf(nombreConsultorio)>-1){
      resultado[nombreConsultorio] = turnosDisponiblesAgendaLibre_(nombreConsultorio, ocupados, ahora);
      return;
    }
    const ventanas = CONSULTORIOS_HORARIOS_[nombreConsultorio];
    if(!ventanas) return;
    const duracion = duracionTurnoConsultorio_(nombreConsultorio);
    const porFecha = {};
    for(let semana=0; semana<SEMANAS_AGENDA_; semana++){
      for(let dias=0; dias<7; dias++){
        const fecha = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate() + semana*7 + dias);
        const diaSemana = fecha.getDay();
        ventanas.filter(function(v){ return v.dia === diaSemana; }).forEach(function(v){
          const inicioMin = minutosDesdeHHMM_(v.desde);
          const finMin = minutosDesdeHHMM_(v.hasta);
          const fechaStr = fechaYYYYMMDD_(fecha);
          for(let m=inicioMin; m+duracion<=finMin; m+=duracion){
            const hh = String(Math.floor(m/60)).padStart(2,'0');
            const mm = String(m%60).padStart(2,'0');
            const horaStr = hh+':'+mm;
            const slotDateTime = new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate(), Math.floor(m/60), m%60);
            if(slotDateTime <= ahora) continue; // no ofrecer horarios ya pasados
            const clave = nombreConsultorio+'|'+fechaStr+'|'+horaStr;
            if(ocupados[clave]) continue;
            if(!porFecha[fechaStr]) porFecha[fechaStr] = [];
            porFecha[fechaStr].push(horaStr);
          }
        });
      }
    }
    resultado[nombreConsultorio] = Object.keys(porFecha).sort().map(function(f){
      return { fecha: f, horarios: porFecha[f] };
    });
  });

  // duracion_min ahora es un mapa por consultorio (antes era un solo número
  // para todos) — cada consultorio puede durar distinto (ver
  // CONSULTORIOS_DURACION_MIN_ más arriba).
  const duracionesPorConsultorio = {};
  nombresConsultorios.forEach(function(n){ duracionesPorConsultorio[n] = duracionTurnoConsultorio_(n); });
  return { success: true, consultorios: resultado, duracion_min: duracionesPorConsultorio };
}

// Para un consultorio de agenda libre (ver CONSULTORIOS_AGENDA_LIBRE_):
// en vez de calcular horarios a partir de un patrón semanal fijo, los lee
// tal cual los fue cargando la secretaria en Horarios_Libres, descartando
// los que ya pasaron o que ya están ocupados en Turnos_Agenda (mismo
// control de choque que usa el resto de la agenda).
function turnosDisponiblesAgendaLibre_(nombreConsultorio, ocupados, ahora){
  const sheet = hojaHorariosLibres_();
  const filas = sheet.getDataRange().getValues();
  const porFecha = {};
  for(let i=1;i<filas.length;i++){
    const consultorio = filas[i][2], fechaStr = filas[i][3], horaStr = filas[i][4];
    const estado = String(filas[i][6]||'').toLowerCase();
    if(consultorio !== nombreConsultorio) continue;
    if(estado !== 'disponible') continue; // 'eliminado' o cualquier otro estado, no se ofrece
    const clave = nombreConsultorio+'|'+fechaStr+'|'+horaStr;
    if(ocupados[clave]) continue; // ya reservado en Turnos_Agenda
    const partes = String(horaStr||'0:0').split(':');
    const fechaPartes = String(fechaStr||'').split('-').map(Number);
    if(fechaPartes.length!==3) continue;
    const slotDateTime = new Date(fechaPartes[0], fechaPartes[1]-1, fechaPartes[2], parseInt(partes[0],10)||0, parseInt(partes[1],10)||0);
    if(slotDateTime <= ahora) continue; // ya pasó
    if(!porFecha[fechaStr]) porFecha[fechaStr] = [];
    if(porFecha[fechaStr].indexOf(horaStr)===-1) porFecha[fechaStr].push(horaStr);
  }
  return Object.keys(porFecha).sort().map(function(f){
    return { fecha: f, horarios: porFecha[f].sort() };
  });
}

// ══════════════════════════════════════════════════════════════════
// CALENDARIO DEL MES (oct 2026, pedido de Luis) — vista mensual día por día
// (lunes a viernes) de TODOS los turnos presenciales: cada horario de cada
// consultorio con su estado (libre / ocupado). Los NOMBRES de los pacientes
// solo se incluyen si viene el token del médico; sin token se ve únicamente
// libre/ocupado.
// ══════════════════════════════════════════════════════════════════
function celdaFecha_(v){
  if(v instanceof Date) return Utilities.formatDate(v, 'America/Argentina/Mendoza', 'yyyy-MM-dd');
  return String(v||'').slice(0,10);
}
function celdaHora_(v){
  if(v instanceof Date) return Utilities.formatDate(v, 'America/Argentina/Mendoza', 'HH:mm');
  const s = String(v||'');
  const m = s.match(/^(\d{1,2}):(\d{2})/);
  return m ? (m[1].padStart(2,'0')+':'+m[2]) : s;
}
function getAgendaMes(params){
  params = params || {};
  const hoy = new Date();
  let mes = String(params.mes||'').trim();
  if(!/^\d{4}-\d{2}$/.test(mes)) mes = Utilities.formatDate(hoy, 'America/Argentina/Mendoza', 'yyyy-MM');
  const anio = parseInt(mes.slice(0,4),10), nMes = parseInt(mes.slice(5,7),10);
  const verNombres = tokenValido_(params.token); // sin token configurado en el backend, tokenValido_ deja pasar
  const diasEnMes = new Date(anio, nMes, 0).getDate();

  // Turnos reservados del mes (todo lo que no esté cancelado)
  const ocupados = {}; // "Consultorio|fecha|hora" -> {id, paciente, estado}
  const filas = hojaTurnosAgenda_().getDataRange().getValues();
  for(let i=1;i<filas.length;i++){
    if(String(filas[i][11]||'').toLowerCase()==='cancelado') continue;
    const f = celdaFecha_(filas[i][3]);
    if(f.slice(0,7)!==mes) continue;
    ocupados[filas[i][2]+'|'+f+'|'+celdaHora_(filas[i][4])] = { id: filas[i][0], paciente: verNombres ? String(filas[i][7]||'') : '', estado: String(filas[i][11]||'') };
  }
  const dias = {};
  const agregar = function(f, consultorio, hora, dur){
    const clave = consultorio+'|'+f+'|'+hora;
    const oc = ocupados[clave];
    if(!dias[f]) dias[f] = [];
    dias[f].push({ consultorio: consultorio, hora: hora, duracion: dur, estado: oc ? 'ocupado' : 'libre', paciente: oc ? oc.paciente : '', id: oc ? oc.id : '' });
    if(oc) oc.usado = true;
  };
  // Consultorios con horario fijo
  Object.keys(CONSULTORIOS_HORARIOS_).forEach(function(nombre){
    const dur = duracionTurnoConsultorio_(nombre);
    for(let d=1; d<=diasEnMes; d++){
      const fecha = new Date(anio, nMes-1, d);
      const f = Utilities.formatDate(fecha, 'America/Argentina/Mendoza', 'yyyy-MM-dd');
      CONSULTORIOS_HORARIOS_[nombre].filter(function(v){ return v.dia===fecha.getDay(); }).forEach(function(v){
        const ini = minutosDesdeHHMM_(v.desde), fin = minutosDesdeHHMM_(v.hasta);
        for(let m=ini; m+dur<=fin; m+=dur){
          agregar(f, nombre, String(Math.floor(m/60)).padStart(2,'0')+':'+String(m%60).padStart(2,'0'), dur);
        }
      });
    }
  });
  // Consultorios de agenda libre (se cargan mes a mes)
  const libres = hojaHorariosLibres_().getDataRange().getValues();
  for(let i=1;i<libres.length;i++){
    if(String(libres[i][6]||'').toLowerCase()!=='disponible') continue;
    const f = celdaFecha_(libres[i][3]);
    if(f.slice(0,7)!==mes) continue;
    agregar(f, libres[i][2], celdaHora_(libres[i][4]), duracionTurnoConsultorio_(libres[i][2]));
  }
  // Turnos reservados que ya no corresponden a ningún horario del patrón (por si cambió): igual se muestran
  Object.keys(ocupados).forEach(function(k){
    const o = ocupados[k]; if(o.usado) return;
    const p = k.split('|');
    if(!dias[p[1]]) dias[p[1]] = [];
    dias[p[1]].push({ consultorio: p[0], hora: p[2], duracion: duracionTurnoConsultorio_(p[0]), estado: 'ocupado', paciente: o.paciente, id: o.id });
  });
  Object.keys(dias).forEach(function(f){ dias[f].sort(function(a,b){ return a.hora.localeCompare(b.hora) || a.consultorio.localeCompare(b.consultorio); }); });
  return { success:true, mes: mes, dias: dias, ve_nombres: verNombres };
}

// ══════════════════════════════════════════════════════════════════
// SEGUIMIENTO DE PACIENTES (oct 2026, pedido de Luis) — resumen liviano de cada
// paciente del circuito (tratamiento médico → prequirúrgico → programación →
// postop) para que el celular/web lo vea igual que la PC. OCM lo sube al guardar
// el Libro Quirúrgico y a las 6:00, 12:30, 17:00 y 21:00. NO lleva fotos ni la
// historia clínica: solo lo necesario para la lista. Todo con token.
// ══════════════════════════════════════════════════════════════════
const SEG_COLS_ = ['DNI','Paciente','Obra_Social','Telefono','Mail','Etapa','Diagnostico','CIE10','Procedimiento','Institucion','Checklist_JSON','Ultima_Consulta','Ultima_Consulta_ISO','Proxima_Cita','Postop_Control','Caso','Actualizado'];
function hojaSeguimiento_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Seguimiento_Pacientes');
  if(!sh){
    sh = ss.insertSheet('Seguimiento_Pacientes');
    sh.appendRow(SEG_COLS_);
    sh.getRange(1,1,1,SEG_COLS_.length).setBackground('#0f766e').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function sincronizarSeguimientoLote(d){
  if(!d || !Array.isArray(d.pacientes)) return {success:false, error:'Falta la lista de pacientes'};
  const lock = LockService.getScriptLock();
  try{
    lock.waitLock(20000);
    const sh = hojaSeguimiento_();
    const data = sh.getDataRange().getValues();
    const filas = [], idx = {};
    for(let i=1;i<data.length;i++){
      const r = data[i].slice(0,SEG_COLS_.length); while(r.length<SEG_COLS_.length) r.push('');
      filas.push(r); if(r[0]) idx[String(r[0])] = filas.length-1;
    }
    let nuevos=0, actualizados=0;
    const ahora = new Date().toISOString();
    d.pacientes.forEach(function(p){
      if(!p || !p.dni) return;
      const cl = (typeof p.checklist==='string') ? p.checklist : JSON.stringify(p.checklist||[]);
      const fila = [String(p.dni), p.paciente||'', p.obra_social||'', p.telefono||'', p.mail||'', p.etapa||'', p.diagnostico||'', p.cie10||'', p.procedimiento||'', p.institucion||'', cl.slice(0,40000), p.ultima_consulta||'', p.ultima_consulta_iso||'', p.proxima_cita||'', p.postop_control||'', String(p.caso||'').slice(0,600), ahora];
      const k = String(p.dni);
      if(idx[k]!==undefined){ filas[idx[k]] = fila; actualizados++; } else { filas.push(fila); idx[k]=filas.length-1; nuevos++; }
    });
    if(filas.length){
      sh.getRange(2,1,filas.length,SEG_COLS_.length).setNumberFormat('@').setValues(filas.map(function(r){ return r.map(function(v){ return v===null||v===undefined?'':String(v); }); }));
    }
    return {success:true, nuevos:nuevos, actualizados:actualizados, total:filas.length};
  }catch(e){
    return {success:false, error:'No se pudo guardar el seguimiento: '+e};
  }finally{
    try{ lock.releaseLock(); }catch(_){}
  }
}
function getSeguimiento(){
  const data = hojaSeguimiento_().getDataRange().getValues();
  const items = [];
  for(let i=1;i<data.length;i++){
    const r = data[i]; if(!r[0]) continue;
    let cl = []; try{ cl = r[10] ? JSON.parse(r[10]) : []; }catch(_){ cl = []; }
    items.push({ dni:String(r[0]), paciente:String(r[1]||''), obra_social:String(r[2]||''), telefono:String(r[3]||''), mail:String(r[4]||''),
      etapa:String(r[5]||''), diagnostico:String(r[6]||''), cie10:String(r[7]||''), procedimiento:String(r[8]||''), institucion:String(r[9]||''),
      checklist:cl, ultima_consulta:String(r[11]||''), ultima_consulta_iso:String(r[12]||''), proxima_cita:String(r[13]||''),
      postop_control:String(r[14]||''), caso:String(r[15]||''), actualizado:String(r[16]||'') });
  }
  return {success:true, items:items};
}

// ══════════════════════════════════════════════════════════════════
// PROGRAMACIÓN QUIRÚRGICA (oct 2026, pedido de Luis) — cada cirugía
// programada: fecha, hora, sanatorio, paciente, obra social y tipo de cirugía.
// Alimenta el calendario mensual de OCM ("Agenda de Cirugía"); los miércoles
// desde las 13 hs entran 3–4 turnos quirúrgicos. Todo va con token.
// ══════════════════════════════════════════════════════════════════
function hojaProgramacionQx_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Programacion_Qx');
  if(!sheet){
    sheet = ss.insertSheet('Programacion_Qx');
    sheet.appendRow(['ID','Timestamp','Fecha','Hora','Institucion','DNI','Paciente','Obra_Social','Procedimiento','Estado','Notas']);
    sheet.getRange(1,1,1,11).setBackground('#14532d').setFontColor('white').setFontWeight('bold');
    sheet.getRange('C:D').setNumberFormat('@');
    sheet.setFrozenRows(1);
  }
  return sheet;
}
function getProgramacionQx(params){
  params = params || {};
  const mes = String(params.mes||'').trim(); // 'YYYY-MM' o vacío = todo
  const filas = hojaProgramacionQx_().getDataRange().getValues();
  const items = [];
  for(let i=1;i<filas.length;i++){
    if(String(filas[i][9]||'').toLowerCase()==='cancelado') continue;
    const f = celdaFecha_(filas[i][2]);
    if(mes && f.slice(0,7)!==mes) continue;
    items.push({ id:filas[i][0], fecha:f, hora:celdaHora_(filas[i][3]), institucion:String(filas[i][4]||''), dni:String(filas[i][5]||''),
      paciente:String(filas[i][6]||''), obra_social:String(filas[i][7]||''), procedimiento:String(filas[i][8]||''), notas:String(filas[i][10]||'') });
  }
  items.sort(function(a,b){ return (a.fecha+a.hora).localeCompare(b.fecha+b.hora); });
  return { success:true, items:items };
}
function guardarProgramacionQx(d){
  if(!d) return {success:false, error:'Sin datos'};
  const fecha = String(d.fecha||'').trim(), hora = String(d.hora||'').trim();
  if(!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) return {success:false, error:'Falta la fecha de la cirugía'};
  if(!/^\d{1,2}:\d{2}$/.test(hora)) return {success:false, error:'Falta la hora de la cirugía'};
  if(!String(d.institucion||'').trim()) return {success:false, error:'Falta el sanatorio / institución'};
  if(!String(d.paciente||'').trim()) return {success:false, error:'Falta el paciente'};
  const lock = LockService.getScriptLock();
  try{
    lock.waitLock(10000);
    const sheet = hojaProgramacionQx_();
    const filas = sheet.getDataRange().getValues();
    const horaN = hora.padStart(5,'0');
    // Un mismo sanatorio no puede tener dos cirugías a la misma fecha y hora (salvo que sea la misma fila que se edita)
    for(let i=1;i<filas.length;i++){
      if(String(filas[i][9]||'').toLowerCase()==='cancelado') continue;
      if(d.id && filas[i][0]===d.id) continue;
      if(celdaFecha_(filas[i][2])===fecha && celdaHora_(filas[i][3])===horaN && String(filas[i][4])===String(d.institucion)){
        return {success:false, error:'Ya hay una cirugía en '+d.institucion+' ese día a las '+horaN+' ('+filas[i][6]+')'};
      }
    }
    const fila = [d.id||Utilities.getUuid(), new Date().toISOString(), fecha, horaN, d.institucion, String(d.dni||''), d.paciente, d.obra_social||'', d.procedimiento||'', 'programada', d.notas||''];
    if(d.id){
      for(let i=1;i<filas.length;i++){
        if(filas[i][0]===d.id){ sheet.getRange(i+1,1,1,11).setValues([fila]); return {success:true, id:d.id, actualizado:true}; }
      }
    }
    sheet.getRange(sheet.getLastRow()+1,1,1,11).setNumberFormat('@').setValues([fila]);
    return {success:true, id:fila[0]};
  }catch(e){
    return {success:false, error:'No se pudo guardar la cirugía: '+e};
  }finally{
    try{ lock.releaseLock(); }catch(_){}
  }
}
function eliminarProgramacionQx(d){
  if(!d || !d.id) return {success:false, error:'Falta el ID'};
  const sheet = hojaProgramacionQx_();
  const filas = sheet.getDataRange().getValues();
  for(let i=1;i<filas.length;i++){
    if(filas[i][0]===d.id){ sheet.getRange(i+1,10).setValue('cancelado'); return {success:true}; }
  }
  return {success:false, error:'No se encontró esa cirugía'};
}

// Reserva un turno puntual. Vuelve a comprobar, en el momento de guardar,
// que ese Consultorio+Fecha+Hora no se haya ocupado mientras tanto (por
// ejemplo, la secretaria de un consultorio y el médico desde la
// videollamada mirando el mismo horario libre al mismo tiempo) — el que
// llega primero se queda con el turno, el segundo recibe un error claro en
// vez de pisar la reserva.
// d = {consultorio, fecha:'yyyy-MM-dd', hora:'HH:mm', dni, nombre, telefono,
//      email, origen, notas}
function reservarTurno(d){
  if(!d || !String(d.consultorio||'').trim()) return {success:false, error:'Falta el consultorio'};
  const esAgendaLibre = CONSULTORIOS_AGENDA_LIBRE_.indexOf(d.consultorio)>-1;
  if(!CONSULTORIOS_HORARIOS_[d.consultorio] && !esAgendaLibre) return {success:false, error:'Consultorio no reconocido'};
  if(!String(d.fecha||'').trim() || !String(d.hora||'').trim()) return {success:false, error:'Falta la fecha o el horario'};
  if(!String(d.nombre||'').trim()) return {success:false, error:'Falta el nombre del paciente'};
  if(!String(d.telefono||'').trim() && !String(d.email||'').trim()){
    return {success:false, error:'Dejá al menos un teléfono o un mail de contacto del paciente'};
  }

  // BUGFIX (sep 2026): antes se chequeaba si el horario ya estaba tomado y
  // recién después se agregaba la fila, sin ningún bloqueo — si la
  // secretaria (Turnos_Secretaria.html) y un paciente reservando online
  // apretaban "Confirmar" para el mismo horario casi al mismo tiempo, los
  // dos pedidos podían leer la planilla ANTES de que el otro terminara de
  // escribir, los dos veían el horario libre, y los dos quedaban
  // "confirmado" — doble reserva para el mismo horario, sin ningún aviso.
  // El lock hace que el segundo pedido espere a que el primero termine de
  // escribir antes de leer, así el segundo si o si encuentra el horario ya
  // ocupado y lo rechaza en vez de duplicarlo.
  const lock = LockService.getScriptLock();
  let id;
  // SEÑA (oct 2026): se calcula antes de tomar el lock, no depende de la
  // hoja — ver CONSULTORIOS_CON_SENA_ más arriba.
  const requiereSena = CONSULTORIOS_CON_SENA_.indexOf(d.consultorio) > -1;
  const montoSena = requiereSena ? montoSenaTurno_() : '';
  try{
    lock.waitLock(10000);
    const sheet = hojaTurnosAgenda_();
    // Columnas de seña y recordatorios — se agregan a demanda, igual que en
    // Cola_Mensajes, para no romper la hoja ya desplegada en producción.
    const colSenaReq = asegurarColumnaColaMensajes_(sheet, 'Sena_Requerida');
    const colSenaEstado = asegurarColumnaColaMensajes_(sheet, 'Sena_Estado');
    const colSenaMonto = asegurarColumnaColaMensajes_(sheet, 'Sena_Monto');
    asegurarColumnaColaMensajes_(sheet, 'Sena_Comprobante_URL');
    const colRecDia = asegurarColumnaColaMensajes_(sheet, 'Recordatorio_DiaAntes_Enviado');
    const colRec3hs = asegurarColumnaColaMensajes_(sheet, 'Recordatorio_3hs_Enviado');
    const anchoFila = Math.max(13, colSenaReq, colSenaEstado, colSenaMonto, colRecDia, colRec3hs);
    const filas = sheet.getDataRange().getValues();
    for(let i=1;i<filas.length;i++){
      const estado = String(filas[i][11]||'').toLowerCase();
      if(estado === 'cancelado') continue;
      if(filas[i][2]===d.consultorio && filas[i][3]===d.fecha && filas[i][4]===d.hora){
        return {success:false, error:'Ese horario ya se acaba de reservar — elegí otro.'};
      }
    }
    id = Utilities.getUuid();
    const fila = new Array(anchoFila).fill('');
    fila[0]=id; fila[1]=new Date().toISOString(); fila[2]=d.consultorio; fila[3]=d.fecha; fila[4]=d.hora;
    fila[5]=duracionTurnoConsultorio_(d.consultorio);
    fila[6]=d.dni||''; fila[7]=d.nombre||''; fila[8]=d.telefono||''; fila[9]=d.email||'';
    fila[10]=d.origen||'Sin especificar'; fila[11]='confirmado'; fila[12]=d.notas||'';
    fila[colSenaReq-1] = requiereSena ? 'Sí' : 'No';
    fila[colSenaEstado-1] = requiereSena ? 'Pendiente' : 'No aplica';
    fila[colSenaMonto-1] = montoSena;
    sheet.getRange(sheet.getLastRow()+1, 1, 1, fila.length).setValues([fila]);
  }catch(e){
    return {success:false, error:'No se pudo confirmar el turno (sistema ocupado) — probá de nuevo.'};
  }finally{
    lock.releaseLock();
  }

  if(String(d.email||'').trim()){
    try{
      const linkGestion = baseSimm_() + 'Gestion_Turno.html?id=' + id;
      let cuerpoExtra;
      if(requiereSena){
        cuerpoExtra =
          '\n\nPara asegurar este turno te pedimos una seña de $' + montoSena.toLocaleString('es-AR') + ', por transferencia, a este alias:\n' +
          ALIAS_TRANSFERENCIA_SENA_ + '\n\n' +
          'Después de transferir, subí el comprobante acá (es el mismo link para cancelar si hiciera falta):\n' + linkGestion + '\n\n' +
          'Si cancelás el turno, se pierde la seña. Si lo cancela el médico, la seña queda como crédito para tu próximo turno.';
      } else {
        cuerpoExtra = '\n\nSi necesitás cancelar, podés hacerlo acá:\n' + linkGestion;
      }
      MailApp.sendEmail(d.email,
        'Turno confirmado — Dr. Muratori',
        'Hola ' + (d.nombre||'') + ',\n\n' +
        'Tu turno quedó confirmado:\n\n' +
        'Día: ' + formatearFechaHoraLegible_(d.fecha, d.hora) + '\n' +
        'Lugar: ' + d.consultorio +
        (d.notas ? ('\n\nNotas: ' + d.notas) : '') +
        cuerpoExtra + '\n\n' +
        'Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943',
        { name: 'Dr. Luis Alberto Muratori' });
    }catch(e){
      // El turno ya quedó guardado — que falle el mail de confirmación no
      // debe hacer perder la reserva.
    }
  }

  return { success:true, id: id, requiere_sena: requiereSena, monto_sena: montoSena };
}

// Libera un turno (por si el paciente cancela o hay que reprogramar) — no
// borra la fila, la marca 'cancelado' para que el horario vuelva a
// aparecer como disponible sin perder el registro de lo que pasó.
// canceladoPor: 'paciente' (default — link del mail/Gestion_Turno.html) o
// 'medico' (panel de la secretaria, Turnos_Secretaria.html). Afecta la
// seña (oct 2026, pedido de Luis): si cancela el paciente y ya la había
// pagado, se pierde; si cancela el médico, queda como crédito para el
// próximo turno. El sistema nunca devuelve plata sola — eso sigue siendo
// a mano si hiciera falta.
function cancelarTurno(d){
  if(!d || !String(d.id||'').trim()) return {success:false, error:'Falta el ID del turno'};
  const canceladoPor = (d.canceladoPor === 'medico') ? 'medico' : 'paciente';
  const sheet = hojaTurnosAgenda_();
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const colSenaEstado = headers.indexOf('Sena_Estado'); // -1 si es una reserva vieja, de antes de la seña
  const filas = sheet.getDataRange().getValues();
  for(let i=1;i<filas.length;i++){
    if(filas[i][0] === d.id){
      sheet.getRange(i+1, 12).setValue('cancelado');
      let senaResultado = null;
      // "Comprobante recibido — verificando" cuenta igual que "Pagada" para
      // esta regla: el paciente ya transfirió (aunque nadie lo haya
      // confirmado todavía), así que la plata igual está en juego.
      const estadoPagoActual = colSenaEstado > -1 ? String(filas[i][colSenaEstado]||'') : '';
      if(estadoPagoActual === 'Pagada' || estadoPagoActual === 'Comprobante recibido — verificando'){
        senaResultado = (canceladoPor === 'medico') ? 'Reconocida — crédito próximo turno' : 'Perdida';
        sheet.getRange(i+1, colSenaEstado+1).setValue(senaResultado);
      }
      return {success:true, sena_resultado: senaResultado};
    }
  }
  return {success:false, error:'No se encontró ese turno'};
}

// Lectura pública y acotada de un turno por su ID (el mismo UUID ya
// funciona como token, no es adivinable) — para Gestion_Turno.html, la
// pantalla donde el paciente sube el comprobante de seña o cancela.
function getTurnoPublico(d){
  const id = (d && d.id) || '';
  if(!String(id).trim()) return {success:false, error:'Falta el ID del turno'};
  const sheet = hojaTurnosAgenda_();
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const col = function(nombre){ return headers.indexOf(nombre); };
  const colSenaReq = col('Sena_Requerida'), colSenaEstado = col('Sena_Estado'), colSenaMonto = col('Sena_Monto');
  const filas = sheet.getDataRange().getValues();
  for(let i=1;i<filas.length;i++){
    if(filas[i][0] === id){
      return {
        success:true,
        consultorio: filas[i][2], fecha: filas[i][3], hora: filas[i][4], nombre: filas[i][7],
        estado: String(filas[i][11]||''),
        sena_requerida: colSenaReq>-1 && filas[i][colSenaReq]==='Sí',
        sena_estado: colSenaEstado>-1 ? filas[i][colSenaEstado] : 'No aplica',
        sena_monto: colSenaMonto>-1 ? filas[i][colSenaMonto] : ''
      };
    }
  }
  return {success:false, error:'No se encontró ese turno'};
}

// d = {id, comprobanteBase64} — el paciente sube la foto del comprobante
// de la seña desde Gestion_Turno.html. Deja el estado en "verificando" en
// vez de "Pagada" directamente: alguien del consultorio confirma que el
// depósito realmente llegó (ver verificarSenaTurno) antes de darlo por
// pagado — mismo criterio que ya se usa en el QR de asesoramiento médico.
function subirComprobanteSenaTurno(d){
  if(!d || !String(d.id||'').trim()) return {success:false, error:'Falta el ID del turno'};
  if(!d.comprobanteBase64) return {success:false, error:'Falta adjuntar el comprobante'};
  const sheet = hojaTurnosAgenda_();
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const colSenaEstado = headers.indexOf('Sena_Estado');
  const colSenaComprobante = headers.indexOf('Sena_Comprobante_URL');
  if(colSenaEstado===-1 || colSenaComprobante===-1) return {success:false, error:'Este turno no tiene seña configurada'};
  const filas = sheet.getDataRange().getValues();
  for(let i=1;i<filas.length;i++){
    if(filas[i][0] === d.id){
      const url = guardarComprobantePagoQR_(d.comprobanteBase64, filas[i][7] || 'paciente');
      sheet.getRange(i+1, colSenaComprobante+1).setValue(url || 'adjuntado — revisar mail si no se guardó el archivo');
      sheet.getRange(i+1, colSenaEstado+1).setValue('Comprobante recibido — verificando');
      try{
        MailApp.sendEmail('lmuratori@gmail.com',
          'Comprobante de seña recibido — ' + (filas[i][7]||''),
          'Llegó un comprobante de seña para el turno de ' + (filas[i][7]||'') + ' en ' + filas[i][2] + ', el ' + filas[i][3] + ' ' + filas[i][4] + ' hs.\n\n' +
          'Verificalo y marcalo como pagado desde Turnos_Secretaria.html (apartado de señas pendientes de verificar).' +
          (url ? ('\n\nComprobante: ' + url) : ''));
      }catch(e){}
      return {success:true};
    }
  }
  return {success:false, error:'No se encontró ese turno'};
}

// d = {id} — el médico/secretaria confirma que la transferencia de la
// seña realmente llegó a la cuenta.
function verificarSenaTurno(d){
  if(!d || !String(d.id||'').trim()) return {success:false, error:'Falta el ID del turno'};
  const sheet = hojaTurnosAgenda_();
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const colSenaEstado = headers.indexOf('Sena_Estado');
  if(colSenaEstado===-1) return {success:false, error:'Este turno no tiene seña configurada'};
  const filas = sheet.getDataRange().getValues();
  for(let i=1;i<filas.length;i++){
    if(filas[i][0] === d.id){
      sheet.getRange(i+1, colSenaEstado+1).setValue('Pagada');
      return {success:true};
    }
  }
  return {success:false, error:'No se encontró ese turno'};
}

// Lista de turnos con comprobante de seña recibido pero todavía sin
// verificar — la usa el apartado correspondiente en Turnos_Secretaria.html.
function getSenasPendientesVerificar(){
  const sheet = hojaTurnosAgenda_();
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const colSenaEstado = headers.indexOf('Sena_Estado');
  const colSenaComprobante = headers.indexOf('Sena_Comprobante_URL');
  if(colSenaEstado===-1) return {success:true, pendientes:[]};
  const filas = sheet.getDataRange().getValues();
  const pendientes = [];
  for(let i=1;i<filas.length;i++){
    if(String(filas[i][colSenaEstado]||'') === 'Comprobante recibido — verificando'){
      pendientes.push({
        id: filas[i][0], consultorio: filas[i][2], fecha: filas[i][3], hora: filas[i][4],
        nombre: filas[i][7], comprobante_url: colSenaComprobante>-1 ? filas[i][colSenaComprobante] : ''
      });
    }
  }
  return {success:true, pendientes: pendientes};
}

// Próximos turnos confirmados (hasta 60), para que el médico/secretaria
// pueda cancelar uno desde Turnos_Secretaria.html — esa cancelación usa
// canceladoPor:'medico', que reconoce la seña como crédito en vez de
// perderla (ver cancelarTurno más arriba).
function getProximosTurnos(){
  const sheet = hojaTurnosAgenda_();
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const colSenaReq = headers.indexOf('Sena_Requerida'), colSenaEstado = headers.indexOf('Sena_Estado');
  const filas = sheet.getDataRange().getValues();
  const ahora = new Date();
  const proximos = [];
  for(let i=1;i<filas.length;i++){
    const fila = filas[i];
    if(String(fila[11]||'').toLowerCase() !== 'confirmado') continue;
    const fechaHora = fechaHoraTurno_(fila[3], fila[4]);
    if(fechaHora && fechaHora.getTime() < ahora.getTime()) continue;
    proximos.push({
      id: fila[0], consultorio: fila[2], fecha: fila[3], hora: fila[4], nombre: fila[7],
      sena_requerida: colSenaReq>-1 && fila[colSenaReq]==='Sí',
      sena_estado: colSenaEstado>-1 ? fila[colSenaEstado] : 'No aplica'
    });
  }
  proximos.sort(function(a,b){ return (a.fecha+a.hora).localeCompare(b.fecha+b.hora); });
  return {success:true, proximos: proximos.slice(0,60)};
}

// ══════════════════════════════════════════════════════════════════
// RECORDATORIOS AUTOMÁTICOS DE TURNO (oct 2026, pedido de Luis) — corre
// cada hora (activar una sola vez con crearTriggerRecordatoriosTurnos,
// desde el editor). Para cada turno confirmado con mail:
//   - Consultorios con seña (CONSULTORIOS_CON_SENA_): recordatorio el día
//     anterior (con el alias si la seña sigue pendiente) y otro 3 horas
//     antes.
//   - Hospital Sícoli (sin seña): solo el recordatorio del día anterior,
//     pidiendo que confirme si va a poder asistir.
// Cada recordatorio se manda una sola vez (columnas Recordatorio_*_Enviado)
// y siempre lleva el link de Gestion_Turno.html para cancelar.
// ══════════════════════════════════════════════════════════════════
function enviarRecordatoriosTurnos(){
  const sheet = hojaTurnosAgenda_();
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const col = function(nombre){ return headers.indexOf(nombre); };
  const colSenaReq = col('Sena_Requerida'), colSenaEstado = col('Sena_Estado');
  const colRecDia = col('Recordatorio_DiaAntes_Enviado'), colRec3hs = col('Recordatorio_3hs_Enviado');
  if(colRecDia===-1 || colRec3hs===-1) return; // hoja vieja sin estas columnas — no hace nada, no rompe

  const filas = sheet.getDataRange().getValues();
  const ahora = new Date();
  for(let i=1;i<filas.length;i++){
    const fila = filas[i];
    if(String(fila[11]||'').toLowerCase() !== 'confirmado') continue;
    if(!String(fila[9]||'').trim()) continue; // sin mail no se puede recordar
    const fechaHora = fechaHoraTurno_(fila[3], fila[4]);
    if(!fechaHora) continue;
    const horasFaltantes = (fechaHora.getTime() - ahora.getTime()) / 3600000;
    if(horasFaltantes < 0) continue;

    const requiereSena = colSenaReq>-1 && fila[colSenaReq] === 'Sí';
    const senaEstado = colSenaEstado>-1 ? String(fila[colSenaEstado]||'') : '';
    const yaDiaAntes = String(fila[colRecDia]||'') === 'Sí';
    const ya3hs = String(fila[colRec3hs]||'') === 'Sí';
    const linkGestion = baseSimm_() + 'Gestion_Turno.html?id=' + fila[0];

    // Ventana de 18 a 30 hs antes (12 hs de ancho): suficiente para que el
    // trigger cada hora no se la pierda, pero acotada para que "mañana" en
    // el texto sea siempre cierto — si alguien reserva con menos de 18 hs
    // de aviso, directamente no hay "día antes" que recordar, y pasa
    // derecho al de 3 horas (si aplica).
    if(!yaDiaAntes && horasFaltantes <= 30 && horasFaltantes > 18){
      const diaLegible = formatearFechaHoraLegible_(fila[3], fila[4]);
      let cuerpo = 'Hola ' + (fila[7]||'') + ',\n\nTe recordamos tu turno:\n\n' +
        'Día: ' + diaLegible + '\nLugar: ' + fila[2] + '\n\n';
      if(requiereSena && senaEstado === 'Pendiente'){
        cuerpo += 'Todavía no registramos tu seña. Por favor transferí a este alias:\n' + ALIAS_TRANSFERENCIA_SENA_ +
          '\ny subí el comprobante acá: ' + linkGestion + '\n\nSi no llega la seña, el turno puede liberarse.\n\n';
      } else if(requiereSena && senaEstado === 'Comprobante recibido — verificando'){
        cuerpo += 'Ya recibimos tu comprobante de seña y lo estamos verificando — no hace falta que hagas nada más.\n\n';
      } else if(requiereSena){
        cuerpo += 'Tu seña ya está confirmada.\n\n';
      } else {
        cuerpo += 'Por favor confirmanos si vas a poder asistir.\n\n';
      }
      cuerpo += 'Si necesitás cancelar: ' + linkGestion + '\n\n' + 'Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943';
      try{
        MailApp.sendEmail(fila[9], 'Recordatorio de turno — ' + diaLegible + ' — Dr. Muratori', cuerpo, { name: 'Dr. Luis Alberto Muratori' });
        sheet.getRange(i+1, colRecDia+1).setValue('Sí');
      }catch(e){}
      continue;
    }

    // Ventana de 2 a 4 hs antes (solo en consultorios con seña) — si se
    // reservó con menos de 2 hs de aviso, no sale este recordatorio: ya no
    // tendría sentido avisar "en 3 horas" cuando en realidad quedan 40
    // minutos.
    if(requiereSena && !ya3hs && horasFaltantes <= 4 && horasFaltantes > 2){
      const diaLegible3hs = formatearFechaHoraLegible_(fila[3], fila[4]);
      const cuerpo = 'Hola ' + (fila[7]||'') + ',\n\nTu turno es en pocas horas:\n\n' +
        'Día: ' + diaLegible3hs + '\nLugar: ' + fila[2] + '\n\n' +
        'Si necesitás cancelar: ' + linkGestion + ' (a esta altura se pierde la seña)\n\n' +
        'Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943';
      try{
        MailApp.sendEmail(fila[9], 'Recordatorio de turno — ' + diaLegible3hs + ' — Dr. Muratori', cuerpo, { name: 'Dr. Luis Alberto Muratori' });
        sheet.getRange(i+1, colRec3hs+1).setValue('Sí');
      }catch(e){}
    }
  }
}

// Día de la semana + fecha + hora, siempre explícitos (pedido de Luis, oct
// 2026: "los turnos no deben expresar hoy o mañana, es el día, la fecha y
// la hora para que sea inconfundible") — se usa en TODOS los mails de
// turno (confirmación y los dos recordatorios), para que nunca haya
// ambigüedad de a qué día se refiere.
const DIAS_SEMANA_ES_ = ['domingo','lunes','martes','miércoles','jueves','viernes','sábado'];
const MESES_ES_ = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
function formatearFechaHoraLegible_(fechaStr, horaStr){
  const d = fechaHoraTurno_(fechaStr, horaStr);
  if(!d) return fechaStr + ', ' + horaStr + ' hs'; // respaldo si la fecha vino en un formato raro
  return DIAS_SEMANA_ES_[d.getDay()] + ' ' + d.getDate() + ' de ' + MESES_ES_[d.getMonth()] + ', ' + horaStr + ' hs';
}

// Fecha ("AAAA-MM-DD") + Hora ("HH:MM") de la hoja → objeto Date, en la
// zona horaria configurada en el proyecto de Apps Script.
function fechaHoraTurno_(fechaStr, horaStr){
  try{
    const pf = String(fechaStr).split('-').map(Number);
    const ph = String(horaStr).split(':').map(Number);
    if(pf.length<3 || ph.length<2) return null;
    return new Date(pf[0], pf[1]-1, pf[2], ph[0], ph[1]);
  }catch(e){ return null; }
}

// Ejecutar UNA SOLA VEZ a mano desde el editor de Apps Script (▶ Ejecutar)
// para activar los recordatorios automáticos — mismo patrón que
// crearTriggerRecordatorioSemanal.
function crearTriggerRecordatoriosTurnos(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if(t.getHandlerFunction()==='enviarRecordatoriosTurnos') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('enviarRecordatoriosTurnos')
    .timeBased().everyHours(1).create();
}

function testGeminiDirecto(){
  try{
    const resp = UrlFetchApp.fetch(getGeminiUrl_(), {
      method:'post', contentType:'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: 'Decime hola en una palabra' }] }],
        generationConfig: {
          thinkingConfig: { thinkingLevel: 'minimal' }
        }
      }),
      muteHttpExceptions: true
    });
    const codigo = resp.getResponseCode();
    const texto = resp.getContentText();
    MailApp.sendEmail('lmuratori@gmail.com', 'Test Gemini — resultado (optimizado)', 'Código HTTP: ' + codigo + '\n\nRespuesta completa:\n' + texto);
  }catch(e){
    MailApp.sendEmail('lmuratori@gmail.com', 'Test Gemini — ERROR', 'No se pudo ni siquiera llamar a Gemini: ' + e + '\n\nSi el error dice "Falta configurar GEMINI_API_KEY", hay que correr configurarGeminiAPI() primero.');
  }
}

// ══════════════════════════════════════════════════════════════════
// Limpieza de la casilla: mails YA RESPONDIDOS desde Agenda de Mensajes.
// Los de más de 7 días con la etiqueta SIM-M/Procesado y cuya fila en
// Cola_Mensajes figura "respondido": si NO tienen adjuntos van a la papelera
// (Gmail la vacía sola a los 30 días); si tienen adjuntos (estudios, informes)
// solo se archivan, para no perderlos. Lo que está pendiente de revisión o no
// figura en la cola NUNCA se toca. Activar una sola vez con
// crearTriggerLimpiezaProcesados().
// ══════════════════════════════════════════════════════════════════
function limpiarMailsRespondidos(){
  const sheet = hojaColaMensajes_();
  const datos = sheet.getDataRange().getValues();
  const respondidos = {};
  for(let i=1;i<datos.length;i++){
    if(datos[i][10]==='respondido' && datos[i][12]) respondidos[String(datos[i][12])] = true;
  }
  const hilos = GmailApp.search('label:"' + LABEL_PROCESADO + '" older_than:7d', 0, 100);
  let papelera = 0, archivados = 0;
  hilos.forEach(function(h){
    if(!respondidos[h.getId()]) return;
    const conAdjuntos = h.getMessages().some(function(m){ return m.getAttachments().length > 0; });
    if(conAdjuntos){ h.moveToArchive(); archivados++; }
    else { h.moveToTrash(); papelera++; }
  });
  Logger.log('limpiarMailsRespondidos: ' + papelera + ' a la papelera, ' + archivados + ' archivados (con adjuntos).');
  return {papelera:papelera, archivados:archivados};
}
// Limpia de Agenda de Mensajes lo que ya entró y NO es médico (por reglas de remitente).
// Los pasa a estado "descartado" (no se borran de la hoja) y deja el mail de Gmail
// otra vez como NO LEÍDO, sin la etiqueta Procesado y con SIM-M/No-medico.
// Nunca toca filas con DNI detectado, ni las que vienen de QR/app (columna Origen).
// Pregunta corta a Gemini: ¿es un mail médico? Devuelve true/false, o null si no pudo decidir
// (ante la duda se conserva el mail).
function esMedicoIA_(asunto, cuerpo){
  try{
    const prompt = 'Sos el asistente del Dr. Luis Alberto Muratori, traumatólogo. Decidí si este mail es médico para su agenda de mensajes de pacientes. esMedico: true SOLO si es un mensaje de un paciente, de un familiar o de un colega sobre un problema de salud o un caso clínico: consulta, síntomas, estudios, informes o certificados de un paciente, pedido de turno o receta, plantillas ortopédicas, o respuesta a un mail del doctor. false si es administrativo o comercial: facturas, honorarios, retenciones, recupero de gastos, contratos, recordatorios de turnos de un centro médico, trámites de la clínica donde trabaja, cadenas o saludos masivos, publicidad, redes sociales, compras, suscripciones, webinars, newsletters, notificaciones automáticas, banca o spam.\n\nASUNTO: ' + asunto + '\nCUERPO:\n' + cuerpo + '\n\nDevolvé SOLO JSON: {"esMedico": true o false}';
    const resp = UrlFetchApp.fetch(getGeminiUrl_(), {method:'post', contentType:'application/json', muteHttpExceptions:true,
      payload: JSON.stringify({contents:[{parts:[{text:prompt}]}], generationConfig:{thinkingConfig:{thinkingLevel:'minimal'}}})});
    const data = JSON.parse(resp.getContentText());
    const j = JSON.parse(data.candidates[0].content.parts[0].text.replace(/```json|```/g,'').trim());
    return (typeof j.esMedico === 'boolean') ? j.esMedico : null;
  }catch(e){ return null; }
}

const FILTRO_MAILS_VERSION = '2026-10-12-v20';
// Diagnóstico: dice qué versión del filtro está pegada y cuántos mensajes pendientes
// hay, cuántos descartaría por reglas y de qué remitentes. No modifica nada.
// Limpia de la Agenda lo que quedó pendiente de ANTES y que no fue dirigido a
// consultas.sism (mails de la casilla personal). No toca los de QR/app ni los que
// no se pueden verificar. Primero correr _VER (solo informa) y después _APLICAR.
function descartarPendientesNoConsultas_VER(){ return descartarPendientesNoConsultas_(false); }
function descartarPendientesNoConsultas_APLICAR(){ return descartarPendientesNoConsultas_(true); }
function descartarPendientesNoConsultas_(aplicar){
  const sheet = hojaColaMensajes_();
  const datos = sheet.getDataRange().getValues();
  const enc = datos[0] || [];
  const iOrigen = enc.indexOf('Origen');
  let revisados = 0, descartados = 0, mantenidos = 0, sinVerificar = 0;
  const inicio = Date.now();
  for(let i=1;i<datos.length;i++){
    if(Date.now() - inicio > 280000){ Logger.log('Corte por tiempo: volvé a correr la función para seguir.'); break; }
    const f = datos[i];
    if(f[10] !== 'pendiente_revision') continue;
    if(iOrigen >= 0 && f[iOrigen]) { mantenidos++; continue; }          // QR / app: siempre quedan
    const tid = String(f[12]||'');
    if(!tid){ sinVerificar++; continue; }
    revisados++;
    try{
      const hilo = GmailApp.getThreadById(tid);
      if(!hilo){ sinVerificar++; continue; }
      const dest = hilo.getMessages().map(function(m){ return (m.getTo()||'') + ' ' + (m.getCc()||''); }).join(' ').toLowerCase();
      if(dest.indexOf(MAIL_CONSULTAS_) !== -1){ mantenidos++; continue; }
      descartados++;
      Logger.log((aplicar ? 'DESCARTADO: ' : 'se descartaría: ') + String(f[2]).slice(0,50) + ' | ' + String(f[3]).slice(0,50));
      if(aplicar) sheet.getRange(i+1, 11).setValue('descartado');
    }catch(e){ sinVerificar++; }
  }
  Logger.log('Revisados: ' + revisados + ' · ' + (aplicar ? 'descartados: ' : 'se descartarían: ') + descartados + ' · quedan: ' + mantenidos + ' · sin poder verificar (quedan): ' + sinVerificar);
  return {revisados:revisados, descartados:descartados, mantenidos:mantenidos, sinVerificar:sinVerificar};
}
function diagnosticoFiltroMails(){
  const sheet = hojaColaMensajes_();
  const datos = sheet.getDataRange().getValues();
  const enc = datos[0] || [];
  const iOrigen = enc.indexOf('Origen');
  let pend = 0, porReglas = 0, conOrigen = 0;
  const muestra = [];
  for(let i=1;i<datos.length;i++){
    const f = datos[i];
    if(f[10] !== 'pendiente_revision') continue;
    pend++;
    if(iOrigen >= 0 && f[iOrigen]){ conOrigen++; continue; }
    if(esMailPropioOSistema_(f[2], f[3]) || esRemitenteNoMedico_(f[2], '', f[3])) porReglas++;
    else if(muestra.length < 15) muestra.push(String(f[2]).slice(0,60) + ' | ' + String(f[3]).slice(0,50));
  }
  Logger.log('Versión del filtro: ' + FILTRO_MAILS_VERSION);
  Logger.log('Filas en la cola: ' + (datos.length-1) + ' · pendientes de revisión: ' + pend + ' · de QR/app: ' + conOrigen + ' · se descartarían por reglas: ' + porReglas);
  Logger.log('Cuántas decidiría Gemini: ' + (pend - conOrigen - porReglas));
  try{
    const yoD = String(Session.getEffectiveUser().getEmail()||'').toLowerCase();
    const q = filtroSoloConsultas_(yoD);
    Logger.log('Cuenta que corre: ' + yoD + ' · modo: ' + (q ? 'SOLO mails dirigidos a ' + MAIL_CONSULTAS_ : 'toda la bandeja'));
    const hConsultas = GmailApp.search(q + 'in:inbox newer_than:7d', 0, 50);
    Logger.log('Hilos de los últimos 7 días que el sistema tomaría: ' + hConsultas.length + (hConsultas.length===0 && q ? ' (si esperabas mails: revisá que consultas.sism reenvíe a esta casilla)' : ''));
  }catch(e){ Logger.log('No se pudo contar los hilos de consultas: ' + e); }
  muestra.forEach(function(x){ Logger.log('  para Gemini: ' + x); });
  return {version:FILTRO_MAILS_VERSION, pendientes:pend, porReglas:porReglas};
}

function descartarMensajesNoMedicosCola(){
  Logger.log('Versión del filtro: ' + FILTRO_MAILS_VERSION);
  const sheet = hojaColaMensajes_();
  const datos = sheet.getDataRange().getValues();
  const enc = datos[0] || [];
  const iOrigen = enc.indexOf('Origen');
  const iDni = enc.indexOf('DNI_Detectado');
  let descartados = 0, llamadasIA = 0, cortado = false;
  const t0 = Date.now();
  const restantes = [];
  for(let i=1;i<datos.length;i++){
    const f = datos[i];
    if(f[10] !== 'pendiente_revision') continue;
    if(Date.now() - t0 > 300000){ cortado = true; break; }   // tope de Apps Script: 6 min
    if(iOrigen >= 0 && f[iOrigen]) continue;
    const propio = esMailPropioOSistema_(f[2], f[3]);
    if(!propio && iDni >= 0 && f[iDni]) continue;
    let noMed = propio || esRemitenteNoMedico_(f[2], '', f[3]);
    if(!noMed && llamadasIA < 150){ llamadasIA++; noMed = (esMedicoIA_(f[3], String(f[4]||'').slice(0,1500)) === false); }
    if(noMed){
      sheet.getRange(i+1, 11).setValue('descartado');
      descartados++;
      try{
        const h = f[12] ? GmailApp.getThreadById(String(f[12])) : null;
        if(h){
          const lp = GmailApp.getUserLabelByName(LABEL_PROCESADO);
          if(lp) h.removeLabel(lp);
          let ln = GmailApp.getUserLabelByName(LABEL_NO_MEDICO);
          if(!ln) ln = GmailApp.createLabel(LABEL_NO_MEDICO);
          h.addLabel(ln);
          h.markUnread();
        }
      }catch(e){ Logger.log('descartar: no pude restaurar el hilo ' + f[12] + ': ' + e); }
    } else {
      restantes.push(String(f[2]) + ' | ' + String(f[3]));
    }
  }
  Logger.log('descartarMensajesNoMedicosCola: ' + descartados + ' descartados. Quedan para revisar a mano: ' + restantes.length + (cortado ? ' — SE CORTÓ POR TIEMPO: ejecutala de nuevo para seguir.' : ''));
  restantes.forEach(function(x){ Logger.log('  queda: ' + x); });
  return {descartados:descartados, restantes:restantes.length};
}

function crearTriggerLimpiezaProcesados(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if(t.getHandlerFunction()==='limpiarMailsRespondidos') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('limpiarMailsRespondidos').timeBased().everyDays(1).atHour(4).create();
  Logger.log('Disparador creado: limpiarMailsRespondidos corre todos los días a las 4 AM.');
}

// ══════════════════════════════════════════════════════════════════
// Disparador automático de procesarBandejaEntrada
// Ejecutar UNA SOLA VEZ manualmente desde el editor (▶ Ejecutar).
// ══════════════════════════════════════════════════════════════════
function crearTriggerProcesarBandejaEntrada(){
  ScriptApp.getProjectTriggers().forEach(t=>{
    if(t.getHandlerFunction()==='procesarBandejaEntrada') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('procesarBandejaEntrada')
    .timeBased()
    .everyMinutes(15)
    .create();
  Logger.log('Disparador creado: procesarBandejaEntrada va a correr cada 15 minutos.');
}

// ══════════════════════════════════════════════════════════════════
// PARCHE — Sincronización de Catastro de Empresas + cruce de Investigación
// Antes el Catastro vivía SOLO en IndexedDB del navegador de OCM: nunca
// llegaba a la nube, así que no había forma real de compartir el cruce de
// Investigación con nadie más (aunque el Dashboard de OCM lo mostrara).
// Esto agrega una copia del catastro en una hoja compartida y replica acá
// el mismo cruce por reglas que antes solo corría en el navegador del
// médico, para que KSM (y futuras apps de colegas) lo puedan consultar.
// ══════════════════════════════════════════════════════════════════
function sheetCatastroEmpresas_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Catastro_Empresas');
  if(!sh){
    sh = ss.insertSheet('Catastro_Empresas');
    sh.appendRow(['Timestamp','Nombre','Localidad','Agentes_Quimicos_JSON','Agentes_Fisicos_JSON','Agentes_Biologicos_JSON','Registrada']);
    sh.getRange(1,1,1,7).setBackground('#0f172a').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  asegurarColumnasCatastro_(sh);
  return sh;
}
// Columnas de vía hídrica (afluente/canal de riego) — se agregan a la hoja
// ya existente sin tocar los registros que ya había.
function asegurarColumnasCatastro_(sh){
  let headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const faltantes = ['Afluente','Afluente_Toxico','Afluente_Organico'].filter(h=>headers.indexOf(h)===-1);
  faltantes.forEach(h=>{
    sh.getRange(1, sh.getLastColumn()+1).setValue(h).setBackground('#0f172a').setFontColor('white').setFontWeight('bold');
  });
  return sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
}
// La llama OCM cada vez que se agrega una empresa al Catastro (además de
// guardarla en su IndexedDB local, como ya hacía).
// Upsert por Nombre+Localidad (antes hacía appendRow siempre, así que
// guardar la misma empresa dos veces —o correr la migración de datos
// locales más de una vez— duplicaba filas en el Catastro).
function sincronizarEmpresaCatastro(d){
  if(!d || !d.nombre || !d.localidad) return{success:false, error:'Falta nombre o localidad de la empresa'};
  const sh = sheetCatastroEmpresas_();
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const fila = new Array(headers.length).fill('');
  const set=(nombreCol,valor,destino)=>{const i=headers.indexOf(nombreCol);if(i>-1)destino[i]=valor;};
  const valores = {};
  set('Timestamp', new Date().toISOString(), valores);
  set('Nombre', d.nombre, valores);
  set('Localidad', d.localidad, valores);
  set('Agentes_Quimicos_JSON', JSON.stringify(d.agentes_quimicos||[]), valores);
  set('Agentes_Fisicos_JSON', JSON.stringify(d.agentes_fisicos||[]), valores);
  set('Agentes_Biologicos_JSON', JSON.stringify(d.agentes_biologicos||[]), valores);
  set('Registrada', d.registrada!==false ? 'SI':'NO', valores);
  // Vía hídrica: descarga de efluentes a canal de riego. El contenido
  // orgánico y/o tóxico puede afectar a pacientes de OTRA localidad,
  // aguas abajo del mismo canal — por eso se cruza aparte de la OIT.
  set('Afluente', d.afluente||'', valores);
  set('Afluente_Toxico', d.afluente_toxico ? 'SI':'NO', valores);
  set('Afluente_Organico', d.afluente_organico ? 'SI':'NO', valores);
  const nombreCol = headers.indexOf('Nombre'), localidadCol = headers.indexOf('Localidad');
  const data = sh.getDataRange().getValues();
  const nombreNorm = quitarAcentosGs_(d.nombre), localidadNorm = quitarAcentosGs_(d.localidad);
  for(let i=1;i<data.length;i++){
    if(quitarAcentosGs_(data[i][nombreCol])===nombreNorm && quitarAcentosGs_(data[i][localidadCol])===localidadNorm){
      const filaExistente = data[i].slice();
      headers.forEach((h,idx)=>{ if(valores[idx]!==undefined) filaExistente[idx]=valores[idx]; });
      sh.getRange(i+1,1,1,filaExistente.length).setValues([filaExistente]);
      return{success:true, message:'Empresa actualizada en el catastro'};
    }
  }
  headers.forEach((h,idx)=>{ if(valores[idx]!==undefined) fila[idx]=valores[idx]; });
  sh.appendRow(fila);
  return{success:true, message:'Empresa agregada al catastro'};
}

// Misma tabla de reglas OIT que usa OCM (Dashboard → Investigación) —
// si se actualiza una, hay que actualizar la otra a mano; no comparten
// archivo porque uno corre en el navegador y el otro en Apps Script.
const DIAGNOSTICO_AGENTE_MAP_ = [
  {palabras:['asma','bronqui','epoc','respirator','neumon','pulmon'], agentes:['Polvos y partículas','Humos metálicos','Gases y vapores','Asbesto','Hidrocarburos'], categoria:'Respiratorio'},
  {palabras:['dermat','eczema','alergia cutanea','alergia de piel','piel'], agentes:['Ácidos y álcalis','Solventes orgánicos','Plaguicidas / agroquímicos'], categoria:'Dermatológico'},
  {palabras:['neuropat','temblor','convuls','encefalopat','neurolog'], agentes:['Solventes orgánicos','Metales pesados','Plaguicidas / agroquímicos'], categoria:'Neurológico'},
  {palabras:['hipoacusia','sordera','audit'], agentes:['Ruido'], categoria:'Auditivo'},
  {palabras:['vibracion','sindrome vibratorio','raynaud'], agentes:['Vibración'], categoria:'Osteoarticular por vibración'},
  {palabras:['cancer','oncolog','tumor','leucemia'], agentes:['Asbesto','Hidrocarburos','Plaguicidas / agroquímicos','Radiación ionizante'], categoria:'Oncológico'},
  {palabras:['parasit','infeccios','zoonosis','bacteri','viral'], agentes:['Bacterias','Virus','Parásitos','Riesgo zoonótico (contacto con animales)'], categoria:'Infeccioso / zoonótico'},
  {palabras:['malformacion','congenit'], agentes:['Plaguicidas / agroquímicos'], categoria:'Congénito'}
];
function quitarAcentosGs_(s){ return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase(); }

// Cruce de Investigación para UN paciente puntual — usa Pacientes.Localidad
// + el/los diagnóstico(s) ya sincronizados del Libro de Discapacidad
// (Discapacidad_CUD) + Catastro_Empresas.
function getInvestigacionPaciente(dni){
  if(!dni) return{success:false, error:'Falta DNI'};
  const cudRows = leerPorDni_('Discapacidad_CUD', dni);
  const diagnosticos = [...new Set(cudRows.map(r=>r.Diagnostico).filter(Boolean))];
  // Localidad y afluente vienen de la página 1 del Libro de Discapacidad
  // (mismo origen que usaba el cruce original en OCM). Si por algún motivo
  // no está cargada ahí, se usa Pacientes.Localidad como respaldo.
  let localidad = (cudRows.map(r=>r.Localidad).filter(Boolean)[0]) || '';
  const afluente = (cudRows.map(r=>r.Afluente).filter(Boolean)[0]) || '';
  if(!localidad){
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pSheet = ss.getSheetByName('Pacientes');
    if(pSheet){
      const pData = pSheet.getDataRange().getValues(), pHeaders = pData[0];
      const dniCol = pHeaders.indexOf('DNI'), locCol = pHeaders.indexOf('Localidad');
      const fila = pData.find(row=>String(row[dniCol])===String(dni));
      if(fila) localidad = fila[locCol]||'';
    }
  }
  if((!localidad && !afluente) || !diagnosticos.length){
    return{success:true, localidad, afluente, hallazgos:[], hallazgos_hidricos:[], mensaje: (!localidad && !afluente) ? 'El paciente no tiene localidad ni afluente cargados.' : 'El paciente todavía no tiene diagnóstico sincronizado desde el Libro de Discapacidad.'};
  }
  const empresas = sheetToObjects('Catastro_Empresas').map(e=>{
    let aq=[],af=[],ab=[];
    try{ aq = JSON.parse(e.Agentes_Quimicos_JSON||'[]'); }catch(_){}
    try{ af = JSON.parse(e.Agentes_Fisicos_JSON||'[]'); }catch(_){}
    try{ ab = JSON.parse(e.Agentes_Biologicos_JSON||'[]'); }catch(_){}
    return {
      nombre:e.Nombre, localidad:e.Localidad, agentes_quimicos:aq, agentes_fisicos:af, agentes_biologicos:ab,
      afluente:e.Afluente||'', afluente_toxico:String(e.Afluente_Toxico).toUpperCase()==='SI', afluente_organico:String(e.Afluente_Organico).toUpperCase()==='SI'
    };
  });

  // 1) Vía OIT clásica — misma localidad + agente asociado al diagnóstico.
  const localidadNorm = quitarAcentosGs_(localidad);
  const hallazgos = [];
  if(localidad){
    diagnosticos.forEach(diag=>{
      const diagNorm = quitarAcentosGs_(diag);
      DIAGNOSTICO_AGENTE_MAP_.forEach(regla=>{
        if(!regla.palabras.some(pal=>diagNorm.includes(pal))) return;
        empresas.forEach(e=>{
          if(quitarAcentosGs_(e.localidad)!==localidadNorm) return;
          const agentesEmpresa = [].concat(e.agentes_quimicos, e.agentes_fisicos, e.agentes_biologicos);
          const coincidentes = agentesEmpresa.filter(a=>regla.agentes.indexOf(a)>-1);
          if(coincidentes.length) hallazgos.push({categoria:regla.categoria, diagnostico:diag, empresa:e.nombre, agentes:coincidentes});
        });
      });
    });
  }

  // 2) Vía hídrica — mismo canal de riego/afluente, SIN importar la
  // localidad: los efluentes viajan por el canal y pueden llegar a otra
  // zona aguas abajo. No depende de la tabla de diagnósticos-OIT: alcanza
  // con compartir el afluente para que sea una alerta a revisar.
  const hallazgos_hidricos = [];
  if(afluente){
    const afluenteNorm = quitarAcentosGs_(afluente);
    empresas.forEach(e=>{
      if(!e.afluente || quitarAcentosGs_(e.afluente)!==afluenteNorm) return;
      const categorias = [];
      if(e.afluente_toxico) categorias.push('Tóxico (vía hídrica/riego)');
      if(e.afluente_organico) categorias.push('Orgánico (vía hídrica/riego)');
      if(!categorias.length) categorias.push('Vía hídrica/riego (sin clasificar)');
      hallazgos_hidricos.push({categoria:categorias.join(' + '), empresa:e.nombre, afluente:e.afluente});
    });
  }

  return{success:true, localidad, afluente, hallazgos, hallazgos_hidricos};
}
// Compuerta de permisos — la usa doGet (get_investigacion_paciente).
function getInvestigacionPacienteGateado_(dni, profesionalEmail){
  if(!tieneModuloAutorizado_(profesionalEmail, 'investigacion')) return{success:false, error:'No autorizado para el módulo de Investigación'};
  return getInvestigacionPaciente(dni);
}

// ══════════════════════════════════════════════════════════════════
// PARCHE — Sincronización del Libro de Discapacidad (CUD) a la nube
// Antes el Libro de Discapacidad (las 9 páginas, incluida la de CUD) se
// guardaba SOLO en localStorage del navegador de OCM — nunca llegaba a la
// hoja Discapacidad_CUD, así que la vista consolidada que ya usa KSM
// quedaba estructuralmente lista pero vacía en la práctica. Esto agrega
// el guardado real, con upsert por DNI (una fila por paciente — se
// actualiza en vez de duplicar cada vez que el médico toca "Guardar").
// ══════════════════════════════════════════════════════════════════
function guardarLibroDiscapacidadCUD(d){
  if(!d || !d.dni) return{success:false, error:'Falta DNI'};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Discapacidad_CUD');
  if(!sh){
    sh = ss.insertSheet('Discapacidad_CUD');
    sh.appendRow(['Timestamp','DNI','Nombre','Telefono','Fecha_Nac','Diagnostico','Codigos_CIF','Tipo',
                  'Subtipo','Modalidad','Orientacion','Fecha_Emision','Fecha_Vencimiento','Junta',
                  'Firmantes_JSON','Efectores','Tipo_Pension','Trazabilidad']);
  }
  let headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  // 'Origen' distingue este registro (viene del Libro de Discapacidad de OCM)
  // del que arma el Lector CUD. 'Localidad' y 'Afluente' vienen de la página 1
  // del Libro — son la base del cruce de Investigación (OIT + vía hídrica).
  const faltantes = ['Origen','Localidad','Afluente'].filter(h=>headers.indexOf(h)===-1);
  if(faltantes.length){
    faltantes.forEach(h=> sh.getRange(1, sh.getLastColumn()+1).setValue(h));
    headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  }
  const data = sh.getDataRange().getValues();
  const dniCol = headers.indexOf('DNI'), origenCol = headers.indexOf('Origen');
  const valores = {
    Timestamp: new Date().toISOString(), DNI: d.dni, Nombre: d.nombre||'',
    Diagnostico: d.diagnostico||'', Codigos_CIF: d.codigos_cif||'', Tipo: d.tipo||'',
    Subtipo: d.subtipo||'', Fecha_Emision: d.emision||'', Fecha_Vencimiento: d.vencimiento||'',
    Localidad: d.localidad||'', Afluente: d.afluente||'',
    Origen: 'libro_discapacidad_ocm'
  };
  for(let i=1;i<data.length;i++){
    if(String(data[i][dniCol])===String(d.dni) && String(data[i][origenCol])==='libro_discapacidad_ocm'){
      const fila = data[i].slice();
      headers.forEach((h,idx)=>{ if(valores[h]!==undefined) fila[idx]=valores[h]; });
      sh.getRange(i+1,1,1,fila.length).setValues([fila]);
      return{success:true, message:'Libro de Discapacidad actualizado en la nube'};
    }
  }
  const filaNueva = headers.map(h=> valores[h]!==undefined ? valores[h] : '');
  sh.appendRow(filaNueva);
  return{success:true, message:'Libro de Discapacidad sincronizado a la nube por primera vez'};
}
// Compuerta de permisos — la usa doGet (get_discapacidad_paciente). No
// reutiliza get_vista_consolidada tal cual porque esa acción la sigue
// usando OCM sin mail de profesional — esta es la puerta gateada para KSM.
function getDiscapacidadPacienteGateado_(dni, profesionalEmail){
  if(!tieneModuloAutorizado_(profesionalEmail, 'cud')) return{success:false, error:'No autorizado para el módulo de Discapacidad/CUD'};
  if(!dni) return{success:false, error:'Falta DNI'};
  return{success:true, discapacidad: construirVistaConsolidada(dni).discapacidad};
}

// ══════════════════════════════════════════════════════════════════
// PARCHE — Sincronización de Plantillas (para compartir SOLO estado y
// detalle clínico con especialistas autorizados — a propósito, sin datos
// de pago). Igual que Catastro y el Libro de Discapacidad, Plantillas
// vivía nada más que en IndexedDB local del navegador de OCM.
// ══════════════════════════════════════════════════════════════════
function sheetPlantillasSync_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Plantillas_Sync');
  if(!sh){
    sh = ss.insertSheet('Plantillas_Sync');
    sh.appendRow(['Timestamp','DNI','Paciente','NSerie','Estado','Detalle','Controles_JSON','Mail']);
    sh.getRange(1,1,1,8).setBackground('#334155').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
  } else if(sh.getLastColumn()<8){
    // Hoja creada antes del registro de controles: se le agregan las dos columnas nuevas
    sh.getRange(1,7,1,2).setValues([['Controles_JSON','Mail']]).setBackground('#334155').setFontColor('white').setFontWeight('bold');
  }
  return sh;
}
// Alta/actualización de uno o varios pedidos (upsert por N° de pedido) leyendo la hoja UNA sola vez.
// Cada item: {dni, paciente, nSerie, estado, detalle, controles(objeto con fechas dd/mm/aaaa y *_hecho), mail}
// controles/mail solo se pisan si vienen en el pedido.
function upsertPlantillas_(lista){
  const sh = sheetPlantillasSync_();
  const lock = LockService.getScriptLock();
  try{
    lock.waitLock(15000);
    const data = sh.getDataRange().getValues();
    const idx = {}; // nSerie -> posición en filas
    const filas = [];
    for(let i=1;i<data.length;i++){
      const r = data[i].slice(0,8); while(r.length<8) r.push('');
      filas.push(r);
      if(r[3]) idx[String(r[3])] = filas.length-1;
    }
    let nuevos=0, actualizados=0;
    const ahora = new Date().toISOString();
    lista.forEach(function(d){
      if(!d || !d.dni || !d.nSerie) return;
      const k = String(d.nSerie);
      const ctl = (d.controles && typeof d.controles==='object') ? JSON.stringify(d.controles) : (typeof d.controles==='string' ? d.controles : null);
      if(idx[k]!==undefined){
        const r = filas[idx[k]];
        r[0]=ahora; r[1]=d.dni; r[2]=d.paciente||r[2];
        if(d.estado!==undefined) r[4]=d.estado||''; if(d.detalle!==undefined) r[5]=d.detalle||'';
        if(ctl!==null) r[6]=ctl; if(d.mail) r[7]=d.mail;
        actualizados++;
      }else{
        filas.push([ahora, d.dni, d.paciente||'', d.nSerie, d.estado||'', d.detalle||'', ctl||'', d.mail||'']);
        idx[k]=filas.length-1; nuevos++;
      }
    });
    if(filas.length){
      sh.getRange(2,1,filas.length,8).setNumberFormat('@').setValues(filas.map(function(r){ return r.map(function(v){ return v===null||v===undefined?'':String(v); }); }));
    }
    return {success:true, nuevos:nuevos, actualizados:actualizados};
  }catch(e){
    return {success:false, error:'No se pudo guardar plantillas: '+e};
  }finally{
    try{ lock.releaseLock(); }catch(_){}
  }
}
// La llama OCM al crear el pedido y al cambiar de estado / marcar un control.
function sincronizarPlantilla(d){
  if(!d || !d.dni || !d.nSerie) return{success:false, error:'Falta DNI o N° de pedido'};
  return upsertPlantillas_([d]);
}
function sincronizarPlantillasLote(d){
  if(!d || !Array.isArray(d.plantillas)) return {success:false, error:'Falta la lista de plantillas'};
  return upsertPlantillas_(d.plantillas);
}
function getPlantillasPaciente(dni){
  if(!dni) return{success:true, plantillas:[]};
  const sh = sheetPlantillasSync_();
  const data = sh.getDataRange().getValues();
  const plantillas = [];
  for(let i=1;i<data.length;i++){
    if(String(data[i][1])===String(dni)) plantillas.push({nSerie:data[i][3], estado:data[i][4], detalle:data[i][5], fecha:data[i][0]});
  }
  return{success:true, plantillas};
}
// Devuelve TODOS los pedidos de plantillas registrados en la nube (no solo
// los de un paciente) — lo usa OCM al arrancar para traer al navegador
// actual los pedidos que se cargaron desde otra computadora/navegador, que
// antes solo vivían en el IndexedDB de la compu donde se habían cargado.
// Protegida por token igual que el resto de las acciones "masivas" (ver
// ACCIONES_GET_PROTEGIDAS_), porque expone datos de todos los pacientes.
function getTodasPlantillasSync(){
  const sh = sheetPlantillasSync_();
  const data = sh.getDataRange().getValues();
  const plantillas = [];
  for(let i=1;i<data.length;i++){
    if(!data[i][3]) continue; // sin N° de pedido no hay con qué emparejar
    let ctl = null; try{ ctl = data[i][6] ? JSON.parse(data[i][6]) : null; }catch(_){ ctl = null; }
    plantillas.push({fecha:data[i][0], dni:data[i][1], paciente:data[i][2], nSerie:data[i][3], estado:data[i][4], detalle:data[i][5], controles:ctl, mail:data[i][7]||''});
  }
  return{success:true, plantillas};
}
// ══════════════════════════════════════════════════════════════════
// QR de pedido de plantillas (oct 2026) — QR_Plantillas.html
// El paciente manda desde el celular: (1) foto de los talones de pie,
// (2) foto de la plantilla vieja con una regla (y el largo en cm que él lee),
// (3) foto de la huella con talco. Las fotos NO se guardan en Drive ni en el
// sistema: salen como adjuntos al mail del médico. En OCM aparece sola una
// fila "Solicitada" (vía Plantillas_Sync). El sistema NO mide ni interpreta
// nada: el largo en cm es lo que declara el paciente y las fotos las
// revisa el médico. Segundo paso (mismo link, con DNI): el informe de las
// radiografías (diferencia de longitud de miembros inferiores y ángulos de
// Costa-Bartani), que también sale por mail y se anota en la misma fila.
// ══════════════════════════════════════════════════════════════════
const QRPL_MAX_B64_ = 3500000;               // ~2,6 MB por archivo (la página ya comprime)
const QRPL_TIPOS_IMG_ = ['image/jpeg','image/png','image/webp'];
const QRPL_TIPOS_RX_  = ['image/jpeg','image/png','image/webp','application/pdf'];
// Bocas de expendio (donde el paciente retira la plantilla). Las claves son las
// mismas que CENTROS en OCM, así OCM completa el centro solo al importar el pedido.
const QRPL_BOCAS_ = {
  sicoli:'Hospital Dr. Domingo Sícoli', espanol:'Hospital Español de Mendoza', palmares:'Centro Médico Palmares',
  sandona:'Clínica San Donà', altomaipu:'Centro Médico Alto Maipú', rehab:'Rehabilitarte Instituto Médico',
  norte:'Consultorios Privados del Norte', santamaria:'Clínica Santa María'
};

function qrplMailMedico_(){
  var m = '';
  try{ m = (PropertiesService.getScriptProperties().getProperty('PLANTILLAS_QR_MAIL')||'').trim(); }catch(e){}
  return m || 'lmuratori@gmail.com';
}
function qrplMailValido_(m){ return /^[^@\s]+@[^@\s.]+(\.[^@\s.]+)+$/.test(String(m||'').trim()); }
function qrplBlob_(dataUrl, nombre, tiposOk){
  const s = String(dataUrl||'');
  const i = s.indexOf(',');
  if(i<0 || i>120) throw new Error('El archivo "'+nombre+'" no se pudo leer. Volvé a sacarlo.');
  const m = /^data:([a-z0-9.+\/-]+);base64$/i.exec(s.slice(0,i));
  if(!m) throw new Error('El archivo "'+nombre+'" no tiene un formato válido.');
  const mime = m[1].toLowerCase();
  if(tiposOk.indexOf(mime)<0) throw new Error('El archivo "'+nombre+'" tiene un formato no admitido ('+mime+').');
  const b64 = s.slice(i+1);
  if(b64.length > QRPL_MAX_B64_) throw new Error('El archivo "'+nombre+'" es demasiado pesado. Sacá la foto de nuevo.');
  const ext = {'image/jpeg':'jpg','image/png':'png','image/webp':'webp','application/pdf':'pdf'}[mime];
  return Utilities.newBlob(Utilities.base64Decode(b64), mime, nombre+'.'+ext);
}
// Tope por DNI y por hora — evita que alguien llene el mail/hoja con basura.
function qrplLimite_(clave, max){
  try{
    const c = CacheService.getScriptCache();
    const n = parseInt(c.get(clave)||'0',10);
    if(n>=max) return false;
    c.put(clave, String(n+1), 3600);
  }catch(_){}
  return true;
}
function qrplLimpioDni_(v){ return String(v||'').replace(/[^0-9]/g,''); }
function qrplTexto_(v, max){ return String(v==null?'':v).replace(/[\u0000-\u001f]+/g,' ').trim().slice(0,max); }

// d = {apellido, nombre, dni, telefono, email, largoCm, observ, origen, web(honeypot),
//      fotoTalones, fotoPlantilla, fotoHuella}  (data:image/...;base64,...)
function registrarPlantillasQR(d){
  if(!d) return {success:false, error:'Faltan datos'};
  if(d.web) return {success:true, nSerie:'', omitido:true};   // honeypot: un bot llenó el campo oculto
  const apellido = qrplTexto_(d.apellido,60), nombre = qrplTexto_(d.nombre,60);
  const dni = qrplLimpioDni_(d.dni);
  const tel = qrplTexto_(d.telefono,30);
  const mail = String(d.email||'').trim().toLowerCase();
  const largo = parseFloat(String(d.largoCm||'').replace(',','.'));
  if(!apellido || !nombre) return {success:false, error:'Falta tu apellido y nombre'};
  if(dni.length<6 || dni.length>9) return {success:false, error:'El DNI no parece válido (solo números, sin puntos)'};
  if(!mail && !tel) return {success:false, error:'Dejá un mail o un teléfono de contacto'};
  if(mail && !qrplMailValido_(mail)) return {success:false, error:'El mail no es válido — revisá el punto (gmail.com)'};
  if(!(largo>=10 && largo<=40)) return {success:false, error:'Escribí el largo de tu plantilla en centímetros, tal como lo lees en la regla (entre 10 y 40)'};
  if(!d.fotoTalones)   return {success:false, error:'Falta la foto de los talones de pie'};
  if(!d.fotoPlantilla) return {success:false, error:'Falta la foto de la plantilla con la regla'};
  if(!d.fotoHuella)    return {success:false, error:'Falta la foto de la huella del pie'};
  const bocaKey = String(d.centro||'').trim().toLowerCase();
  if(!QRPL_BOCAS_[bocaKey]) return {success:false, error:'Elegí dónde vas a retirar tu plantilla'};
  if(!d.firmaPng) return {success:false, error:'Falta tu firma para autorizar el envío de las fotos'};
  if(!qrplLimite_('qrpl_'+dni, 3)) return {success:false, error:'Ya recibimos tus fotos hace instantes. Si necesitás corregir algo, escribinos por los canales del consultorio.'};
  let cuota = null; try{ cuota = MailApp.getRemainingDailyQuota(); }catch(_){}
  if(cuota!==null && cuota<=1) return {success:false, error:'El sistema de mails está al límite por hoy. Probá mañana o consultá en el consultorio.'};

  let blobs;
  try{
    blobs = [
      qrplBlob_(d.fotoTalones,   '1_talones_de_pie',     QRPL_TIPOS_IMG_),
      qrplBlob_(d.fotoPlantilla, '2_plantilla_con_regla', QRPL_TIPOS_IMG_),
      qrplBlob_(d.fotoHuella,    '3_huella_con_talco',    QRPL_TIPOS_IMG_)
    ];
    const firma = qrplBlob_(d.firmaPng, 'firma_consentimiento', ['image/png']);
    if(firma.getBytes().length > 400000) throw new Error('La firma es demasiado pesada. Borrala y firmá de nuevo.');
    blobs.push(firma);
  }catch(e){ return {success:false, error:String(e.message||e)}; }

  const fecha = Utilities.formatDate(new Date(), 'America/Argentina/Mendoza', 'yyyyMMdd');
  const nSerie = 'QR-' + fecha + '-' + dni;
  const pacienteTxt = (apellido+' '+nombre).toUpperCase();
  const observ = qrplTexto_(d.observ,400);
  const ahora = Utilities.formatDate(new Date(), 'America/Argentina/Mendoza', 'dd/MM/yyyy HH:mm');
  const bocaNombre = QRPL_BOCAS_[bocaKey];
  const textoConsent = qrplTexto_(d.textoConsent, 1500);
  const cuerpo =
    'Pedido de plantillas ingresado por QR.\n\n' +
    'BOCA DE EXPENDIO (donde retira): ' + bocaNombre + '\n' +
    'Paciente: ' + pacienteTxt + '\n' +
    'DNI: ' + dni + '\n' +
    'Teléfono: ' + (tel||'—') + '\n' +
    'Email: ' + (mail||'—') + '\n' +
    'Origen del QR: ' + qrplTexto_(d.origen||'sin especificar',80) + '\n\n' +
    'Largo de la plantilla actual, SEGÚN LO QUE LEYÓ EL PACIENTE en la regla: ' + largo + ' cm\n' +
    '(dato declarado por el paciente — el sistema no lo midió ni lo verificó; confirmalo con la foto 2).\n' +
    (observ ? ('\nObservaciones del paciente: ' + observ + '\n') : '') +
    '\nAdjuntos:\n  1) talones de pie (bipedestación)\n  2) plantilla con regla\n  3) huella con talco en el piso\n  4) firma del paciente (consentimiento)\n\n' +
    '── Consentimiento para compartir las imágenes ──\n' +
    'Firmado desde el celular del paciente el ' + ahora + ' (hora de Mendoza). Texto que el paciente leyó y aceptó:\n' +
    (textoConsent ? ('"' + textoConsent + '"') : '(la página no mandó el texto)') + '\n\n' +
    'El informe de las radiografías (diferencia de longitud de miembros inferiores y ángulos de Costa-Bartani) se pide en un segundo paso y llega en otro mail.\n\n' +
    'En OCM → Plantillas ya figura como "Solicitada" (N° ' + nSerie + ') después de actualizar.';
  try{
    MailApp.sendEmail(qrplMailMedico_(), 'Plantillas por QR — ' + pacienteTxt + ' — DNI ' + dni, cuerpo, {
      attachments: blobs, name:'S.I.M-M — Plantillas QR',
      replyTo: (mail || 'drluismuratori@gmail.com')
    });
  }catch(e){
    return {success:false, error:'No se pudo enviar el pedido: '+e};
  }

  // Alta en Plantillas_Sync (OCM la importa como "Solicitada"). Si ya hay un pedido QR del mismo día para el mismo DNI se actualiza, no se duplica.
  let advertencia = '';
  try{
    const prev = qrplBuscarFila_(dni, nSerie);
    const detalle = 'Solicitud por QR · boca de expendio: ' + bocaNombre + ' [centro:' + bocaKey + '] · largo declarado por el paciente ' + largo + ' cm · fotos (talones, plantilla con regla, huella) recibidas por mail · consentimiento firmado ' + ahora +
                    (prev && /informe RX: recibido/.test(String(prev.detalle||'')) ? ' · ' + String(prev.detalle).replace(/^.*?(informe RX: recibido[^·]*).*$/,'$1') : ' · informe RX: pendiente');
    const r = upsertPlantillas_([{dni:dni, paciente:pacienteTxt, nSerie:nSerie, estado:(prev&&prev.estado)||'solicitada', detalle:detalle, mail:mail}]);
    if(!r.success) advertencia = r.error;
  }catch(e){ advertencia = String(e); }

  // También en la Agenda de Mensajes (una sola vez por pedido del día)
  try{
    const c = CacheService.getScriptCache(); const kc = 'qrpl_cola_'+nSerie;
    if(!c.get(kc)){
      c.put(kc,'1',21600);
      qrplEncolarAgenda_({
        contacto: mail || tel, dni: dni, origen: 'QR plantillas' + (d.origen ? (' — ' + qrplTexto_(d.origen,60)) : ''),
        tipo: 'Pedido de plantillas (QR) — retiro: ' + bocaNombre,
        asunto: 'Pedido de plantillas por QR — ' + pacienteTxt + ' — DNI ' + dni,
        cuerpo:
          'Pedido de plantillas ingresado por QR.\n\n' +
          'Paciente: ' + pacienteTxt + '\nDNI: ' + dni + '\nTeléfono: ' + (tel||'—') + '\nEmail: ' + (mail||'—') + '\n' +
          'Boca de expendio (retiro): ' + bocaNombre + '\n' +
          'Largo de la plantilla actual, según lo que leyó el paciente en la regla: ' + largo + ' cm (dato declarado, no medido por el sistema)\n' +
          (observ ? ('Observaciones: ' + observ + '\n') : '') +
          '\nLas 3 fotos (talones de pie, plantilla con regla, huella con talco) y la firma del consentimiento llegaron como adjuntos al mail enviado a ' + qrplMailMedico_() + ' (asunto: "Plantillas por QR — ' + pacienteTxt + ' — DNI ' + dni + '"). Consentimiento firmado ' + ahora + '.\n' +
          'En OCM → Plantillas figura como "Solicitada" (N° ' + nSerie + ').',
        borrador:
          'Hola ' + nombre + ',\n\nRecibimos tus fotos para el pedido de plantillas (retiro en: ' + bocaNombre + '). El Dr. Muratori las va a revisar y te vamos a avisar los pasos siguientes.\n\n' +
          'Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943'
      });
    }
  }catch(_){}

  // Confirmación al paciente (si dejó mail). Texto sin datos clínicos.
  if(mail){
    try{
      const link = baseSimm_() + 'QR_Plantillas.html?paso=rx&dni=' + encodeURIComponent(dni);
      MailApp.sendEmail(mail, 'Recibimos tu pedido de plantillas — Dr. Muratori',
        'Hola ' + nombre + ',\n\nRecibimos tus fotos para el pedido de plantillas. El Dr. Muratori las va a revisar y te avisa los pasos siguientes.\n\n' +
        'Si tenés el informe de las radiografías (longitud de miembros inferiores y ángulos de Costa-Bartani), podés mandarlo desde este link:\n' + link + '\n\n' +
        'Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943',
        {replyTo:'drluismuratori@gmail.com', name:'Dr. Luis Alberto Muratori'});
    }catch(_){}
  }
  return {success:true, nSerie:nSerie, advertencia:advertencia};
}

// Deja el pedido también en la Agenda de Mensajes ("Mensaje Mail") como
// pendiente de revisión, igual que las consultas por QR. Las fotos NO se
// guardan acá: viajan solo como adjuntos del mail (ver qrplMailMedico_()).
// Nunca bloquea el envío: si algo falla, el mail y la fila de plantillas ya salieron.
function qrplEncolarAgenda_(o){
  try{
    const sheet = hojaColaMensajes_();
    const colOrigen = asegurarColumnaOrigenColaMensajes_(sheet);
    const colTipo = asegurarColumnaColaMensajes_(sheet, 'Tipo_Solicitud');
    const colDni = asegurarColumnaColaMensajes_(sheet, 'DNI_Detectado');
    const ancho = Math.max(13, colOrigen, colTipo, colDni);
    const fila = new Array(ancho).fill('');
    fila[0] = Utilities.getUuid();
    fila[1] = new Date().toISOString();
    fila[2] = o.contacto || '';
    fila[3] = o.asunto;
    fila[4] = o.cuerpo;
    fila[5] = 'privado';
    fila[6] = 1;
    fila[7] = false;            // requierePago
    fila[8] = 'no_aplica';
    fila[9] = o.borrador || '';
    fila[10] = 'pendiente_revision';
    fila[colOrigen-1] = o.origen || 'QR plantillas';
    fila[colTipo-1] = o.tipo;
    fila[colDni-1] = o.dni || '';
    sheet.appendRow(fila);
  }catch(e){ Logger.log('qrplEncolarAgenda_: ' + e); }
}

// Busca el pedido QR más reciente de un DNI (o uno exacto por nSerie).
function qrplBuscarFila_(dni, nSerieExacto){
  const sh = sheetPlantillasSync_();
  const data = sh.getDataRange().getValues();
  let hallado = null;
  for(let i=1;i<data.length;i++){
    if(String(data[i][1])!==String(dni)) continue;
    const ns = String(data[i][3]||'');
    if(nSerieExacto){ if(ns===nSerieExacto) return {nSerie:ns, estado:data[i][4], detalle:data[i][5], paciente:data[i][2], mail:data[i][7]||''}; continue; }
    if(/^QR-/.test(ns)) hallado = {nSerie:ns, estado:data[i][4], detalle:data[i][5], paciente:data[i][2], mail:data[i][7]||''};
  }
  return hallado;
}

// d = {dni, nota, archivos:[dataURL imagen/PDF, ...hasta 3], web}
function registrarInformeRxQR(d){
  if(!d) return {success:false, error:'Faltan datos'};
  if(d.web) return {success:true, omitido:true};
  const dni = qrplLimpioDni_(d.dni);
  if(dni.length<6 || dni.length>9) return {success:false, error:'El DNI no parece válido (solo números, sin puntos)'};
  const arch = Array.isArray(d.archivos) ? d.archivos.filter(Boolean).slice(0,3) : [];
  if(!arch.length) return {success:false, error:'Adjuntá la foto o el PDF del informe de las radiografías'};
  if(d.acepta!==true) return {success:false, error:'Tildá la casilla que autoriza el envío del informe'};
  if(!qrplLimite_('qrplrx_'+dni, 6)) return {success:false, error:'Ya recibimos varios archivos con este DNI hace instantes.'};
  let cuota = null; try{ cuota = MailApp.getRemainingDailyQuota(); }catch(_){}
  if(cuota!==null && cuota<=1) return {success:false, error:'El sistema de mails está al límite por hoy. Probá mañana.'};
  let blobs;
  try{ blobs = arch.map(function(a,i){ return qrplBlob_(a, 'informe_RX_'+(i+1), QRPL_TIPOS_RX_); }); }
  catch(e){ return {success:false, error:String(e.message||e)}; }

  let fila = null;
  try{ fila = qrplBuscarFila_(dni, ''); }catch(_){}
  const nota = qrplTexto_(d.nota,500);
  const hoy = Utilities.formatDate(new Date(), 'America/Argentina/Mendoza', 'dd/MM/yyyy');
  const cuerpo =
    'Informe de radiografías enviado por el paciente (QR de plantillas).\n\n' +
    'DNI: ' + dni + '\n' +
    'Paciente: ' + (fila ? fila.paciente : '(no se encontró un pedido de plantillas por QR con ese DNI — revisá que sea el mismo)') + '\n' +
    (fila ? ('Pedido N°: ' + fila.nSerie + '\n') : '') +
    'Archivos: ' + blobs.length + '\n' +
    (nota ? ('\nNota del paciente: ' + nota + '\n') : '') +
    '\nSe espera: diferencia de longitud de miembros inferiores y ángulos de Costa-Bartani. El sistema no interpreta el informe: lo ves vos.';
  try{
    MailApp.sendEmail(qrplMailMedico_(), 'Plantillas por QR — informe RX — DNI ' + dni + (fila ? (' — ' + fila.paciente) : ''), cuerpo, {
      attachments: blobs, name:'S.I.M-M — Plantillas QR', replyTo: (fila && fila.mail) || 'drluismuratori@gmail.com'
    });
  }catch(e){
    return {success:false, error:'No se pudo enviar el informe: '+e};
  }
  if(fila){
    try{
      const det = String(fila.detalle||'').replace(/informe RX: pendiente/, 'informe RX: recibido ' + hoy + ' (' + blobs.length + ' arch., por mail)');
      upsertPlantillas_([{dni:dni, paciente:fila.paciente, nSerie:fila.nSerie, detalle:(det.indexOf('informe RX: recibido')>-1 ? det : det + ' · informe RX: recibido ' + hoy + ' (' + blobs.length + ' arch., por mail)')}]);
    }catch(_){}
  }
  qrplEncolarAgenda_({
    contacto: (fila && fila.mail) || '', dni: dni, origen: 'QR plantillas',
    tipo: 'Informe de radiografías (QR plantillas)',
    asunto: 'Informe RX por QR plantillas — DNI ' + dni + (fila ? (' — ' + fila.paciente) : ''),
    cuerpo: 'El paciente mandó ' + blobs.length + ' archivo(s) con el informe de radiografías (longitud de miembros inferiores y ángulos de Costa-Bartani).\n' +
            'Paciente: ' + (fila ? fila.paciente : '(sin pedido de plantillas por QR con ese DNI)') + '\nDNI: ' + dni + '\n' +
            (nota ? ('Nota del paciente: ' + nota + '\n') : '') +
            '\nLos archivos llegaron como adjuntos al mail enviado a ' + qrplMailMedico_() + ' (asunto: "Plantillas por QR — informe RX — DNI ' + dni + '"). El sistema no interpreta el informe.',
    borrador: fila ? ('Hola,\n\nRecibimos el informe de tus radiografías. El Dr. Muratori lo va a revisar junto con tus fotos y te avisamos los pasos siguientes.\n\nDr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943') : ''
  });
  return {success:true, vinculado:!!fila};
}

// Compuerta de permisos — la usa doGet (get_plantillas_paciente).
function getPlantillasPacienteGateado_(dni, profesionalEmail){
  if(!tieneModuloAutorizado_(profesionalEmail, 'plantillas')) return{success:false, error:'No autorizado para el módulo de Plantillas'};
  return getPlantillasPaciente(dni);
}

// ══════════════════════════════════════════════════════════════════
// Listas personalizadas — usado por CAM (Herramientas Clínicas):
// diccionario propio de síndromes/ítems que el médico va agregando,
// guardado en la nube (misma base que el resto de S.I.M-M) y disponible
// para siempre en cualquier dispositivo. Genérico por "tipo" para poder
// reutilizarse con otras listas futuras (instituciones, procedimientos,
// códigos, etc.) sin tener que tocar el backend de nuevo.
// ══════════════════════════════════════════════════════════════════
function hojaListasPersonalizadas_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Listas_Personalizadas');
  if(!sheet){
    sheet = ss.insertSheet('Listas_Personalizadas');
    sheet.appendRow(['tipo','valor','fecha']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
function getListaQuirurgico(tipo){
  const sheet = hojaListasPersonalizadas_();
  const rows = sheet.getDataRange().getValues();
  const data = [];
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][0])===String(tipo)) data.push(rows[i][1]);
  }
  return {success:true, data};
}
function guardarItemQuirurgico(d){
  if(!d || !d.tipo || !d.valor) return {success:false, error:'Falta tipo o valor'};
  const sheet = hojaListasPersonalizadas_();
  sheet.appendRow([d.tipo, d.valor, new Date().toISOString()]);
  return {success:true};
}

// ══════════════════════════════════════════════════════════════════
// MINI-RIMP — Interconsultas pre-quirúrgicas por especialidad
//
// Etapa actual (desarrollo): estos son los campos consensuados con cada
// especialista sobre el Mini-RIMP de su especialidad (ver las páginas de
// validación enviadas a cada uno). Acá quedan cargados como columnas
// reales de una hoja de cálculo — una hoja por especialidad — para que
// cuando el circuito completo esté armado en OCM (el médico pide la
// interconsulta → se dispara el envío al especialista → el especialista
// escanea la receta, carga sus datos profesionales y completa el
// Mini-RIMP → la conclusión se refleja en la historia clínica) los datos
// tengan dónde guardarse de verdad, sin inventar nada de la clínica.
//
// MINIRIMP_CAMPOS es la ÚNICA fuente de verdad de qué campo tiene cada
// especialidad — agregar o sacar un campo acá alcanza para que
// setupHojasMiniRIMP() actualice la hoja correspondiente (agrega
// columnas nuevas sin tocar lo ya cargado; nunca borra columnas viejas).
// ══════════════════════════════════════════════════════════════════
const MINIRIMP_CAMPOS = {
  // Campos de Anestesiología ampliados (sep 2026) tras comparar contra
  // "Esquemas para la evaluación preoperatoria" (Jegier, anestesia.org.ar):
  // se agregaron antecedentes anestésicos propios/familiares más allá de
  // hipertermia maligna, tabaquismo, consumo de alcohol, medicación
  // habitual, estado dentario (piezas rotas/flojas, no solo prótesis) y
  // distancia tiromentoniana.
  // Ampliado nuevamente (sep 2026) comparando contra la planilla real de
  // "Evaluación Preoperatoria" que completa el anestesiólogo en el
  // pre-quirúrgico: se agregaron test de mordida, movilidad cervical,
  // antecedente de intubación dificultosa, los 5 factores de riesgo de
  // ventilación dificultosa (obesidad, barba, edad>55, SAOS/roncador,
  // edentado), consumo de sustancias ilícitas, requerimiento de cama UTI
  // y disponibilidad de hemoderivados — todos tomados literalmente de esa
  // planilla, sin inventar nada.
  anestesiologia: ['Apertura bucal','Distancia tiromentoniana','Uso de prótesis dental removible','Piezas dentarias rotas o flojas','Piercing corporal','Clasificación de Mallampati','Clasificación ASA','Test de mordida','Movilidad cervical','Antecedente de intubación dificultosa','Factor de riesgo de ventilación dificultosa: obesidad (IMC mayor a 26)','Factor de riesgo de ventilación dificultosa: barba','Factor de riesgo de ventilación dificultosa: edad mayor a 55 años','Factor de riesgo de ventilación dificultosa: SAOS / roncador','Factor de riesgo de ventilación dificultosa: edentado','Enfermedades de la infancia','Esquema de vacunación e inmunidad antitetánica','Antecedente de hipertermia maligna (personal o familiar)','Antecedentes anestésicos previos y complicaciones (personales o familiares)','Tabaquismo','Consumo de alcohol','Consumo de sustancias ilícitas','Medicación habitual actual','Convivencia con mascotas','Alergia al látex','Requerimiento de cama UTI','Disponibilidad de hemoderivados'],
  odontologia: ['Foco infeccioso oral activo','Piezas dentarias móviles o en mal estado','Profilaxis antibiótica en paciente con prótesis articular','Particularidades pediátricas','Conclusión del odontólogo'],
  // Cardiología: campos verificados contra el Consenso Argentino SAC 2016
  // de Evaluación de Riesgo Cardiovascular en Cirugía No Cardíaca (sigue
  // vigente, no hay versión posterior). Se agregó el manejo perioperatorio
  // de antiagregantes/anticoagulantes porque el consenso lo trata como
  // recomendación fuerte (Clase I/III según el fármaco) y no estaba
  // preguntado como campo propio.
  cardiologia: ['Capacidad funcional','Factores de riesgo de Lee (RCRI)','Criterios clínicos mayores (Paso 1, Consenso SAC 2016)','Criterios clínicos menores (Paso 1, Consenso SAC 2016)','Riesgo quirúrgico según tipo de procedimiento (Paso 2)','Riesgo clínico-quirúrgico combinado (Paso 3)','Modalidad de internación planificada (a definir en OCM)','Antecedentes cardiovasculares','Medicación cardiovascular vigente','Plan de manejo perioperatorio de antiagregantes/anticoagulantes (suspensión, puente)','Electrocardiograma reciente','Biomarcadores (BNP/NT-proBNP, troponina)'],
  neurologia: ['Diagnóstico de epilepsia o trastorno convulsivo','Esquema anticonvulsivante actual','Última crisis registrada','Disponibilidad de vía IV si el paciente queda en ayuno prolongado'],
  hematologia: ['Anticoagulante actual','Motivo de la anticoagulación','Necesidad de terapia puente (bridging)','Último control de coagulación (RIN u otro)'],
  infectologia_ortopedia: ['Screening nasal de Staphylococcus aureus (incl. SAMR)','Screening extranasal de S. aureus — conducto auditivo','Screening extranasal de S. aureus — axila','Screening extranasal de S. aureus — ingle','Estado de las láminas ungueales (manos y pies)','Folículos pilosos con signos de infección (foliculitis)','Rasurado preoperatorio','Profilaxis antibiótica (cefazolina)','Serologías infectocontagiosas','Consentimiento informado para VIH','Esquema de vacunación e inmunidad antitetánica'],
  pediatria: ['Peso y talla actuales (con percentilo)','Antecedentes perinatales relevantes','Infección respiratoria alta reciente (últimas 2-4 semanas)','Antecedente de apneas, ronquido o sospecha de apnea obstructiva del sueño','Esquema de vacunación','Alergias conocidas (alimentarias, medicamentosas)','Desarrollo madurativo acorde a la edad'],
  clinico: ['Antecedentes patológicos personales','Medicación habitual completa','Antecedentes quirúrgicos y anestésicos previos','Hábitos (tabaquismo, consumo de alcohol)','Control de signos vitales y examen físico general','Estudios de laboratorio pre-quirúrgico'],
  nutricion: ['Índice de masa corporal (IMC)','Pérdida de peso no intencional reciente','Albuminemia / proteínas totales (si disponible)','Restricciones o particularidades alimentarias','Plan nutricional pre y postoperatorio'],
  psiquiatria: ['Antecedentes psiquiátricos personales','Medicación psiquiátrica vigente','Nivel de ansiedad ante la cirugía','Red de apoyo y contención post-quirúrgica','Conclusión del especialista'],
  dermatologia: ['Lesiones cutáneas activas en la zona a operar','Lesiones cutáneas activas a distancia','Antecedente de cicatrización patológica','Alergias cutáneas a antisépticos o apósitos'],
  urologia: ['Antecedentes urológicos relevantes','Sedimento urinario y urocultivo','Función renal (creatinina, filtrado estimado)','Sonda vesical o dispositivos urológicos permanentes'],
};

// Nombres de hoja cortos (Sheets no distingue bien nombres muy largos con
// barras "/"), pero se mantienen legibles y estables — no cambiarlos una
// vez que haya datos cargados, porque romperían la referencia.
const MINIRIMP_NOMBRE_HOJA_ = {
  anestesiologia:'MiniRIMP_Anestesiologia', odontologia:'MiniRIMP_Odontologia', cardiologia:'MiniRIMP_Cardiologia',
  neurologia:'MiniRIMP_Neurologia', hematologia:'MiniRIMP_Hematologia', infectologia_ortopedia:'MiniRIMP_InfectOrtopedia',
  pediatria:'MiniRIMP_Pediatria', clinico:'MiniRIMP_Clinico', nutricion:'MiniRIMP_Nutricion',
  psiquiatria:'MiniRIMP_Psiquiatria', dermatologia:'MiniRIMP_Dermatologia', urologia:'MiniRIMP_Urologia'
};
function nombreHojaMiniRIMP_(especialidad){
  return MINIRIMP_NOMBRE_HOJA_[especialidad] || null;
}
// Columnas fijas, iguales en las 12 hojas, antes de las columnas propias
// de cada especialidad (que salen de MINIRIMP_CAMPOS).
const MINIRIMP_COLUMNAS_FIJAS_ = [
  'Timestamp','ID_Interconsulta','DNI_Paciente','Nombre_Paciente',
  'Nombre_Especialista','Matricula_Especialista','Email_Especialista',
  'Estado','Fecha_Respuesta','Aprobacion','Sugerencias_Mejora',
  'Resumen_Quirurgico'
];

// Crea (o actualiza) las 12 hojas de Mini-RIMP con sus columnas reales.
// Se corre UNA VEZ a mano desde el editor de Apps Script (Ejecutar >
// setupHojasMiniRIMP), igual que ya se hace con setupSheets(). Si una
// hoja ya existe, solo agrega las columnas que falten al final —
// nunca borra ni reordena columnas con datos ya cargados.
function setupHojasMiniRIMP(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const creadas = [];
  Object.keys(MINIRIMP_CAMPOS).forEach(especialidad=>{
    const nombreHoja = nombreHojaMiniRIMP_(especialidad);
    const headersDeseados = MINIRIMP_COLUMNAS_FIJAS_.concat(MINIRIMP_CAMPOS[especialidad]);
    let sh = ss.getSheetByName(nombreHoja);
    if(!sh){
      sh = ss.insertSheet(nombreHoja);
      sh.getRange(1,1,1,headersDeseados.length).setValues([headersDeseados])
        .setBackground('#0f5c56').setFontColor('white').setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.setFrozenColumns(2);
      creadas.push(nombreHoja+' (nueva)');
    } else {
      const headersActuales = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
      let agregadas = 0;
      headersDeseados.forEach(h=>{
        if(headersActuales.indexOf(h)===-1){
          sh.getRange(1, sh.getLastColumn()+1).setValue(h).setBackground('#0f5c56').setFontColor('white').setFontWeight('bold');
          agregadas++;
        }
      });
      creadas.push(nombreHoja+' ('+agregadas+' columna/s nueva/s)');
    }
  });
  // BUGFIX (sep 2026): esta función se pensó para correrla a mano desde el
  // editor de Apps Script (por eso el alert), pero guardarRespuestaMiniRIMP()
  // también la llama sola cuando falta la hoja de una especialidad — y esa
  // otra llamada pasa por doPost, un pedido web sin ventana de Apps Script.
  // SpreadsheetApp.getUi() no existe en ese contexto y tira error; sin este
  // try/catch, la primera respuesta de un especialista para una especialidad
  // sin hoja todavía se perdía (el request fallaba entero, sin guardar nada,
  // y sin que el especialista ni el médico se enteraran de por qué).
  try{ SpreadsheetApp.getUi().alert('✅ Hojas Mini-RIMP listas:\n\n'+creadas.join('\n')); }
  catch(e){ Logger.log('Hojas Mini-RIMP listas (alert omitido, no hay UI): '+creadas.join(' | ')); }
  return {success:true, hojas: creadas};
}

// Guarda la respuesta de un especialista para un paciente puntual.
// d = { especialidad, dni, paciente, nombreEspecialista, matriculaEspecialista,
//       emailEspecialista, estado, aprobacion, sugerencias, campos:{...} }
// 'campos' trae SOLO los nombres que están en MINIRIMP_CAMPOS[especialidad];
// cualquier otro nombre se ignora (no se inventa columna sobre la marcha).
function guardarRespuestaMiniRIMP(d){
  if(!d || !d.especialidad) return {success:false, error:'Falta especialidad'};
  const nombreHoja = nombreHojaMiniRIMP_(d.especialidad);
  if(!nombreHoja) return {success:false, error:'Especialidad desconocida: '+d.especialidad};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nombreHoja);
  if(!sh){ setupHojasMiniRIMP(); sh = ss.getSheetByName(nombreHoja); }
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const fila = new Array(headers.length).fill('');
  const set=(nombreCol,valor)=>{ const i=headers.indexOf(nombreCol); if(i>-1) fila[i]=valor; };
  const id = d.id || Utilities.getUuid();
  set('Timestamp', new Date().toISOString());
  set('ID_Interconsulta', id);
  set('DNI_Paciente', d.dni||'');
  set('Nombre_Paciente', d.paciente||'');
  set('Nombre_Especialista', d.nombreEspecialista||'');
  set('Matricula_Especialista', d.matriculaEspecialista||'');
  set('Email_Especialista', d.emailEspecialista||'');
  set('Estado', d.estado||'Completado');
  set('Fecha_Respuesta', new Date().toISOString());
  set('Aprobacion', d.aprobacion||'');
  set('Sugerencias_Mejora', d.sugerencias||'');
  set('Resumen_Quirurgico', d.resumenQuirurgico||'');
  (MINIRIMP_CAMPOS[d.especialidad]||[]).forEach(campo=>{
    if(d.campos && Object.prototype.hasOwnProperty.call(d.campos, campo)) set(campo, d.campos[campo]);
  });
  sh.appendRow(fila);
  return {success:true, id: id};
}

// Lee lo cargado para un paciente (o todo, si no se pasa dni) de una
// especialidad puntual. Queda protegida por token (ACCIONES_GET_PROTEGIDAS_)
// porque expone datos de pacientes.
function getMiniRIMP(especialidad, dni){
  const nombreHoja = nombreHojaMiniRIMP_(especialidad);
  if(!nombreHoja) return {success:false, error:'Especialidad desconocida: '+especialidad};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(nombreHoja);
  if(!sh) return {success:true, registros:[]};
  const data = sh.getDataRange().getValues();
  if(data.length<2) return {success:true, registros:[]};
  const headers = data[0];
  const dniCol = headers.indexOf('DNI_Paciente');
  const registros = data.slice(1)
    .filter(row => !dni || String(row[dniCol])===String(dni))
    .map(row => { const obj={}; headers.forEach((h,i)=>{obj[h]=row[i];}); return obj; });
  return {success:true, registros};
}

// Solo el esquema de campos (sin datos de pacientes) — lo usa el
// frontend para dibujar el formulario correspondiente a la especialidad.
// No necesita token: no expone nada sensible, es la lista de preguntas.
function getCamposMiniRIMP(especialidad){
  if(!especialidad) return {success:true, especialidades: Object.keys(MINIRIMP_CAMPOS)};
  if(!MINIRIMP_CAMPOS[especialidad]) return {success:false, error:'Especialidad desconocida: '+especialidad};
  return {success:true, especialidad, campos: MINIRIMP_CAMPOS[especialidad]};
}


// ═════════════════════════════════════════════════════════════════════════
// KEP — Evaluador de Parámetros Kinésicos (oct 2026)
//
// Circuito de tres puntas, calcado del que ya existe para los médicos
// (RMYT: autoevaluación semanal → semáforo → panel del médico):
//
//   PACIENTE   →  KEP_Evaluador_Paciente.html (link por mail, sin token)
//                 acción pública kep_submit
//   KINESIÓLOGO → KSM_Kinesiologia.html: ve el semáforo y la tendencia,
//                 carga sus mediciones objetivas (rango articular, fuerza
//                 0-5, test funcional), manda el link al paciente y puede
//                 DERIVAR al médico (kep_medicion, kep_enviar_link,
//                 kep_derivar_medico, get_kep — todas con token)
//   MÉDICO     →  pestaña "Kinesiología" del Seguimiento de OCM
//                 (get_kep_dashboard) y respuesta al kinesiólogo
//                 (kep_responder_medico).
//
// Reglas del semáforo (documentadas para que se puedan revisar):
//   - Dolor en aumento: EVA en movimiento sube ≥ KEP_UMBRAL_EVA_ puntos
//     respecto de la carga anterior. 2 puntos (~30 %) es el cambio de dolor
//     que en la literatura se considera clínicamente importante (Farrar
//     2001, Pain 94:149-158; Salaffi 2004, Eur J Pain 8:283-291).
//   - Función empeora: el promedio de la PSFS (Patient-Specific Functional
//     Scale, 0-10) baja ≥ KEP_UMBRAL_PSFS_ puntos. Cambio "mediano" de la
//     PSFS ≈ 2,3 (Abbott & Schmitt 2014, JOSPT 44:560-564).
//   - Banderas rojas (criterio de NICE NG234 y del marco IFOMPT 2020 para
//     columna: dolor nocturno/progresivo, déficit neurológico, alteración
//     de esfínteres o anestesia en silla de montar): una sola bandera es un
//     indicador débil, importan las combinaciones. Por eso: ROJO si hay
//     alteración de esfínteres/silla de montar, pérdida de fuerza
//     progresiva, o dos o más banderas juntas; AMARILLO con una sola.
//     Fiebre y articulación caliente son criterio clínico habitual (no
//     tienen una guía verificada).
//   - Es un semáforo de ayuda: NO diagnostica ni reemplaza el criterio
//     médico, y las alertas siempre lo dicen.
// ═════════════════════════════════════════════════════════════════════════
const KEP_UMBRAL_EVA_ = 2;
const KEP_UMBRAL_PSFS_ = 2;
const KEP_DIAS_SIN_RESPUESTA_ = 14;
const KEP_MAX_ENVIOS_DIA_ = 20;
const KEP_BANDERAS_ = {
  noche_progresivo: 'Dolor intenso que no cede en reposo o que despierta de noche, y va en aumento',
  fuerza_progresiva: 'Pérdida de fuerza que va en aumento en un brazo o una pierna',
  neuro_aumento: 'Hormigueo o adormecimiento constante o que va en aumento',
  esfinteres_silla: 'Dificultad para controlar la orina o la materia fecal, o adormecimiento en la zona genital/entrepierna',
  fiebre: 'Fiebre o malestar general junto con el dolor',
  hinchazon_caliente: 'Hinchazón, calor y enrojecimiento nuevos en una articulación o en la pierna',
  caida_golpe: 'Caída o golpe nuevo desde la última evaluación'
};
const KEP_BANDERAS_GRAVES_ = ['esfinteres_silla','fuerza_progresiva'];
const KEP_CAMBIO_GLOBAL_ = ['Mucho mejor','Algo mejor','Igual','Algo peor','Mucho peor'];

function kepHoja_(nombre, headers){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nombre);
  if(!sh){
    sh = ss.insertSheet(nombre);
    sh.appendRow(headers);
    sh.getRange(1,1,1,headers.length).setBackground('#166534').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
  } else {
    headers.forEach(function(h){ asegurarColumnaColaMensajes_(sh, h); });
  }
  return sh;
}
const KEP_HDR_EVAL_ = ['Timestamp','DNI','Paciente','Region','Cumplimiento_Ejercicios','Cumplimiento_Sesiones','EVA_Reposo','EVA_Movimiento','Dolor_Nocturno','PSFS_Prom','PSFS_JSON','Cambio_Global','Tolerancia','Banderas','Alerta','Nivel','Dificultades','Datos_JSON'];
const KEP_HDR_MED_ = ['Timestamp','DNI','Paciente','Region','Kinesiologo','Kinesiologo_Email','ROM_JSON','Fuerza_JSON','Test_Funcional','Observaciones'];
const KEP_HDR_SEG_ = ['DNI','Paciente','Region','Kinesiologo','Kinesiologo_Email','Fecha_Inicio','Ultimo_Link','Activo'];
const KEP_HDR_DER_ = ['Timestamp','DNI','Paciente','Kinesiologo','Kinesiologo_Email','Motivo','Resumen','Estado','Respuesta_Medico','Fecha_Respuesta'];

function kepFilas_(sh){
  const d = sh.getDataRange().getValues();
  if(d.length < 2) return [];
  const h = d[0];
  const out = [];
  for(let i=1;i<d.length;i++){
    const o = {_row:i+1};
    for(let j=0;j<h.length;j++) o[h[j]] = d[i][j] instanceof Date ? d[i][j].toISOString() : d[i][j];
    out.push(o);
  }
  return out;
}
function kepNum_(v){ return (v===''||v===null||v===undefined||isNaN(v)) ? null : Number(v); }
function kepPaciente_(dni){
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pacientes');
  if(!sh) return null;
  const d = sh.getDataRange().getValues(), h = d[0]||[];
  const iD = h.indexOf('DNI'), iA = h.indexOf('Apellido'), iN = h.indexOf('Nombre'), iE = h.indexOf('Email'), iDx = h.indexOf('Diagnostico_Principal'), iOS = h.indexOf('Obra_Social'), iT = h.indexOf('Telefono');
  for(let i=1;i<d.length;i++){
    if(String(d[i][iD])===String(dni)){
      return { dni:String(dni), apellido:String(d[i][iA]||''), nombre:String(d[i][iN]||''), email:String(iE>-1?d[i][iE]:''),
        dx:String(iDx>-1?d[i][iDx]:''), os:String(iOS>-1?d[i][iOS]:''), tel:String(iT>-1?d[i][iT]:'') };
    }
  }
  // Paciente que solo se inscribió en kinesiología (QR del centro): vive en la historia clínica kinésica
  const kh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KIN_HC');
  if(kh){
    const kd = kh.getDataRange().getValues(), kk = kd[0]||[], ci = function(n){ return kk.indexOf(n); };
    for(let j=kd.length-1;j>=1;j--){
      if(String(kd[j][ci('DNI')])===String(dni)){
        return { dni:String(dni), apellido:String(kd[j][ci('Apellido')]||''), nombre:String(kd[j][ci('Nombre')]||''), email:String(kd[j][ci('Email')]||''),
          dx:String(kd[j][ci('Diagnostico')]||''), os:String(kd[j][ci('Obra_Social')]||''), tel:String(kd[j][ci('Telefono')]||'') };
      }
    }
  }
  return null;
}
function kepNombre_(p, dni){ return p ? ((p.apellido||'')+', '+(p.nombre||'')).replace(/^, |, $/g,'') : ('DNI '+dni); }

// % de discapacidad equivalente a partir de la PSFS (0-10): 10 = hace todo como antes.
function kepPsfsAPctDisc_(psfs){ return psfs==null ? null : Math.round((10 - psfs) * 100) / 10; }
// Percepción global → 0-100 (aproximación gruesa, solo para detectar discordancia entre lo que siente el paciente y lo medido)
function kepCambioAPct_(c){ const i = KEP_CAMBIO_GLOBAL_.indexOf(c); return i<0 ? null : [90,70,50,30,10][i]; }

function kepEvaluarAlertas_(cur, prev, prev2){
  const alertas = []; let nivel = 'verde';
  const subir = function(n){ if(n==='rojo') nivel='rojo'; else if(n==='amarillo' && nivel!=='rojo') nivel='amarillo'; };
  const banderas = cur.banderas || [];
  const graves = banderas.filter(function(b){ return KEP_BANDERAS_GRAVES_.indexOf(b)>-1; });
  if(graves.length || banderas.length >= 2){ alertas.push('BANDERAS_ROJAS'); subir('rojo'); }
  else if(banderas.length === 1){ alertas.push('BANDERA_A_REVISAR'); subir('amarillo'); }
  if(prev && cur.eva_mov != null && prev.eva_mov != null && (cur.eva_mov - prev.eva_mov) >= KEP_UMBRAL_EVA_){ alertas.push('DOLOR_EN_AUMENTO'); subir('amarillo'); }
  if(prev && cur.psfs != null && prev.psfs != null && (prev.psfs - cur.psfs) >= KEP_UMBRAL_PSFS_){ alertas.push('FUNCION_EMPEORA'); subir('amarillo'); }
  if(cur.cambio === 'Mucho peor' || (cur.cambio === 'Algo peor' && prev && (prev.cambio === 'Algo peor' || prev.cambio === 'Mucho peor'))){ alertas.push('PEOR_PERCIBIDO'); subir('amarillo'); }
  if(cur.cump != null && cur.cump <= 3){ alertas.push('BAJO_CUMPLIMIENTO'); subir('amarillo'); }
  if(cur.tolerancia === 'Mala' && prev && prev.tolerancia === 'Mala'){ alertas.push('NO_TOLERA_EJERCICIOS'); subir('amarillo'); }
  return { alertas: alertas, nivel: nivel };
}
function kepEnLista_(v){ return String(v||'').split(',').map(function(x){ return x.trim(); }).filter(Boolean); }
function kepFilaAReg_(f){
  return { fecha:String(f.Timestamp||''), region:String(f.Region||''), eva_reposo:kepNum_(f.EVA_Reposo), eva_mov:kepNum_(f.EVA_Movimiento),
    eva_movimiento:kepNum_(f.EVA_Movimiento), psfs:kepNum_(f.PSFS_Prom), cump:(function(){ const a=kepNum_(f.Cumplimiento_Ejercicios), b=kepNum_(f.Cumplimiento_Sesiones); const v=[a,b].filter(function(x){return x!==null;}); return v.length ? Math.round(v.reduce(function(x,y){return x+y;},0)/v.length*10)/10 : null; })(),
    cambio:String(f.Cambio_Global||''), tolerancia:String(f.Tolerancia||''), banderas:kepEnLista_(f.Banderas), alerta:String(f.Alerta||''), nivel:String(f.Nivel||''),
    escala_pct_discapacidad:kepPsfsAPctDisc_(kepNum_(f.PSFS_Prom)), recuperacion_pct:kepCambioAPct_(String(f.Cambio_Global||'')) };
}
function kepTendencia_(regsCron){
  try{ return calcularTendenciaRMYT_(regsCron); }catch(e){ return null; }
}

// ── PACIENTE: guarda la autoevaluación semanal (acción pública) ─────────
function guardarKepEvaluacion(d){
  const dni = String(d.dni||'').replace(/\D/g,'');
  if(!dni) return {success:false, error:'Falta el DNI en el link — pedile a tu kinesiólogo/a que te lo reenvíe.'};
  const pac = kepPaciente_(dni);
  if(!pac) return {success:false, error:'No encontramos tu ficha en el consultorio. Avisale a tu kinesiólogo/a.'};
  const v = d.datos || {};
  const num = function(x, lo, hi){ const n = Number(x); return (x===''||x===null||x===undefined||isNaN(n)) ? null : Math.max(lo, Math.min(hi, n)); };
  const sh = kepHoja_('KEP_Evaluaciones', KEP_HDR_EVAL_);
  const filas = kepFilas_(sh);
  const propias = filas.filter(function(f){ return String(f.DNI)===dni; });
  // tope diario por DNI (evita llenar la hoja si alguien adivina DNI)
  const hoy = new Date().toISOString().slice(0,10);
  if(propias.filter(function(f){ return String(f.Timestamp).slice(0,10)===hoy; }).length >= KEP_MAX_ENVIOS_DIA_) return {success:false, error:'Ya cargaste muchas evaluaciones hoy. Probá mañana.'};

  const psfsItems = (Array.isArray(v.psfs) ? v.psfs : []).slice(0,5).map(function(x){ return { actividad: String(x.actividad||'').slice(0,120), valor: num(x.valor,0,10) }; }).filter(function(x){ return x.actividad && x.valor!==null; });
  const psfsProm = psfsItems.length ? Math.round(psfsItems.reduce(function(a,b){ return a+b.valor; },0)/psfsItems.length*10)/10 : null;
  const banderas = (Array.isArray(v.banderas) ? v.banderas : []).filter(function(b){ return KEP_BANDERAS_[b]; });
  const cur = {
    eva_mov: num(v.eva_movimiento,0,10), eva_reposo: num(v.eva_reposo,0,10), psfs: psfsProm,
    cambio: KEP_CAMBIO_GLOBAL_.indexOf(v.cambio_global)>-1 ? v.cambio_global : '',
    tolerancia: ['Buena','Regular','Mala'].indexOf(v.tolerancia)>-1 ? v.tolerancia : '',
    banderas: banderas
  };
  const cE = num(v.cumplimiento_ejercicios,0,10), cS = num(v.cumplimiento_sesiones,0,10);
  const vc = [cE,cS].filter(function(x){ return x!==null; });
  cur.cump = vc.length ? Math.round(vc.reduce(function(a,b){ return a+b; },0)/vc.length*10)/10 : null;
  const prevF = propias.length ? propias[propias.length-1] : null;
  const prev = prevF ? kepFilaAReg_(prevF) : null;
  const ev = kepEvaluarAlertas_(cur, prev, null);

  const fila = {
    Timestamp: new Date().toISOString(), DNI: dni, Paciente: kepNombre_(pac, dni), Region: String(d.region||v.region||'').slice(0,60),
    Cumplimiento_Ejercicios: cE===null?'':cE, Cumplimiento_Sesiones: cS===null?'':cS,
    EVA_Reposo: cur.eva_reposo===null?'':cur.eva_reposo, EVA_Movimiento: cur.eva_mov===null?'':cur.eva_mov,
    Dolor_Nocturno: banderas.indexOf('noche_progresivo')>-1 ? 'Sí' : '',
    PSFS_Prom: psfsProm===null?'':psfsProm, PSFS_JSON: JSON.stringify(psfsItems), Cambio_Global: cur.cambio, Tolerancia: cur.tolerancia,
    Banderas: banderas.join(', '), Alerta: ev.alertas.join(' + '), Nivel: ev.nivel,
    Dificultades: String(v.dificultades||'').slice(0,1500), Datos_JSON: JSON.stringify({ banderas: banderas, version: 'KEP-1' })
  };
  const h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  sh.appendRow(h.map(function(c){ return fila[c]===undefined ? '' : fila[c]; }));

  // Pregunta opcional para el kinesiólogo: usa el mismo canal que ya existía
  const pregunta = String(v.pregunta||'').trim();
  if(pregunta) guardarPreguntaKine({dni:dni, paciente:kepNombre_(pac,dni), pregunta:pregunta.slice(0,1500)});

  // Avisos (nunca bloquean el guardado)
  try{ kepNotificar_(pac, dni, ev, cur, String(d.region||v.region||''), String(v.dificultades||'')); }catch(e){ Logger.log('KEP notificar: '+e); }

  let mensaje = 'Gracias, tu kinesiólogo/a ya recibió tu autoevaluación.';
  if(ev.nivel === 'rojo') mensaje = 'Gracias. Por lo que marcaste, avisamos a tu kinesiólogo/a y al Dr. Muratori para que te contacten. Si tenés dificultad para controlar la orina o la materia fecal, adormecimiento en la entrepierna o perdés fuerza rápidamente, no esperes: andá a una guardia ahora.';
  else if(ev.nivel === 'amarillo') mensaje = 'Gracias, tu kinesiólogo/a va a revisar lo que marcaste.';
  return {success:true, nivel:ev.nivel, mensaje:mensaje};
}

function kepSeguimientoDe_(dni){
  try{
    const sh = kepHoja_('KEP_Seguimientos', KEP_HDR_SEG_);
    const l = kepFilas_(sh).filter(function(f){ return String(f.DNI)===String(dni) && String(f.Activo)!=='no'; });
    return l.length ? l[l.length-1] : null;
  }catch(e){ return null; }
}
function kepNotificar_(pac, dni, ev, cur, region, dificultades){
  if(ev.nivel === 'verde') return;
  const seg = kepSeguimientoDe_(dni);
  const nombre = kepNombre_(pac, dni);
  const txt = ev.alertas.map(function(a){ return a.replace(/_/g,' '); }).join(', ');
  const cuerpo = 'Paciente: ' + nombre + ' (DNI ' + dni + ')\n' +
    (region ? 'Región: ' + region + '\n' : '') +
    'Semáforo: ' + ev.nivel.toUpperCase() + ' — ' + txt + '\n\n' +
    'EVA en movimiento: ' + (cur.eva_mov==null?'—':cur.eva_mov) + ' · EVA en reposo: ' + (cur.eva_reposo==null?'—':cur.eva_reposo) + ' · PSFS promedio: ' + (cur.psfs==null?'—':cur.psfs) + ' · Percepción: ' + (cur.cambio||'—') + '\n' +
    (cur.banderas.length ? 'Banderas marcadas: ' + cur.banderas.map(function(b){ return KEP_BANDERAS_[b]; }).join('; ') + '\n' : '') +
    (dificultades ? 'Dificultades: ' + dificultades.slice(0,400) + '\n' : '') +
    '\nEs un semáforo de ayuda automático: no es un diagnóstico y no reemplaza el criterio clínico.';
  const asunto = 'SIM-M KEP · ' + (ev.nivel==='rojo' ? '🔴 ALERTA' : '🟡 A revisar') + ' · ' + nombre;
  const destinos = [];
  if(seg && seg.Kinesiologo_Email) destinos.push(String(seg.Kinesiologo_Email));
  if(ev.nivel === 'rojo') destinos.push(MEDICO_MAIL_KINESIOLOGIA);
  const unicos = destinos.filter(function(x,i){ return x && destinos.indexOf(x)===i; });
  if(unicos.length) MailApp.sendEmail(unicos.join(','), asunto, cuerpo, { name:'SIM-M — Evaluador KEP' });
}

// ── KINESIÓLOGO ─────────────────────────────────────────────────────────
function getKep(dni){
  dni = String(dni||'').replace(/\D/g,'');
  if(!dni) return {success:false, error:'Falta el DNI'};
  const evs = kepFilas_(kepHoja_('KEP_Evaluaciones', KEP_HDR_EVAL_)).filter(function(f){ return String(f.DNI)===dni; });
  const regs = evs.map(kepFilaAReg_).sort(function(a,b){ return new Date(a.fecha)-new Date(b.fecha); });
  const meds = kepFilas_(kepHoja_('KEP_Mediciones', KEP_HDR_MED_)).filter(function(f){ return String(f.DNI)===dni; });
  const ders = kepFilas_(kepHoja_('KEP_Derivaciones', KEP_HDR_DER_)).filter(function(f){ return String(f.DNI)===dni; });
  const seg = kepSeguimientoDe_(dni);
  const ultima = regs.length ? regs[regs.length-1] : null;
  const dias = ultima ? Math.floor((Date.now()-new Date(ultima.fecha).getTime())/86400000) : null;
  return { success:true, dni:dni,
    evaluaciones: regs.slice().reverse().slice(0,30),
    mediciones: meds.map(function(m){ return { fecha:String(m.Timestamp), region:String(m.Region||''), kinesiologo:String(m.Kinesiologo||''), rom:kepJson_(m.ROM_JSON), fuerza:kepJson_(m.Fuerza_JSON), test:String(m.Test_Funcional||''), observaciones:String(m.Observaciones||'') }; }).sort(function(a,b){ return new Date(b.fecha)-new Date(a.fecha); }).slice(0,20),
    derivaciones: ders.map(function(x){ return { fila:x._row, fecha:String(x.Timestamp), motivo:String(x.Motivo||''), estado:String(x.Estado||''), respuesta:String(x.Respuesta_Medico||''), fecha_respuesta:String(x.Fecha_Respuesta||'') }; }).sort(function(a,b){ return new Date(b.fecha)-new Date(a.fecha); }),
    seguimiento: seg ? { region:String(seg.Region||''), kinesiologo:String(seg.Kinesiologo||''), desde:String(seg.Fecha_Inicio||''), ultimo_link:String(seg.Ultimo_Link||'') } : null,
    tendencia: kepTendencia_(regs), dias_desde_ultima: dias,
    sin_respuesta: dias!==null && dias > KEP_DIAS_SIN_RESPUESTA_,
    umbrales: { eva: KEP_UMBRAL_EVA_, psfs: KEP_UMBRAL_PSFS_, dias_sin_respuesta: KEP_DIAS_SIN_RESPUESTA_ } };
}
function kepJson_(s){ try{ const x = JSON.parse(String(s||'[]')); return x; }catch(e){ return []; } }

function guardarKepMedicion(d){
  const dni = String(d.dni||'').replace(/\D/g,'');
  if(!dni) return {success:false, error:'Falta el DNI'};
  const rom = (Array.isArray(d.rom)?d.rom:[]).slice(0,30).map(function(x){ return { mov:String(x.mov||'').slice(0,60), lado:String(x.lado||'').slice(0,20), grados:kepNum_(x.grados) }; }).filter(function(x){ return x.mov && x.grados!==null; });
  const fz = (Array.isArray(d.fuerza)?d.fuerza:[]).slice(0,30).map(function(x){ const g = kepNum_(x.mrc); return { musculo:String(x.musculo||'').slice(0,60), lado:String(x.lado||'').slice(0,20), mrc: g===null?null:Math.max(0,Math.min(5,Math.round(g))) }; }).filter(function(x){ return x.musculo && x.mrc!==null; });
  const test = String(d.test||'').slice(0,300), obs = String(d.observaciones||'').slice(0,1500);
  if(!rom.length && !fz.length && !test && !obs) return {success:false, error:'No cargaste ninguna medición'};
  const pac = kepPaciente_(dni);
  const sh = kepHoja_('KEP_Mediciones', KEP_HDR_MED_);
  const fila = { Timestamp:new Date().toISOString(), DNI:dni, Paciente:kepNombre_(pac,dni), Region:String(d.region||'').slice(0,60), Kinesiologo:String(d.kinesiologoNombre||'').slice(0,100), Kinesiologo_Email:String(d.kinesiologoEmail||'').slice(0,120), ROM_JSON:JSON.stringify(rom), Fuerza_JSON:JSON.stringify(fz), Test_Funcional:test, Observaciones:obs };
  const h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  sh.appendRow(h.map(function(c){ return fila[c]===undefined ? '' : fila[c]; }));
  return {success:true, message:'Mediciones guardadas'};
}

function kepLinkPaciente_(dni, nombre, region){
  return baseSimm_() + 'KEP_Evaluador_Paciente.html?dni=' + encodeURIComponent(dni) +
    (nombre ? '&nombre=' + encodeURIComponent(nombre) : '') + (region ? '&region=' + encodeURIComponent(region) : '');
}
function kepEnviarLink(d){
  const dni = String(d.dni||'').replace(/\D/g,'');
  if(!dni) return {success:false, error:'Falta el DNI'};
  const pac = kepPaciente_(dni);
  if(!pac) return {success:false, error:'El paciente no está en la base'};
  const email = String(d.emailPaciente || pac.email || '').trim();
  const region = String(d.region||'').slice(0,60);
  const link = kepLinkPaciente_(dni, pac.nombre, region);
  const sh = kepHoja_('KEP_Seguimientos', KEP_HDR_SEG_);
  const filas = kepFilas_(sh).filter(function(f){ return String(f.DNI)===dni; });
  const h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const reg = { DNI:dni, Paciente:kepNombre_(pac,dni), Region:region, Kinesiologo:String(d.kinesiologoNombre||'').slice(0,100), Kinesiologo_Email:String(d.kinesiologoEmail||'').slice(0,120), Fecha_Inicio:new Date().toISOString(), Ultimo_Link:new Date().toISOString(), Activo:'si' };
  if(filas.length){
    const f = filas[filas.length-1];
    reg.Fecha_Inicio = f.Fecha_Inicio || reg.Fecha_Inicio;
    sh.getRange(f._row,1,1,h.length).setValues([h.map(function(c){ return reg[c]===undefined ? '' : reg[c]; })]);
  } else sh.appendRow(h.map(function(c){ return reg[c]===undefined ? '' : reg[c]; }));
  if(!d.enviarMail) return {success:true, link:link, message:'Link generado'};
  if(!email) return {success:false, error:'El paciente no tiene mail cargado', link:link};
  const cuerpo = 'Hola ' + (pac.nombre||'') + ',\n\n' +
    'Para acompañar tu tratamiento de kinesiología te pedimos que completes, una vez por semana, esta breve evaluación (te lleva 2 minutos). Le llega directo a tu kinesiólogo/a y, si algo requiere atención médica, también al Dr. Muratori.\n\n' +
    link + '\n\n' +
    'Si en algún momento tenés dificultad para controlar la orina o la materia fecal, adormecimiento en la entrepierna o perdés fuerza rápidamente, no esperes el formulario: consultá en una guardia.\n\n' +
    'Consultorio Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943';
  MailApp.sendEmail(email, 'SIM-M · Evaluación semanal de kinesiología', cuerpo, { replyTo: d.kinesiologoEmail || MEDICO_MAIL_KINESIOLOGIA, name: d.kinesiologoNombre ? (d.kinesiologoNombre + ' — Consultorio Dr. Muratori') : 'Consultorio Dr. Muratori' });
  return {success:true, link:link, message:'Link enviado por mail al paciente'};
}

// El kinesiólogo deriva al médico (con motivo). Llega por mail a consultas.sism
// y queda en la pestaña "Kinesiología" del Seguimiento de OCM.
function kepDerivarMedico(d){
  const dni = String(d.dni||'').replace(/\D/g,'');
  const motivo = String(d.motivo||'').trim();
  if(!dni || !motivo) return {success:false, error:'Falta el DNI o el motivo'};
  const pac = kepPaciente_(dni);
  const nombre = kepNombre_(pac, dni);
  const resumen = String(d.resumen||'').slice(0,4000);
  const sh = kepHoja_('KEP_Derivaciones', KEP_HDR_DER_);
  const fila = { Timestamp:new Date().toISOString(), DNI:dni, Paciente:nombre, Kinesiologo:String(d.kinesiologoNombre||'').slice(0,100), Kinesiologo_Email:String(d.kinesiologoEmail||'').slice(0,120), Motivo:motivo.slice(0,1000), Resumen:resumen, Estado:'pendiente', Respuesta_Medico:'', Fecha_Respuesta:'' };
  const h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  sh.appendRow(h.map(function(c){ return fila[c]===undefined ? '' : fila[c]; }));
  try{
    MailApp.sendEmail(MEDICO_MAIL_KINESIOLOGIA, 'SIM-M KEP · Derivación de kinesiología · ' + nombre,
      'El/la kinesiólogo/a ' + (d.kinesiologoNombre||'') + ' deriva a ' + nombre + ' (DNI ' + dni + ').\n\nMotivo:\n' + motivo + '\n\n' + (resumen ? ('Resumen:\n' + resumen + '\n\n') : '') +
      'Podés responderle desde la pestaña "Kinesiología" del Seguimiento en OCM, o respondiendo este mail.',
      { replyTo: d.kinesiologoEmail || MEDICO_MAIL_KINESIOLOGIA, name:'SIM-M — Evaluador KEP' });
  }catch(e){ return {success:true, message:'Derivación guardada, pero no se pudo enviar el mail: ' + e}; }
  return {success:true, message:'Derivación enviada al médico'};
}
// El médico responde al kinesiólogo (indicación). Se guarda y llega por mail
// al kinesiólogo, con el mail del médico como respuesta directa.
function kepResponderMedico(d){
  const fila = parseInt(d.fila);
  const resp = String(d.respuesta||'').trim();
  if(!fila || !resp) return {success:false, error:'Falta la fila o la respuesta'};
  const sh = kepHoja_('KEP_Derivaciones', KEP_HDR_DER_);
  if(fila < 2 || fila > sh.getLastRow()) return {success:false, error:'Fila inexistente'};
  const h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const v = sh.getRange(fila,1,1,h.length).getValues()[0];
  const get = function(c){ return v[h.indexOf(c)]; };
  sh.getRange(fila, h.indexOf('Respuesta_Medico')+1).setValue(resp.slice(0,2000));
  sh.getRange(fila, h.indexOf('Fecha_Respuesta')+1).setValue(new Date().toISOString());
  sh.getRange(fila, h.indexOf('Estado')+1).setValue('respondida');
  const kmail = String(get('Kinesiologo_Email')||'');
  if(kmail){
    try{
      MailApp.sendEmail(kmail, 'SIM-M KEP · Respuesta del Dr. Muratori · ' + get('Paciente'),
        'El Dr. Muratori respondió tu derivación de ' + get('Paciente') + ' (DNI ' + get('DNI') + '):\n\n"' + String(get('Motivo')||'') + '"\n\nIndicación médica:\n' + resp + '\n\nConsultorio Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943',
        { replyTo:'drluismuratori@gmail.com', name:'Dr. Luis Alberto Muratori' });
    }catch(e){ return {success:true, message:'Respuesta guardada, pero no se pudo enviar el mail: ' + e}; }
  }
  return {success:true, message: kmail ? 'Respuesta guardada y enviada al kinesiólogo' : 'Respuesta guardada (la derivación no tenía mail del kinesiólogo)'};
}

// ── MÉDICO: una fila por paciente en seguimiento kinésico (pestaña de OCM) ──
function getKepDashboard(){
  const evs = kepFilas_(kepHoja_('KEP_Evaluaciones', KEP_HDR_EVAL_));
  const segs = kepFilas_(kepHoja_('KEP_Seguimientos', KEP_HDR_SEG_));
  const ders = kepFilas_(kepHoja_('KEP_Derivaciones', KEP_HDR_DER_));
  const dnis = {};
  evs.forEach(function(f){ dnis[String(f.DNI)] = 1; });
  segs.forEach(function(f){ if(String(f.Activo)!=='no') dnis[String(f.DNI)] = 1; });
  ders.forEach(function(f){ dnis[String(f.DNI)] = 1; });
  kinFilas_Ord_().forEach(function(f){ if(String(f.Estado)==='enviada') dnis[String(f.DNI)] = 1; });
  kinFilasHC_().forEach(function(f){ if(String(f.Estado)!=='alta' || !f.Fecha_Alta || (Date.now()-new Date(String(f.Fecha_Alta)).getTime()) < 90*86400000) dnis[String(f.DNI)] = 1; });
  const items = Object.keys(dnis).filter(Boolean).map(function(dni){
    const pac = kepPaciente_(dni);
    const regs = evs.filter(function(f){ return String(f.DNI)===dni; }).map(kepFilaAReg_).sort(function(a,b){ return new Date(a.fecha)-new Date(b.fecha); });
    const ult = regs.length ? regs[regs.length-1] : null;
    const dias = ult ? Math.floor((Date.now()-new Date(ult.fecha).getTime())/86400000) : null;
    const tend = kepTendencia_(regs);
    const seg = segs.filter(function(f){ return String(f.DNI)===dni; }).pop() || null;
    const dd = ders.filter(function(f){ return String(f.DNI)===dni; }).map(function(x){ return { fila:x._row, fecha:String(x.Timestamp), motivo:String(x.Motivo||''), resumen:String(x.Resumen||''), estado:String(x.Estado||''), respuesta:String(x.Respuesta_Medico||''), kinesiologo:String(x.Kinesiologo||'') }; }).sort(function(a,b){ return new Date(b.fecha)-new Date(a.fecha); });
    return { dni:dni, paciente:kepNombre_(pac,dni), obra_social:pac?pac.os:'', telefono:pac?pac.tel:'', mail:pac?pac.email:'', diagnostico:pac?pac.dx:'',
      kinesiologo: seg ? String(seg.Kinesiologo||'') : (dd[0] ? dd[0].kinesiologo : ''), region: seg ? String(seg.Region||'') : (ult ? ult.region : ''),
      n_evaluaciones: regs.length, ultima: ult ? { fecha:ult.fecha, eva_reposo:ult.eva_reposo, eva_movimiento:ult.eva_mov, psfs:ult.psfs, cambio:ult.cambio, cumplimiento_promedio:ult.cump, alerta:ult.alerta, nivel:ult.nivel, banderas:ult.banderas } : null,
      dias_desde_ultima: dias, sin_respuesta: dias!==null && dias > KEP_DIAS_SIN_RESPUESTA_,
      tendencia_estado: tend ? tend.estado : '', tendencia_mensaje: tend ? tend.mensaje : '',
      derivaciones: dd, derivacion_pendiente: dd.some(function(x){ return x.estado==='pendiente'; }), kin: kinResumenDe_(dni) };
  });
  return {success:true, items:items};
}

// ═════════════════════════════════════════════════════════════════════════
// PÁGINA DE SALUD (drmuratorisalud.ar) ↔ RMYT — v19 (oct 2026)
// El RMYT del paciente se coordina con las Claves A (tratamiento médico) y
// C (estrategia kinésica del equipo) de cada tema: un solo código le abre
// "cómo piensa y trata la patología el equipo profesional". La Clave B
// (quirúrgica) se habilita aparte, cuando se decide operar.
// PRIVACIDAD: a la página solo viaja tema + días + claves. El vínculo
// código ↔ paciente vive SOLO en esta planilla (hoja Accesos_Pagina_Salud).
// CONFIGURACIÓN (una vez): Propiedades del script → SIMM_CLAVE (la misma que
// el secreto SIMM_CLAVE de Cloudflare) y, opcional, SITIO_URL.
// La clave NUNCA va en archivos ni en GitHub.
// ═════════════════════════════════════════════════════════════════════════
const PS_HDR_ = ['Timestamp','DNI','Paciente','Tema','Codigo','Claves','Vence','Plan','Profesional','Estado','Enlace','Historial_JSON'];
const PS_MAX_POR_DIA_ = 40;

function psHoja_(){ return kepHoja_('Accesos_Pagina_Salud', PS_HDR_); }
function psHoy_(){ return new Date(Date.now()-3*3600*1000).toISOString().slice(0,10); }
function psApi_(ruta, cuerpo){
  const clave = PropertiesService.getScriptProperties().getProperty('SIMM_CLAVE');
  if(!clave) return {ok:false, error:'Falta cargar SIMM_CLAVE en las Propiedades del script (Configuración del proyecto).'};
  const sitio = String(PropertiesService.getScriptProperties().getProperty('SITIO_URL') || 'https://drmuratorisalud.ar').replace(/\/$/,'');
  try{
    const opc = { method: cuerpo ? 'post' : 'get', contentType:'application/json', headers:{ Authorization:'Bearer ' + clave }, muteHttpExceptions:true };
    if(cuerpo) opc.payload = JSON.stringify(cuerpo);
    const r = UrlFetchApp.fetch(sitio + '/api/integracion/' + ruta, opc);
    let d = {}; try{ d = JSON.parse(r.getContentText()||'{}'); }catch(e){}
    if(r.getResponseCode() >= 400) return {ok:false, error: d.error || ('La página respondió con error ' + r.getResponseCode())};
    return {ok:true, data:d};
  }catch(e){ return {ok:false, error:'No hay conexión con la página de salud.'}; }
}
function psTemasLista_(){
  const c = CacheService.getScriptCache();
  try{ const t = c.get('ps_temas'); if(t) return JSON.parse(t); }catch(e){}
  const r = psApi_('temas');
  if(!r.ok) return {error:r.error};
  const l = (r.data.temas||[]).map(function(t){ return {slug:String(t.slug), titulo:String(t.titulo||t.slug), claves:(t.claves||[]).join('')}; });
  try{ c.put('ps_temas', JSON.stringify(l), 600); }catch(e){}
  return l;
}
// MÉDICO: temas de la página con las claves que tienen hoy (A, B, C)
function getPaginaSaludTemas(){
  const l = psTemasLista_();
  if(l && l.error) return {success:false, error:l.error};
  return {success:true, temas:l};
}
function psFilas_(){
  const sh = psHoja_();
  return kepFilas_(sh).map(function(f){
    return { fila:f._row, dni:String(f.DNI), paciente:String(f.Paciente||''), tema:String(f.Tema||''), codigo:String(f.Codigo||''), claves:String(f.Claves||''),
      vence:String(f.Vence||'').slice(0,10), plan:String(f.Plan||''), profesional:String(f.Profesional||''), estado:String(f.Estado||''), enlace:String(f.Enlace||'') };
  });
}
function getPaginaSaludAccesos(dni){
  dni = String(dni||'').replace(/\D/g,'');
  if(!dni) return {success:false, error:'Falta el DNI'};
  const hoy = psHoy_();
  const l = psFilas_().filter(function(x){ return x.dni===dni && x.estado!=='revocado'; }).map(function(x){ x.vencido = !!x.vence && x.vence < hoy; return x; });
  return {success:true, accesos:l};
}
function psGuardarHist_(sh, fila, evento){
  const h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const iH = h.indexOf('Historial_JSON')+1;
  let lista = []; try{ lista = JSON.parse(String(sh.getRange(fila,iH,1,1).getValues()[0][0]||'[]')); }catch(e){}
  lista.push(Object.assign({fecha:new Date().toISOString()}, evento));
  sh.getRange(fila,iH).setValue(JSON.stringify(lista.slice(-30)));
}
function psActualizarFila_(sh, fila, campos){
  const h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  Object.keys(campos).forEach(function(k){ const i = h.indexOf(k); if(i>-1) sh.getRange(fila,i+1).setValue(campos[k]); });
}
function psSitio_(){ return String(PropertiesService.getScriptProperties().getProperty('SITIO_URL') || 'https://drmuratorisalud.ar').replace(/\/$/,''); }
function psMailAcceso_(pac, titulo, enlace, claves, vence, prof){
  if(!pac || !pac.email) return false;
  try{
    MailApp.sendEmail(pac.email, 'SIM-M · Tu tratamiento explicado por el equipo — ' + titulo,
      'Hola ' + pac.nombre + ',\n\n' + (claves.indexOf('C')>-1 && claves.indexOf('A')<0 ? 'El equipo de kinesiología que te atiende te habilitó la estrategia kinésica de tu problema: cómo la piensa y la trata el equipo, y en qué evidencia se apoya.' : 'En esta página el equipo profesional te explica cómo piensa y trata tu problema y en qué evidencia se apoya.') +
      '\n\nTu enlace personal (no lo compartas):\n' + enlace + '\n\nHabilitado hasta el ' + vence + '.\n\nLa página es informativa y no reemplaza la consulta. Si algo empeora o aparece una señal de alarma, consultá sin esperar.\n\nQuien te lo habilitó: ' + prof,
      { replyTo:'drluismuratori@gmail.com', name:'Consultorio Dr. Luis Alberto Muratori' });
    return true;
  }catch(e){ return false; }
}
// Crea (o completa) el acceso del paciente a un tema.
//  · RMYT (médico) emite la Clave A: claves 'A'. Si el tema aún no tiene A, con vincular:true se guarda solo el
//    enlace al artículo general (sin código) para que el RMYT conecte igual al paciente con el contenido.
//  · El GRUPO DE KINESIOLOGÍA entrega la Clave C (claves 'C') si el paciente es tratado por el equipo; se suma
//    al mismo código si el paciente ya tiene uno para ese tema.
function crearAccesoPaginaSalud(d){
  const dni = String(d.dni||'').replace(/\D/g,'');
  const tema = String(d.tema||'').trim();
  if(!dni) return {success:false, error:'Falta el DNI'};
  if(!/^[a-z0-9-]{1,80}$/.test(tema)) return {success:false, error:'Tema inválido'};
  const pac = kepPaciente_(dni);
  if(!pac) return {success:false, error:'El DNI no está cargado en Pacientes'};
  let dias = parseInt(d.dias); if(!(dias>=1 && dias<=365)) dias = 30;
  let pedidas = String(d.claves||'A').toUpperCase().replace(/[^ABC]/g,'');
  if(!pedidas) pedidas = 'A';
  const temas = psTemasLista_();
  if(temas && temas.error) return {success:false, error:temas.error};
  const t = temas.filter(function(x){ return x.slug===tema; })[0];
  if(!t) return {success:false, error:'Ese tema todavía no tiene esquema protegido en la página.'};
  const prof = String(d.profesional||'Dr. Muratori').slice(0,80);
  const disponibles = pedidas.split('').filter(function(c){ return t.claves.indexOf(c)>-1; });
  const sh = psHoja_(), hoy = psHoy_(), todas = psFilas_().filter(function(x){ return x.dni===dni && x.tema===tema && x.estado!=='revocado'; });
  const vinculo = todas.filter(function(x){ return !x.codigo; })[0];
  const vigentes = todas.filter(function(x){ return x.codigo && x.vence >= hoy; });
  const artGeneral = psSitio_() + '/articulos/' + tema + '.html';
  if(!disponibles.length){
    if(!d.vincular) return {success:false, error:'El tema todavía no tiene ' + (pedidas==='C' ? 'la Clave C' : 'las claves pedidas') + ' publicada (hoy tiene: ' + t.claves + ').'};
    if(vigentes.length || vinculo) return {success:true, reutilizado:true, sinCodigo:!vigentes.length, tema:tema, titulo:t.titulo, claves:vigentes.length?vigentes[vigentes.length-1].claves:'', enlace:vigentes.length?vigentes[vigentes.length-1].enlace:artGeneral, message:'El RMYT ya estaba vinculado a este tema.'};
    sh.appendRow([new Date().toISOString(), dni, (pac.apellido + ', ' + pac.nombre), tema, '', '', '', 'Vínculo al artículo general', prof, 'sin_codigo', artGeneral, JSON.stringify([{fecha:new Date().toISOString(), accion:'vincular', profesional:prof}])]);
    return {success:true, reutilizado:false, sinCodigo:true, tema:tema, titulo:t.titulo, claves:'', enlace:artGeneral, message:'Tema vinculado al RMYT (todavía sin Clave A publicada: el paciente ve el artículo general).'};
  }
  if(vigentes.length){
    const ex = vigentes[vigentes.length-1];
    let claves = ex.claves, enlace = ex.enlace, vence = ex.vence;
    const faltan = disponibles.filter(function(c){ return claves.indexOf(c)<0; });
    for(let i=0;i<faltan.length;i++){
      const r = psApi_('habilitar', {codigo:ex.codigo, clave:faltan[i]});
      if(!r.ok) return {success:false, error:r.error};
      claves = String(r.data.claves||claves); vence = String(r.data.vence||vence).slice(0,10); enlace = String(r.data.enlace||enlace);
    }
    if(faltan.length){
      psActualizarFila_(sh, ex.fila, {Claves:claves, Vence:vence, Enlace:enlace});
      psGuardarHist_(sh, ex.fila, {accion:'habilitar', claves:faltan.join(''), profesional:prof});
    }
    const mail = (d.enviarMail && faltan.length) ? psMailAcceso_(pac, t.titulo, enlace, claves, vence, prof) : false;
    return {success:true, reutilizado:true, codigo:ex.codigo, tema:tema, titulo:t.titulo, claves:claves, vence:vence, enlace:enlace, mailEnviado:mail,
      message: faltan.length ? ('Se sumó la Clave ' + faltan.join(' y ') + ' al código que el paciente ya tenía: no necesita un QR nuevo.') : 'El paciente ya tenía esa clave habilitada en su código vigente.'};
  }
  const hojaDatos = sh.getDataRange().getValues(), iT = hojaDatos[0].indexOf('Timestamp');
  const hoyN = hojaDatos.slice(1).filter(function(f){ return String(f[iT]).slice(0,10)===new Date().toISOString().slice(0,10); }).length;
  if(hoyN >= PS_MAX_POR_DIA_) return {success:false, error:'Se alcanzó el tope diario de códigos nuevos (' + PS_MAX_POR_DIA_ + ').'};
  const r = psApi_('codigo', {tema:tema, dias:dias, claves:disponibles.join('')});
  if(!r.ok) return {success:false, error:r.error};
  const x = r.data, venceX = String(x.vence).slice(0,10);
  const plan = String(d.plan||('Seguimiento ' + dias + ' días')).slice(0,80);
  const hist = JSON.stringify([{fecha:new Date().toISOString(), accion:'crear', claves:String(x.claves), vence:venceX, profesional:prof}]);
  if(vinculo){
    psActualizarFila_(sh, vinculo.fila, {Codigo:String(x.codigo), Claves:String(x.claves), Vence:venceX, Plan:plan, Profesional:prof, Estado:'activo', Enlace:String(x.enlace||''), Historial_JSON:hist});
  } else {
    sh.appendRow([new Date().toISOString(), dni, (pac.apellido + ', ' + pac.nombre), tema, String(x.codigo), String(x.claves), venceX, plan, prof, 'activo', String(x.enlace||''), hist]);
  }
  const mail = d.enviarMail ? psMailAcceso_(pac, t.titulo, String(x.enlace||''), String(x.claves), venceX, prof) : false;
  return {success:true, reutilizado:false, codigo:String(x.codigo), tema:tema, titulo:t.titulo, claves:String(x.claves), vence:venceX, enlace:String(x.enlace||''), mailEnviado:mail, message:'Código creado con la clave ' + String(x.claves).split('').join(' + ')};
}

// MÉDICO: habilita una clave más (típico: B cuando se decide operar)
function habilitarClavePaginaSalud(d){
  const codigo = String(d.codigo||'').toUpperCase().replace(/\s/g,'');
  const clave = String(d.clave||'').toUpperCase();
  if(!codigo || !/^[ABC]$/.test(clave)) return {success:false, error:'Falta el código o la clave (A, B o C)'};
  const f = psFilas_().filter(function(x){ return x.codigo===codigo; })[0];
  if(!f) return {success:false, error:'Ese código no es de este sistema'};
  const r = psApi_('habilitar', {codigo:codigo, clave:clave});
  if(!r.ok) return {success:false, error:r.error};
  const sh = psHoja_();
  psActualizarFila_(sh, f.fila, {Claves:String(r.data.claves||''), Vence:String(r.data.vence||'').slice(0,10)});
  psGuardarHist_(sh, f.fila, {accion:'habilitar', claves:clave, profesional:String(d.profesional||'')});
  return {success:true, codigo:codigo, claves:String(r.data.claves||''), vence:String(r.data.vence||'').slice(0,10), message:'Clave ' + clave + ' habilitada'};
}
// MÉDICO: extiende la vigencia (membresía de seguimiento)
function extenderAccesoPaginaSalud(d){
  const codigo = String(d.codigo||'').toUpperCase().replace(/\s/g,'');
  let dias = parseInt(d.dias); if(!(dias>=1 && dias<=365)) dias = 30;
  if(!codigo) return {success:false, error:'Falta el código'};
  const f = psFilas_().filter(function(x){ return x.codigo===codigo; })[0];
  if(!f) return {success:false, error:'Ese código no es de este sistema'};
  const r = psApi_('extender', {codigo:codigo, dias:dias});
  if(!r.ok) return {success:false, error:r.error};
  const sh = psHoja_();
  psActualizarFila_(sh, f.fila, {Vence:String(r.data.vence||'').slice(0,10), Estado:'activo', Plan:'Extensión ' + dias + ' días'});
  psGuardarHist_(sh, f.fila, {accion:'extender', dias:dias, vence:String(r.data.vence||''), profesional:String(d.profesional||'')});
  return {success:true, codigo:codigo, vence:String(r.data.vence||'').slice(0,10), enlace:f.enlace, message:'Vigencia extendida hasta ' + String(r.data.vence||'').slice(0,10)};
}

// ═════════════════════════════════════════════════════════════════════════
// KINESIOLOGÍA — portada con QR: inscripción → presente → salida (v20, oct 2026)
// Un mismo QR del centro sirve para todo: la 1ª vez el paciente se inscribe
// (filiación, lesión, médico prescriptor, kinesiólogo actuante, foto de la
// orden, consentimiento firmado en el celular); las siguientes marca su
// PRESENTE (firma + cómo llega) y al terminar su SALIDA (sensación, comentario,
// conformidad y firma). Todo alimenta la historia clínica kinésica, el
// seguimiento mensual, la concordancia paciente/terapeuta y la bitácora del alta.
// Acciones PÚBLICAS (las usa el paciente desde el QR, sin token, como RMYT):
// kin_identificar, kin_inscribir, kin_entrada, kin_salida. Se identifican con
// DNI + fecha de nacimiento, con tope de intentos, y nunca devuelven datos
// clínicos (solo el nombre de pila y el estado de la sesión).
// ═════════════════════════════════════════════════════════════════════════
const KIN_HDR_HC_ = ['Timestamp','DNI','Apellido','Nombre','Fecha_Nac','Sexo','Telefono','Email','Obra_Social','Afiliado','Region','Lado','Diagnostico','Fecha_Lesion','Operado','Medico_Tipo','Medico_Nombre','Kinesiologo','Kinesiologo_Email','Centro','Sesiones_Prescriptas','Turno_Preferido','Orden_URL','Firma_Consent','Estado','Fecha_Alta','Tipo_Alta','Resumen_Estudios','Bitacora_JSON','Consentimiento','Membresia','Membresia_Vence','Orden_Medica'];
const KIN_HDR_ORD_ = ['Timestamp','DNI','Paciente','Telefono','Email','Diagnostico','Region','Lado','Sesiones','Frecuencia','Indicacion','Medico','Resumen_Estudios','Estado','HC_Timestamp','Fecha_Vinculo','Centro','Confirmacion'];
const KIN_HDR_NOT_ = ['Timestamp','DNI','Autor_Tipo','Autor','Texto'];
const KIN_MAIL_MEDICO_ = 'drluismuratori@gmail.com';
const KIN_HDR_SES_ = ['Timestamp','DNI','Paciente','Episodio','N_Sesion','Fecha','Centro','Kinesiologo','Entrada_Hora','Sensacion_Entrada','Firma_Entrada','Salida_Hora','Sensacion_Salida','Comentario','Conforme','Firma_Salida','Estado','Hecho_Kine','Puntos_Eval','Sensacion_Kine','Evolucion_Kine','Nota_Hora'];
const KIN_MAX_INSCRIP_DIA_ = 40;
const KIN_KINE_DEFAULT_ = 'Fernando Polidori';   // kinesiólogo actuante por defecto (Polidori Kinesiología)
const KIN_MARCA_ = 'Polidori Kinesiología';
const KIN_MAX_INTENTOS_ = 8;
const KIN_CENTROS_ = {sicoli:'Hospital Sicoli', espanol:'Hospital Español', palmares:'Centro Médico Palmares', sandona:'Clínica San Donà', altomaipu:'Centro Médico Alto Maipú', rehab:'Rehabilitarte Instituto Médico', norte:'Consultorios del Norte', consultorio:'Consultorio Dr. Muratori'};

function kinHojaHC_(){ const sh = kepHoja_('KIN_HC', KIN_HDR_HC_); try{ sh.getRange('E:E').setNumberFormat('@'); }catch(e){} return sh; }
function kinHojaSes_(){ return kepHoja_('KIN_Sesiones', KIN_HDR_SES_); }
function kinDni_(v){ return String(v==null?'':v).replace(/\D/g,''); }
function kinTexto_(v, n){ return String(v==null?'':v).replace(/[\u0000-\u001f]/g,' ').trim().slice(0,n); }
function kinEsc_(s){ return String(s==null?'':s).replace(/[&<>"']/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]; }); }
function kinMailOk_(m){ return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(m||'')) && String(m).length<=120; }
function kinNum_(v, min, max){ if(v===''||v==null) return null; const n = Number(v); if(!isFinite(n)) return null; return Math.max(min, Math.min(max, Math.round(n*10)/10)); }
function kinFecha_(v){
  if(v instanceof Date) return v.toISOString().slice(0,10);
  const s = String(v==null?'':v).trim();
  let a=null, m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if(m) a=[+m[1],+m[2],+m[3]];
  else { m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/); if(m) a=[+m[3],+m[2],+m[1]]; }
  if(!a || a[0]<1900 || a[0]>2100 || a[1]<1 || a[1]>12 || a[2]<1 || a[2]>31) return '';
  return a[0] + '-' + ('0'+a[1]).slice(-2) + '-' + ('0'+a[2]).slice(-2);
}
function kinFilasHC_(){ return kepFilas_(kinHojaHC_()); }
function kinFilasSes_(){ return kepFilas_(kinHojaSes_()); }
function kinUltimoHC_(dni){ const l = kinFilasHC_().filter(function(f){ return String(f.DNI)===dni; }); return l.length ? l[l.length-1] : null; }
function kinBloqueado_(dni){ try{ return Number(CacheService.getScriptCache().get('kinf_'+dni)||0) >= KIN_MAX_INTENTOS_; }catch(e){ return false; } }
function kinFallo_(dni){ try{ const c = CacheService.getScriptCache(); c.put('kinf_'+dni, String(Number(c.get('kinf_'+dni)||0)+1), 3600); }catch(e){} }
function kinAuth_(dniRaw, fnacRaw){
  const dni = kinDni_(dniRaw), fnac = kinFecha_(fnacRaw);
  if(dni.length<6 || dni.length>9 || !fnac) return {error:'Revisá el DNI y la fecha de nacimiento.'};
  if(kinBloqueado_(dni)) return {error:'Demasiados intentos. Probá de nuevo en una hora o consultá en recepción.'};
  const hc = kinUltimoHC_(dni);
  if(!hc || kinFecha_(hc.Fecha_Nac)!==fnac){ kinFallo_(dni); return {error:'No encontramos ese DNI con esa fecha de nacimiento. Si es tu primera vez, tocá "Inscribirme".'}; }
  return {hc:hc, dni:dni};
}
function kinImagen_(dataUrl, carpeta, nombre, maxChars){
  try{
    const m = String(dataUrl||'').match(/^data:(image\/(?:png|jpeg|jpg|webp));base64,([A-Za-z0-9+\/=]+)$/);
    if(!m || m[2].length > maxChars) return '';
    const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], nombre);
    const it = DriveApp.getFoldersByName(carpeta);
    const f = it.hasNext() ? it.next() : DriveApp.createFolder(carpeta);
    return f.createFile(blob).getUrl();
  }catch(e){ return ''; }
}
function kinFirmaOk_(f){ return /^data:image\/png;base64,/.test(String(f||'')) && String(f).length > 1500 && String(f).length < 150000; }
function kinSesionesEpisodio_(dni, ep){ return kinFilasSes_().filter(function(s){ return String(s.DNI)===dni && String(s.Episodio)===String(ep); }); }
function kinSesObj_(s){
  return { fila:s._row, n:Number(s.N_Sesion)||0, fecha:String(s.Fecha||'').slice(0,10), entrada:String(s.Entrada_Hora||''), salida:String(s.Salida_Hora||''), sens_entrada:kepNum_(s.Sensacion_Entrada), sens_salida:kepNum_(s.Sensacion_Salida),
    comentario:String(s.Comentario||''), conforme:String(s.Conforme||''), estado:String(s.Estado||''), firma_entrada:String(s.Firma_Entrada||''), firma_salida:String(s.Firma_Salida||''),
    hecho:String(s.Hecho_Kine||''), puntos:String(s.Puntos_Eval||''), sens_kine:kepNum_(s.Sensacion_Kine), evolucion_kine:String(s.Evolucion_Kine||''), kinesiologo:String(s.Kinesiologo||'') };
}
// Concordancia paciente ↔ terapeuta (misma escala 0-10: 0 muy mal, 10 muy bien). Criterio práctico, no validado:
// diferencia ≤1 concuerda; 2-3 parcial; ≥4 discrepa.
function kinConcordancia_(ses){
  let c=0, p=0, d=0;
  ses.forEach(function(s){
    if(s.sens_salida==null || s.sens_kine==null) return;
    const dif = Math.abs(s.sens_salida - s.sens_kine);
    if(dif<=1) c++; else if(dif<=3) p++; else d++;
  });
  const n = c+p+d;
  return { n:n, concuerdan:c, parcial:p, discrepan:d, pct: n ? Math.round(c/n*100) : null };
}
function kinResumenEstudios_(dni){
  try{
    const r = getRecomendaciones(dni), c = (r && r.datos && r.datos.consulta) || {};
    const pac = kepPaciente_(dni);
    const partes = [];
    const dx = c.dx || (pac && pac.dx) || '';
    if(dx) partes.push('Diagnóstico médico: ' + dx + (c.cie10 ? ' (CIE-10 ' + c.cie10 + ')' : ''));
    if(c.fecha) partes.push('Última consulta médica: ' + c.fecha);
    if(c.resumen) partes.push('Resumen de la consulta: ' + c.resumen);
    if(c.estudios_realizados) partes.push('Estudios realizados: ' + c.estudios_realizados);
    if(c.estudios) partes.push('Hallazgos: ' + c.estudios);
    if(c.estudios_solicitados) partes.push('Estudios solicitados: ' + c.estudios_solicitados);
    if(c.indicaciones) partes.push('Indicaciones médicas: ' + c.indicaciones);
    return partes.join('\n').slice(0,3500);
  }catch(e){ return ''; }
}
function kinMailKine_(asunto, cuerpo, kmail){
  let kp = ''; try{ kp = String(PropertiesService.getScriptProperties().getProperty('KIN_MAIL_KINE')||''); }catch(e){}
  const para = [MEDICO_MAIL_KINESIOLOGIA].concat(kinMailOk_(kmail) ? [kmail] : []).concat(kinMailOk_(kp) ? [kp] : []).filter(function(x,i,a){ return a.indexOf(x)===i; }).join(',');
  try{ MailApp.sendEmail(para, 'SIM-M KIN · ' + asunto, cuerpo, {name:'S.I.M-M Kinesiología'}); return true; }catch(e){ return false; }
}

// ── PÚBLICO: ¿ya está inscripto? (decide si muestra presente o salida) ──
function kinIdentificar(d){
  const a = kinAuth_(d.dni, d.fnac);
  if(a.error) return {success:false, error:a.error};
  const hc = a.hc, ses = kinSesionesEpisodio_(a.dni, hc.Timestamp);
  const abierta = ses.filter(function(s){ return String(s.Estado)==='abierta'; }).pop();
  return { success:true, nombre:String(hc.Nombre||'').split(' ')[0], estado:String(hc.Estado||'activo'), sesionAbierta:!!abierta, n_proxima:(abierta?Number(abierta.N_Sesion):(ses.length+1)),
    prescriptas:Number(hc.Sesiones_Prescriptas)||0, hechas:ses.filter(function(s){ return String(s.Estado)==='cerrada'; }).length,
    membresia:kinMembEstado_(hc), membresia_vence:String(hc.Membresia_Vence||'').slice(0,10) };
}

// ── PÚBLICO: inscripción = primera vez (historia clínica kinésica inicial) ──
function kinInscribir(d){
  const dni = kinDni_(d.dni), fnac = kinFecha_(d.fnac);
  const apellido = kinTexto_(d.apellido,60), nombre = kinTexto_(d.nombre,60), tel = kinTexto_(d.telefono,30), email = kinTexto_(d.email,120);
  const region = kinTexto_(d.region,60), dx = kinTexto_(d.diagnostico,400);
  if(dni.length<6 || dni.length>9) return {success:false, error:'Revisá el DNI.'};
  if(!fnac) return {success:false, error:'Revisá la fecha de nacimiento.'};
  if(!apellido || !nombre) return {success:false, error:'Falta el nombre y apellido.'};
  if(!tel) return {success:false, error:'Falta un teléfono de contacto.'};
  if(email && !kinMailOk_(email)) return {success:false, error:'El mail no parece válido.'};
  if(!region || !dx) return {success:false, error:'Contanos qué zona te lesionaste y qué te pasó.'};
  if(d.consentimiento!==true) return {success:false, error:'Hace falta aceptar el consentimiento para inscribirte.'};
  if(!kinFirmaOk_(d.firma)) return {success:false, error:'Falta tu firma en el recuadro.'};
  const medicoTipo = d.medico_tipo==='equipo' ? 'equipo' : 'externo';
  const medicoNombre = kinTexto_(d.medico_nombre, 80) || (medicoTipo==='equipo' ? 'Dr. Muratori / equipo' : '');
  if(medicoTipo==='externo' && !medicoNombre) return {success:false, error:'Indicá el nombre del médico que te indicó kinesiología.'};
  const sh = kinHojaHC_();
  const todas = kepFilas_(sh);
  const previo = todas.filter(function(f){ return String(f.DNI)===dni; }).pop();
  if(previo && String(previo.Estado)!=='alta') return {success:false, error:'Ya estás inscripto/a. Para marcar tu asistencia tocá "Ya soy paciente".'};
  if(todas.filter(function(f){ return String(f.Timestamp).slice(0,10)===new Date().toISOString().slice(0,10); }).length >= KIN_MAX_INSCRIP_DIA_) return {success:false, error:'Se alcanzó el tope de inscripciones de hoy. Consultá en recepción.'};
  const centro = KIN_CENTROS_[String(d.centro||'')] ? String(d.centro) : 'consultorio';
  const orden = kinImagen_(d.orden, 'SIMM - Kinesiología (órdenes y firmas)', 'orden_' + dni + '_' + Date.now(), 6000000);
  const firma = kinImagen_(d.firma, 'SIMM - Kinesiología (órdenes y firmas)', 'consentimiento_' + dni + '_' + Date.now(), 200000);
  const sesiones = Math.max(0, Math.min(60, parseInt(d.sesiones)||0));
  const ord = kinOrdenAbierta_(dni);          // orden que el médico ya envió al centro (el paciente aceptó hacer la FKT acá)
  const medicoTipoF = ord ? 'equipo' : medicoTipo, medicoNombreF = ord ? (String(ord.Medico||'') || 'Dr. Muratori / equipo') : medicoNombre;
  const sesionesF = sesiones || (ord ? (Number(ord.Sesiones)||0) : 0);
  const estudios = medicoTipoF==='equipo' ? kinResumenEstudios_(dni) : '';
  const ts = new Date().toISOString();
  sh.appendRow([ts, dni, apellido, nombre, fnac, kinTexto_(d.sexo,12), tel, email, kinTexto_(d.os,60), kinTexto_(d.afiliado,40), region, kinTexto_(d.lado,20), dx, kinFecha_(d.fecha_lesion), d.operado===true?'si':'no', medicoTipoF, medicoNombreF,
    kinTexto_(d.kinesiologo,80) || KIN_KINE_DEFAULT_, kinMailOk_(d.kinesiologo_email)?String(d.kinesiologo_email):'', centro, sesionesF||'', kinTexto_(d.turno_pref,120), orden, firma, 'activo', '', '', estudios, '', 'aceptado ' + ts, d.membresia===true ? 'solicitada' : '', '', ord ? kinTexto_(String(ord.Indicacion||''),800) : '']);
  if(ord) kinOrdenVincular_(ord, ts, apellido + ', ' + nombre);
  kinMailKine_('Nueva inscripción — ' + apellido + ', ' + nombre,
    'Se inscribió un paciente nuevo desde el QR de ' + KIN_CENTROS_[centro] + '.\n\nPaciente: ' + apellido + ', ' + nombre + ' — DNI ' + dni + '\nTeléfono: ' + tel + (email?'\nMail: '+email:'') + '\nZona: ' + region + (d.lado?' ('+kinTexto_(d.lado,20)+')':'') + '\nMotivo / lesión: ' + dx +
    '\nMédico prescriptor: ' + medicoNombre + (medicoTipo==='equipo'?' (equipo del consultorio)':'') + '\nSesiones prescriptas: ' + (sesiones||'sin indicar') + '\nTurno preferido: ' + (kinTexto_(d.turno_pref,120)||'sin indicar') + '\nOrden médica: ' + (orden?'adjunta en Drive':'no adjuntó foto') +
    (estudios?'\n\nResumen de estudios del consultorio médico:\n'+estudios:'') + '\n\nDesde KSM podés ver su historia clínica kinésica.', d.kinesiologo_email);
  let mailPac = false;
  if(email){
    try{ MailApp.sendEmail(email, 'SIM-M · Recibimos tu inscripción de kinesiología', 'Hola ' + nombre + ',\n\nTu inscripción quedó registrada. Nos vamos a comunicar para coordinar tu turno.\n\nCada vez que vengas, escaneá el mismo código QR del centro: te va a pedir DNI y fecha de nacimiento para marcar tu presente y, al terminar, tu salida con tu firma.\n\nCualquier duda, respondé este mail.', {replyTo:MEDICO_MAIL_KINESIOLOGIA, name:KIN_MARCA_}); mailPac = true; }catch(e){}
  }
  return {success:true, message:'Listo, ' + nombre + ': quedaste inscripto/a. La próxima vez que vengas escaneá el mismo QR para marcar tu presente.', ordenGuardada:!!orden, mailEnviado:mailPac, equipo: medicoTipoF==='equipo', ordenVinculada: !!ord};
}

// ── PÚBLICO: PRESENTE (la firma se estampa) ──
function kinEntrada(d){
  const a = kinAuth_(d.dni, d.fnac);
  if(a.error) return {success:false, error:a.error};
  const hc = a.hc, dni = a.dni;
  if(String(hc.Estado)==='alta') return {success:false, error:'Tu tratamiento ya tiene el alta. Si necesitás más sesiones, consultá a tu médico.'};
  if(!kinFirmaOk_(d.firma)) return {success:false, error:'Falta tu firma en el recuadro.'};
  const sh = kinHojaSes_(), hoy = psHoy_();
  const ses = kinSesionesEpisodio_(dni, hc.Timestamp);
  const abiertas = ses.filter(function(s){ return String(s.Estado)==='abierta'; });
  if(abiertas.some(function(s){ return String(s.Fecha).slice(0,10)===hoy; })) return {success:false, abierta:true, error:'Ya marcaste tu presente hoy. Para terminar, escaneá de nuevo y registrá tu salida.'};
  if(ses.some(function(s){ return String(s.Fecha).slice(0,10)===hoy && String(s.Estado)==='cerrada'; })) return {success:false, error:'Ya registraste tu sesión de hoy. ¡Hasta la próxima!'};
  abiertas.forEach(function(s){ sh.getRange(s._row, KIN_HDR_SES_.indexOf('Estado')+1).setValue('sin_salida'); });
  const n = ses.length + 1;
  const firma = kinImagen_(d.firma, 'SIMM - Kinesiología (órdenes y firmas)', 'presente_' + dni + '_' + hoy + '_' + n, 200000);
  const sens = kinNum_(d.sensacion, 0, 10);
  sh.appendRow([new Date().toISOString(), dni, String(hc.Apellido)+', '+String(hc.Nombre), String(hc.Timestamp), n, hoy, String(hc.Centro||''), String(hc.Kinesiologo||''), new Date().toISOString(), sens==null?'':sens, firma, '', '', '', '', '', 'abierta', '', '', '', '', '']);
  const presc = Number(hc.Sesiones_Prescriptas)||0;
  return {success:true, n_sesion:n, prescriptas:presc, message:'Presente registrado. Sesión ' + n + (presc?' de ' + presc:'') + '. ¡Que tengas una buena sesión!' + (presc && n>presc ? ' Ya completaste las sesiones prescriptas: consultá con tu kinesiólogo.' : '')};
}

// ── PÚBLICO: SALIDA (sensación, comentario, conformidad y firma) ──
function kinSalida(d){
  const a = kinAuth_(d.dni, d.fnac);
  if(a.error) return {success:false, error:a.error};
  const hc = a.hc, dni = a.dni;
  const sens = kinNum_(d.sensacion, 0, 10);
  if(sens==null) return {success:false, error:'Marcá cómo te sentís al terminar (0 a 10).'};
  const conforme = d.conforme==='conforme' ? 'conforme' : (d.conforme==='con_reparos' ? 'con_reparos' : '');
  if(!conforme) return {success:false, error:'Indicá si estás conforme con la atención de hoy.'};
  if(!kinFirmaOk_(d.firma)) return {success:false, error:'Falta tu firma en el recuadro.'};
  const sh = kinHojaSes_();
  const abierta = kinSesionesEpisodio_(dni, hc.Timestamp).filter(function(s){ return String(s.Estado)==='abierta'; }).pop();
  if(!abierta) return {success:false, error:'No hay un presente abierto. Primero marcá tu presente.'};
  const firma = kinImagen_(d.firma, 'SIMM - Kinesiología (órdenes y firmas)', 'salida_' + dni + '_' + abierta.Fecha + '_' + abierta.N_Sesion, 200000);
  const col = function(c){ return KIN_HDR_SES_.indexOf(c)+1; };
  const comentario = kinTexto_(d.comentario, 600);
  sh.getRange(abierta._row, col('Salida_Hora')).setValue(new Date().toISOString());
  sh.getRange(abierta._row, col('Sensacion_Salida')).setValue(sens);
  sh.getRange(abierta._row, col('Comentario')).setValue(comentario);
  sh.getRange(abierta._row, col('Conforme')).setValue(conforme);
  sh.getRange(abierta._row, col('Firma_Salida')).setValue(firma);
  sh.getRange(abierta._row, col('Estado')).setValue('cerrada');
  if(sens<=2 || conforme==='con_reparos'){
    kinMailKine_('Atención: salida con ' + (sens<=2 ? 'sensación baja (' + sens + '/10)' : 'reparos') + ' — ' + String(hc.Apellido), String(hc.Apellido)+', '+String(hc.Nombre)+' (DNI '+dni+') cerró la sesión '+abierta.N_Sesion+' con sensación '+sens+'/10 y estado "'+conforme+'".'+(comentario?'\nComentario: '+comentario:''), hc.Kinesiologo_Email);
  }
  return {success:true, message:'¡Gracias, ' + String(hc.Nombre).split(' ')[0] + '! Tu salida quedó registrada con tu firma y tu sensación pasa a tu seguimiento.'};
}

// ── KINESIÓLOGO: historia clínica kinésica completa ──
function getKinHC(dni){
  dni = kinDni_(dni);
  if(!dni) return {success:false, error:'Falta el DNI'};
  const hc = kinUltimoHC_(dni);
  if(!hc) return {success:true, hc:null, orden:kinOrdenObj_(kinOrdenUltima_(dni)), notas:kinNotas_(dni)};
  let estudios = String(hc.Resumen_Estudios||'');
  if(String(hc.Medico_Tipo)==='equipo' && String(hc.Estado)!=='alta'){
    const nuevo = kinResumenEstudios_(dni);
    if(nuevo && nuevo!==estudios){ estudios = nuevo; kinHojaHC_().getRange(hc._row, KIN_HDR_HC_.indexOf('Resumen_Estudios')+1).setValue(nuevo); }
  }
  const ses = kinSesionesEpisodio_(dni, hc.Timestamp).map(kinSesObj_);
  return { success:true,
    hc:{ dni:dni, apellido:String(hc.Apellido), nombre:String(hc.Nombre), fnac:kinFecha_(hc.Fecha_Nac), sexo:String(hc.Sexo||''), telefono:String(hc.Telefono||''), email:String(hc.Email||''), os:String(hc.Obra_Social||''), afiliado:String(hc.Afiliado||''),
      region:String(hc.Region||''), lado:String(hc.Lado||''), diagnostico:String(hc.Diagnostico||''), fecha_lesion:kinFecha_(hc.Fecha_Lesion), operado:String(hc.Operado)==='si', medico_tipo:String(hc.Medico_Tipo||''), medico:String(hc.Medico_Nombre||''),
      kinesiologo:String(hc.Kinesiologo||''), centro:String(hc.Centro||''), centro_nombre:KIN_CENTROS_[String(hc.Centro||'')]||'', prescriptas:Number(hc.Sesiones_Prescriptas)||0, turno_pref:String(hc.Turno_Preferido||''), orden_url:String(hc.Orden_URL||''), firma_consent:String(hc.Firma_Consent||''),
      estado:String(hc.Estado||'activo'), fecha_alta:String(hc.Fecha_Alta||'').slice(0,10), tipo_alta:String(hc.Tipo_Alta||''), inscripto:String(hc.Timestamp||''), membresia:kinMembEstado_(hc), membresia_vence:String(hc.Membresia_Vence||'').slice(0,10) },
    estudios:estudios, sesiones:ses.slice().reverse(), concordancia:kinConcordancia_(ses), orden:kinOrdenObj_(kinOrdenUltima_(dni)), orden_medica:String(hc.Orden_Medica||''), notas:kinNotas_(dni) };
}
function kinResumenDe_(dni){
  const hc = kinUltimoHC_(dni);
  const nts = kinNotas_(dni), un = nts.length ? nts[0] : null;
  if(!hc){
    const o = kinOrdenAbierta_(dni); if(!o) return null;
    return { estado:'orden_enviada', orden_enviada:true, sesiones:0, prescriptas:Number(o.Sesiones)||0, medico_tipo:'equipo', medico:String(o.Medico||''), kinesiologo:KIN_KINE_DEFAULT_, tiene_orden:true, orden_fecha:String(o.Timestamp||'').slice(0,10), notas_n:nts.length, ult_nota:un };
  }
  const ses = kinSesionesEpisodio_(dni, hc.Timestamp).map(kinSesObj_);
  const cerr = ses.filter(function(s){ return s.estado==='cerrada'; });
  const u = cerr[cerr.length-1] || null;
  return { estado:String(hc.Estado||'activo'), sesiones:cerr.length, prescriptas:Number(hc.Sesiones_Prescriptas)||0, ultima_fecha:u?u.fecha:'', sens_pac:u?u.sens_salida:null, sens_kine:u?u.sens_kine:null, conforme:u?u.conforme:'',
    concordancia:kinConcordancia_(ses).pct, medico_tipo:String(hc.Medico_Tipo||''), medico:String(hc.Medico_Nombre||''), kinesiologo:String(hc.Kinesiologo||''), tiene_orden:!!hc.Orden_URL, membresia:kinMembEstado_(hc), notas_n:nts.length, ult_nota:un, fecha_alta:String(hc.Fecha_Alta||'').slice(0,10), tipo_alta:String(hc.Tipo_Alta||'') };
}
// KINESIÓLOGO: qué se hizo + puntos de evaluación + su opinión (misma escala que la del paciente)
function kinNotaSesion(d){
  const dni = kinDni_(d.dni), fila = parseInt(d.fila);
  const sh = kinHojaSes_();
  if(!dni || !fila || fila<2 || fila>sh.getLastRow()) return {success:false, error:'Sesión inexistente'};
  const h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0], v = sh.getRange(fila,1,1,h.length).getValues()[0];
  if(String(v[h.indexOf('DNI')])!==dni) return {success:false, error:'La sesión no es de ese paciente'};
  const sk = kinNum_(d.sensacion_kine, 0, 10);
  const evo = ['mejor','igual','peor'].indexOf(String(d.evolucion))>-1 ? String(d.evolucion) : '';
  const col = function(c){ return KIN_HDR_SES_.indexOf(c)+1; };
  sh.getRange(fila, col('Hecho_Kine')).setValue(kinTexto_(d.hecho, 800));
  sh.getRange(fila, col('Puntos_Eval')).setValue(kinTexto_(d.puntos, 500));
  sh.getRange(fila, col('Sensacion_Kine')).setValue(sk==null?'':sk);
  sh.getRange(fila, col('Evolucion_Kine')).setValue(evo);
  sh.getRange(fila, col('Nota_Hora')).setValue(new Date().toISOString());
  return {success:true, message:'Evolución de la sesión guardada'};
}
// KINESIÓLOGO/MÉDICO: manda por mail el link de la portada (turno + inscripción) al paciente
function kinEnviarPortada(d){
  const email = String(d.email||'').trim();
  if(!kinMailOk_(email)) return {success:false, error:'El mail no parece válido'};
  const centro = KIN_CENTROS_[String(d.centro||'')] ? String(d.centro) : 'consultorio';
  try{
    const c = CacheService.getScriptCache(), k = 'kinpor_' + psHoy_(), n = Number(c.get(k)||0);
    if(n >= 30) return {success:false, error:'Tope diario de envíos alcanzado'};
    c.put(k, String(n+1), 86400);
  }catch(e){}
  const link = baseSimm_() + 'KIN_Portada.html?centro=' + centro;
  try{
    MailApp.sendEmail(email, 'SIM-M · Tu turno de kinesiología', 'Hola' + (d.nombre?' '+kinTexto_(d.nombre,40):'') + ',\n\nPara sacar turno e inscribirte en kinesiología (' + KIN_CENTROS_[centro] + ') entrá a este enlace desde tu celular:\n' + link +
      '\n\nLa primera vez cargás tus datos, tu lesión y una foto de la orden médica. Las siguientes veces, escaneás el mismo QR del centro para marcar tu presente y tu salida.', {replyTo:MEDICO_MAIL_KINESIOLOGIA, name:KIN_MARCA_});
  }catch(e){ return {success:false, error:'No se pudo enviar el mail'}; }
  return {success:true, link:link, message:'Link enviado a ' + email};
}

// ── BITÁCORA DEL ALTA (informe al paciente + dashboard del médico) ──
function kinBitacora_(dni, tipo, extra){
  const hc = kinUltimoHC_(dni), pac = kepPaciente_(dni);
  const ses = hc ? kinSesionesEpisodio_(dni, hc.Timestamp).map(kinSesObj_) : [];
  const cerr = ses.filter(function(s){ return s.estado==='cerrada'; });
  const avg = function(l){ const x = l.filter(function(v){ return v!=null; }); return x.length ? Math.round(x.reduce(function(a,b){ return a+b; },0)/x.length*10)/10 : null; };
  const regs = kepFilas_(kepHoja_('KEP_Evaluaciones', KEP_HDR_EVAL_)).filter(function(f){ return String(f.DNI)===dni; }).map(kepFilaAReg_).sort(function(a,b){ return new Date(a.fecha)-new Date(b.fecha); });
  const meds = kepFilas_(kepHoja_('KEP_Mediciones', KEP_HDR_MED_)).filter(function(f){ return String(f.DNI)===dni; });
  const conc = kinConcordancia_(cerr);
  const b = {
    tipo:tipo, fecha:psHoy_(), paciente: pac ? (pac.apellido + ', ' + pac.nombre) : dni,
    diagnostico: hc ? String(hc.Diagnostico||'') : (pac?pac.dx:''), region: hc ? String(hc.Region||'') : '', medico: hc ? String(hc.Medico_Nombre||'') : '', kinesiologo: hc ? String(hc.Kinesiologo||'') : '',
    desde: hc ? String(hc.Timestamp).slice(0,10) : '', sesiones_asistidas:cerr.length, sesiones_prescriptas: hc ? (Number(hc.Sesiones_Prescriptas)||0) : 0,
    sens_entrada_prom: avg(ses.map(function(s){ return s.sens_entrada; })), sens_salida_prom: avg(cerr.map(function(s){ return s.sens_salida; })), sens_kine_prom: avg(cerr.map(function(s){ return s.sens_kine; })),
    concordancia:conc,
    hecho: cerr.filter(function(s){ return s.hecho; }).map(function(s){ return {n:s.n, fecha:s.fecha, texto:s.hecho}; }),
    puntos: cerr.map(function(s){ return s.puntos; }).filter(Boolean).join(' · ').slice(0,900),
    comentarios: cerr.filter(function(s){ return s.comentario; }).map(function(s){ return {n:s.n, fecha:s.fecha, texto:s.comentario}; }),
    conformes: cerr.filter(function(s){ return s.conforme==='conforme'; }).length,
    kep_primera: regs.length ? regs[0] : null, kep_ultima: regs.length ? regs[regs.length-1] : null, kep_n: regs.length,
    mediciones: meds.length ? { primera:{fecha:String(meds[0].Timestamp).slice(0,10), rom:kepJson_(meds[0].ROM_JSON), fuerza:kepJson_(meds[0].Fuerza_JSON)}, ultima:{fecha:String(meds[meds.length-1].Timestamp).slice(0,10), rom:kepJson_(meds[meds.length-1].ROM_JSON), fuerza:kepJson_(meds[meds.length-1].Fuerza_JSON)} } : null,
    resumen_medico: kinTexto_(extra && extra.resumen, 2000), conclusion: kinTexto_(extra && extra.conclusion, 1500)
  };
  return b;
}
// ═════ Mensaje por mail + membresía de seguimiento (Polidori Kinesiología) ═════
// La membresía NO cobra nada desde el sistema: el valor y el link/alias de pago los pone el profesional en
// Propiedades del script (KIN_MEMBRESIA_PRECIO, KIN_MEMBRESIA_LINK, KIN_MEMBRESIA_ALIAS). El paciente la solicita; el
// kinesiólogo confirma el pago desde KSM y recién ahí queda "activa" (+30 días, renovable).
function kinProp_(k){ try{ return String(PropertiesService.getScriptProperties().getProperty(k)||'').trim(); }catch(e){ return ''; } }
function kinMembEstado_(hc){
  const m = String(hc && hc.Membresia || '');
  if(m==='activa'){ const v = String(hc.Membresia_Vence||'').slice(0,10); if(v && v < psHoy_()) return 'vencida'; }
  return m;
}
function kinMembInfo_(){
  return { precio:kinProp_('KIN_MEMBRESIA_PRECIO').slice(0,120), link:/^https:\/\//.test(kinProp_('KIN_MEMBRESIA_LINK')) ? kinProp_('KIN_MEMBRESIA_LINK') : '', alias:kinProp_('KIN_MEMBRESIA_ALIAS').slice(0,80), kinesiologo:KIN_KINE_DEFAULT_ };
}
// PÚBLICO (con DNI + fecha de nacimiento): el paciente le escribe a su kinesiólogo; le llega por mail y puede responderle directo
function kinMensaje(d){
  const a = kinAuth_(d.dni, d.fnac);
  if(a.error) return {success:false, error:a.error};
  const hc = a.hc, texto = kinTexto_(String(d.mensaje||'').replace(/[\r\n]+/g,'\n'), 1500);
  if(texto.length < 5) return {success:false, error:'Escribí tu mensaje.'};
  try{
    const c = CacheService.getScriptCache(), k = 'kinmsg_' + a.dni + '_' + psHoy_(), n = Number(c.get(k)||0);
    if(n >= 3) return {success:false, error:'Ya enviaste 3 mensajes hoy. Si es urgente, llamá o consultá en recepción.'};
    c.put(k, String(n+1), 86400);
  }catch(e){}
  const resp = kinMailOk_(d.email) ? String(d.email) : (kinMailOk_(String(hc.Email||'')) ? String(hc.Email) : '');
  const cuerpo = 'Mensaje de ' + String(hc.Apellido) + ', ' + String(hc.Nombre) + ' (DNI ' + a.dni + ') — ' + (KIN_CENTROS_[String(hc.Centro||'')]||'') + '\nTeléfono: ' + String(hc.Telefono||'—') + (resp ? '\nResponder a: ' + resp : '\n(No dejó mail: contestale por teléfono.)') + '\n\n' + texto;
  let kp = kinProp_('KIN_MAIL_KINE');
  const para = [MEDICO_MAIL_KINESIOLOGIA].concat(kinMailOk_(String(hc.Kinesiologo_Email||'')) ? [String(hc.Kinesiologo_Email)] : []).concat(kinMailOk_(kp) ? [kp] : []).filter(function(x,i,l){ return l.indexOf(x)===i; }).join(',');
  try{
    const op = {name:'S.I.M-M Kinesiología'}; if(resp) op.replyTo = resp;
    MailApp.sendEmail(para, 'SIM-M KIN · Mensaje de paciente — ' + String(hc.Apellido), cuerpo, op);
  }catch(e){ return {success:false, error:'No se pudo enviar. Probá de nuevo en un rato.'}; }
  return {success:true, message:'Mensaje enviado a ' + KIN_KINE_DEFAULT_ + '. Te va a responder' + (resp ? ' a ' + resp : ' por teléfono') + '.'};
}
// PÚBLICO: datos de la membresía (valor y forma de pago que cargó el profesional)
function kinMembresiaInfo(){ const i = kinMembInfo_(); return {success:true, precio:i.precio, link:i.link, alias:i.alias, kinesiologo:i.kinesiologo, configurada:!!(i.precio||i.link||i.alias)}; }
// PÚBLICO (DNI + fecha de nacimiento): el paciente solicita la membresía de seguimiento
function kinMembresiaSolicitar(d){
  const a = kinAuth_(d.dni, d.fnac);
  if(a.error) return {success:false, error:a.error};
  const hc = a.hc, est = kinMembEstado_(hc), info = kinMembInfo_();
  if(est==='activa') return {success:true, ya:true, estado:'activa', vence:String(hc.Membresia_Vence||'').slice(0,10), message:'Tu membresía ya está activa hasta el ' + String(hc.Membresia_Vence||'').slice(0,10) + '.', precio:info.precio, link:info.link, alias:info.alias};
  const col = KIN_HDR_HC_.indexOf('Membresia')+1;
  kinHojaHC_().getRange(hc._row, col).setValue('solicitada');
  kinMailKine_('Solicitud de membresía de seguimiento — ' + String(hc.Apellido), String(hc.Apellido)+', '+String(hc.Nombre)+' (DNI '+a.dni+', tel. '+String(hc.Telefono||'—')+') solicitó la membresía de seguimiento. Cuando confirmes el pago, activala desde KSM (historia clínica kinésica del paciente).', hc.Kinesiologo_Email);
  return {success:true, estado:'solicitada', message:'Listo: quedó registrada tu solicitud. ' + KIN_KINE_DEFAULT_ + ' confirma el pago y te activa la membresía.', precio:info.precio, link:info.link, alias:info.alias};
}
// KINESIÓLOGO (con token): confirma el pago → activa por N días, o da de baja
function kinMembresiaEstado(d){
  const dni = kinDni_(d.dni), hc = kinUltimoHC_(dni);
  if(!hc) return {success:false, error:'Paciente sin historia kinésica'};
  const est = ['activa','inactiva','solicitada'].indexOf(String(d.estado))>-1 ? String(d.estado) : '';
  if(!est) return {success:false, error:'Estado inválido'};
  const dias = Math.max(1, Math.min(366, parseInt(d.dias)||30));
  let vence = '';
  if(est==='activa'){
    const base = String(hc.Membresia_Vence||'').slice(0,10) > psHoy_() ? new Date(String(hc.Membresia_Vence).slice(0,10)+'T12:00:00') : new Date();
    base.setDate(base.getDate()+dias); vence = base.toISOString().slice(0,10);
  }
  const sh = kinHojaHC_();
  sh.getRange(hc._row, KIN_HDR_HC_.indexOf('Membresia')+1).setValue(est);
  sh.getRange(hc._row, KIN_HDR_HC_.indexOf('Membresia_Vence')+1).setValue(vence);
  let mail = false;
  if(est==='activa' && kinMailOk_(String(hc.Email||''))){
    try{ MailApp.sendEmail(String(hc.Email), 'SIM-M · Tu membresía de seguimiento está activa', 'Hola ' + String(hc.Nombre) + ',\n\nTu membresía de seguimiento con ' + KIN_MARCA_ + ' quedó activa hasta el ' + vence + '. Podés escribirle a tu kinesiólogo desde el QR del centro ("Escribirle a mi kinesiólogo").', {replyTo:MEDICO_MAIL_KINESIOLOGIA, name:KIN_MARCA_}); mail = true; }catch(e){}
  }
  return {success:true, estado:est, vence:vence, mailEnviado:mail, message: est==='activa' ? 'Membresía activa hasta ' + vence : 'Estado actualizado'};
}
// ═════ Orden del médico → centro de kinesiología, y HISTORIA COMPARTIDA médico ↔ kinesiólogo ═════
// 1) El médico, cuando el paciente acepta hacer la rehabilitación en este centro, envía la orden (kin_orden_enviar):
//    le llega al kinesiólogo por mail y queda en KIN_Ordenes ("enviada"); el paciente recibe el link de la portada.
// 2) Cuando el paciente se inscribe por el QR (acepta la rehabilitación), la orden se VINCULA sola a su historia
//    kinésica (prescriptor = equipo, sesiones, indicación, resumen de estudios) y se le avisa al médico.
// 3) Desde ahí hay un único hilo de historia clínica (KIN_Notas): médico y kinesiólogo escriben y leen lo mismo.
function kinHojaOrd_(){ return kepHoja_('KIN_Ordenes', KIN_HDR_ORD_); }
function kinHojaNot_(){ return kepHoja_('KIN_Notas', KIN_HDR_NOT_); }
function kinFilas_Ord_(){ return kepFilas_(kinHojaOrd_()); }
function kinOrdenAbierta_(dni){ const l = kinFilas_Ord_().filter(function(f){ return String(f.DNI)===dni && String(f.Estado)==='enviada'; }); return l.length ? l[l.length-1] : null; }
function kinOrdenUltima_(dni){ const l = kinFilas_Ord_().filter(function(f){ return String(f.DNI)===dni && String(f.Estado)!=='anulada'; }); return l.length ? l[l.length-1] : null; }
function kinOrdenObj_(o){
  if(!o) return null;
  return { fila:o._row, dni:String(o.DNI||''), fecha:String(o.Timestamp||'').slice(0,10), estado:String(o.Estado||''), paciente:String(o.Paciente||''), telefono:String(o.Telefono||''), email:String(o.Email||''), diagnostico:String(o.Diagnostico||''), region:String(o.Region||''), lado:String(o.Lado||''),
    sesiones:Number(o.Sesiones)||0, frecuencia:String(o.Frecuencia||''), indicacion:String(o.Indicacion||''), medico:String(o.Medico||''), estudios:String(o.Resumen_Estudios||''), centro:String(o.Centro||''), vinculada:String(o.Fecha_Vinculo||'').slice(0,10) };
}
function kinNotas_(dni){
  return kepFilas_(kinHojaNot_()).filter(function(f){ return String(f.DNI)===dni; }).map(function(f){ return { fecha:String(f.Timestamp||''), tipo:String(f.Autor_Tipo||''), autor:String(f.Autor||''), texto:String(f.Texto||'') }; }).reverse();
}
function kinOrdenVincular_(ord, hcTs, nombre){
  try{
    const sh = kinHojaOrd_(), c = function(n){ return KIN_HDR_ORD_.indexOf(n)+1; };
    sh.getRange(ord._row, c('Estado')).setValue('vinculada'); sh.getRange(ord._row, c('HC_Timestamp')).setValue(hcTs); sh.getRange(ord._row, c('Fecha_Vinculo')).setValue(psHoy_());
  }catch(e){}
  try{ MailApp.sendEmail(KIN_MAIL_MEDICO_, 'SIM-M KIN · Tu paciente aceptó la rehabilitación — ' + nombre, nombre + ' se inscribió en ' + KIN_MARCA_ + ' (' + KIN_KINE_DEFAULT_ + ') con tu orden. Desde ahora comparten una sola historia clínica: vos ves su evolución, sus sensaciones y la concordancia en OCM (pestaña Kinesiología), y el kinesiólogo ve tus estudios y tus notas.', {name:'S.I.M-M Kinesiología'}); }catch(e){}
}
// MÉDICO (token): envía la orden de kinesiología al centro. Exige confirmar que el paciente aceptó hacerla ahí.
function kinOrdenEnviar(d){
  const dni = kinDni_(d.dni);
  if(!dni) return {success:false, error:'Falta el DNI'};
  if(d.paciente_acepta!==true) return {success:false, error:'Confirmá que el paciente aceptó hacer la rehabilitación en este centro: sin eso la orden no se envía.'};
  const indic = kinTexto_(d.indicacion, 800);
  if(indic.length < 3) return {success:false, error:'Falta la indicación de kinesiología'};
  const pac = kepPaciente_(dni);
  if(!pac) return {success:false, error:'No encontré al paciente en la base'};
  const centro = KIN_CENTROS_[String(d.centro||'')] ? String(d.centro) : 'consultorio';
  const ses = Math.max(0, Math.min(60, parseInt(d.sesiones)||0)), frec = kinTexto_(d.frecuencia, 60), medico = kinTexto_(d.medico, 80) || 'Dr. Luis Alberto Muratori';
  const nombre = (pac.apellido + ', ' + pac.nombre).trim();
  const estudios = kinResumenEstudios_(dni);
  const hc = kinUltimoHC_(dni), yaInscripto = !!(hc && String(hc.Estado)!=='alta');
  const sh = kinHojaOrd_(), ts = new Date().toISOString();
  const fila = [ts, dni, nombre, pac.tel, pac.email, kinTexto_(d.diagnostico, 300) || pac.dx, kinTexto_(d.region, 60), kinTexto_(d.lado, 20), ses||'', frec, indic, medico, estudios, yaInscripto?'vinculada':'enviada', yaInscripto?String(hc.Timestamp):'', yaInscripto?psHoy_():'', centro, 'El médico confirma que el paciente aceptó realizar la rehabilitación en este centro ' + ts];
  const previa = kinOrdenAbierta_(dni);
  if(previa) sh.getRange(previa._row, 1, 1, fila.length).setValues([fila]); else sh.appendRow(fila);
  if(yaInscripto){
    const sc = kinHojaHC_(), c = function(n){ return KIN_HDR_HC_.indexOf(n)+1; };
    sc.getRange(hc._row, c('Orden_Medica')).setValue(indic); sc.getRange(hc._row, c('Medico_Tipo')).setValue('equipo'); sc.getRange(hc._row, c('Medico_Nombre')).setValue(medico);
    if(ses && !Number(hc.Sesiones_Prescriptas)) sc.getRange(hc._row, c('Sesiones_Prescriptas')).setValue(ses);
    sc.getRange(hc._row, c('Resumen_Estudios')).setValue(estudios);
  }
  const link = baseSimm_() + 'KIN_Portada.html?centro=' + centro;
  kinMailKine_('Orden de kinesiología — ' + nombre, 'El ' + medico + ' te envía la orden de kinesiología de ' + nombre + ' (DNI ' + dni + '), que aceptó rehabilitarse en ' + KIN_CENTROS_[centro] + '.\n\nTeléfono: ' + (pac.tel||'—') + (pac.email?'\nMail: '+pac.email:'') + '\nDiagnóstico: ' + (fila[5]||'—') + '\nIndicación: ' + indic + (ses?'\nSesiones: ' + ses + (frec?' × '+frec:''):'') + (estudios?'\n\nResumen de estudios:\n'+estudios:'') +
    '\n\n' + (yaInscripto ? 'El paciente ya estaba inscripto: la orden quedó vinculada a su historia kinésica.' : 'Cuando el paciente se contacte y se inscriba por el QR, la orden se vincula sola a su historia (una sola historia compartida con el médico).') + '\nHistoria compartida: KSM → paciente → "Historia compartida".', '');
  let mailPac = false;
  if(kinMailOk_(pac.email) && !yaInscripto){
    try{ MailApp.sendEmail(pac.email, 'SIM-M · Tu orden de kinesiología llegó a ' + KIN_MARCA_, 'Hola ' + pac.nombre + ',\n\nTu orden de kinesiología ya está en ' + KIN_MARCA_ + ' (' + KIN_KINE_DEFAULT_ + '). Para sacar turno e inscribirte entrá desde tu celular:\n' + link + '\n\nLa inscripción es rápida: tus datos, tu lesión y la firma. Así tu kinesiólogo y tu médico trabajan con una sola historia clínica. Si preferís hacer la kinesiología en otro lugar, avisanos.', {replyTo:KIN_MAIL_MEDICO_, name:'Dr. Luis Alberto Muratori'}); mailPac = true; }catch(e){}
  }
  return {success:true, estado: yaInscripto ? 'vinculada' : 'enviada', mailEnviado:mailPac, link:link, message: yaInscripto ? 'El paciente ya estaba inscripto: la orden quedó vinculada a su historia kinésica.' : 'Orden enviada a ' + KIN_MARCA_ + '.' + (mailPac ? ' Al paciente le llegó el link para inscribirse.' : '') + ' Se vincula sola cuando el paciente se inscriba.'};
}
// KINESIÓLOGO (token): órdenes que el médico envió y el paciente todavía no se inscribió
function getKinOrdenes(){
  const l = kinFilas_Ord_().filter(function(f){ return String(f.Estado)==='enviada'; }).map(kinOrdenObj_).reverse();
  return {success:true, ordenes:l};
}
// MÉDICO o KINESIÓLOGO (token): nota en la historia compartida
function kinNotaCompartida(d){
  const dni = kinDni_(d.dni), tipo = String(d.autor_tipo)==='medico' ? 'medico' : (String(d.autor_tipo)==='kine' ? 'kine' : '');
  const texto = kinTexto_(String(d.texto||'').replace(/\r/g,''), 1000);
  if(!dni || !tipo) return {success:false, error:'Faltan datos'};
  if(texto.length < 3) return {success:false, error:'Escribí la nota'};
  if(!kinUltimoHC_(dni) && !kinOrdenUltima_(dni)) return {success:false, error:'Este paciente todavía no tiene historia kinésica ni orden enviada'};
  const autor = kinTexto_(d.autor, 80) || (tipo==='medico' ? 'Dr. Muratori' : 'Kinesiología');
  kinHojaNot_().appendRow([new Date().toISOString(), dni, tipo, autor, texto]);
  const pac = kepPaciente_(dni), nom = pac ? (pac.apellido + ', ' + pac.nombre) : dni;
  try{
    if(tipo==='medico') kinMailKine_('Nota del médico — ' + nom, autor + ' escribió en la historia compartida de ' + nom + ' (DNI ' + dni + '):\n\n' + texto, '');
    else MailApp.sendEmail(KIN_MAIL_MEDICO_, 'SIM-M KIN · Nota del kinesiólogo — ' + nom, autor + ' escribió en la historia compartida de ' + nom + ' (DNI ' + dni + '):\n\n' + texto + '\n\nLa ves en OCM → Seguimiento → Kinesiología.', {name:'S.I.M-M Kinesiología'});
  }catch(e){}
  return {success:true, message:'Nota agregada a la historia compartida'};
}
function kinBitacoraHtml_(b){
  const e = kinEsc_, tit = {kinesica:'Alta de kinesiología', medica:'Alta médica', quirurgica:'Alta quirúrgica'}[b.tipo] || 'Alta';
  const f = function(v){ return v==null ? '—' : v; };
  let h = '<div style="font-family:Arial,sans-serif;max-width:640px;margin:auto;color:#0f172a"><h2 style="color:#0369a1">' + e(tit) + ' — bitácora de tu tratamiento</h2>' +
    '<p><b>' + e(b.paciente) + '</b> · ' + e(b.fecha) + '<br>' + (b.diagnostico ? 'Motivo / diagnóstico: ' + e(b.diagnostico) + (b.region?' ('+e(b.region)+')':'') + '<br>' : '') + (b.medico ? 'Médico prescriptor: ' + e(b.medico) + '<br>' : '') + (b.kinesiologo ? 'Kinesiólogo/a actuante: ' + e(b.kinesiologo) : '') + '</p>';
  if(b.conclusion) h += '<h3>Conclusión</h3><p>' + e(b.conclusion) + '</p>';
  if(b.resumen_medico) h += '<h3>Resumen médico</h3><p>' + e(b.resumen_medico).replace(/\n/g,'<br>') + '</p>';
  if(b.sesiones_asistidas || b.sesiones_prescriptas) h += '<h3>Asistencia</h3><p>Sesiones asistidas: <b>' + b.sesiones_asistidas + '</b>' + (b.sesiones_prescriptas ? ' de ' + b.sesiones_prescriptas + ' prescriptas' : '') + (b.desde ? ' · desde el ' + e(b.desde) : '') + '.</p>';
  if(b.sens_salida_prom!=null){
    h += '<h3>Cómo te sentiste</h3><p>Sensación promedio al llegar: <b>' + f(b.sens_entrada_prom) + '</b> · al salir: <b>' + f(b.sens_salida_prom) + '</b> (0 = muy mal, 10 = muy bien). Conformidad: ' + b.conformes + ' de ' + b.sesiones_asistidas + ' sesiones.</p>';
    if(b.concordancia.n) h += '<h3>Concordancia entre tu opinión y la del terapeuta</h3><p>Coincidieron en <b>' + b.concordancia.pct + '%</b> de las sesiones evaluadas (' + b.concordancia.concuerdan + ' coinciden, ' + b.concordancia.parcial + ' parcialmente, ' + b.concordancia.discrepan + ' difieren). Promedio del terapeuta: ' + f(b.sens_kine_prom) + '. Es un criterio práctico, no un puntaje clínico.</p>';
  }
  if(b.hecho.length) h += '<h3>Lo que se hizo</h3><ul>' + b.hecho.slice(0,20).map(function(x){ return '<li>Sesión ' + x.n + ' (' + e(x.fecha) + '): ' + e(x.texto) + '</li>'; }).join('') + '</ul>';
  if(b.puntos) h += '<h3>Puntos de evaluación tenidos en cuenta</h3><p>' + e(b.puntos) + '</p>';
  if(b.kep_ultima){
    const p = b.kep_primera, u = b.kep_ultima;
    h += '<h3>Tus evaluaciones semanales (' + b.kep_n + ')</h3><p>Dolor en movimiento: ' + f(p.eva_mov) + ' → <b>' + f(u.eva_mov) + '</b> · Función (PSFS): ' + f(p.psfs) + ' → <b>' + f(u.psfs) + '</b> · Cumplimiento del plan: ' + f(u.cump) + '/10.</p>';
  }
  if(b.mediciones){
    const rm = function(m){ return (m.rom||[]).map(function(x){ return e(x.mov) + ' ' + e(x.lado) + ' ' + e(x.grados) + '°'; }).concat((m.fuerza||[]).map(function(x){ return e(x.musculo) + ' ' + e(x.lado) + ' ' + e(x.mrc) + '/5'; })).join('; ') || '—'; };
    h += '<h3>Mediciones del kinesiólogo</h3><p>Inicial (' + e(b.mediciones.primera.fecha) + '): ' + rm(b.mediciones.primera) + '<br>Final (' + e(b.mediciones.ultima.fecha) + '): ' + rm(b.mediciones.ultima) + '</p>';
  }
  if(b.comentarios.length) h += '<h3>Tus comentarios</h3><ul>' + b.comentarios.slice(0,15).map(function(x){ return '<li>Sesión ' + x.n + ': ' + e(x.texto) + '</li>'; }).join('') + '</ul>';
  return h + '<hr><p style="font-size:12px;color:#64748b">Este informe resume tu tratamiento y no reemplaza tu historia clínica. Con el alta, tu código de acceso a la página de salud se dio de baja. Consultorio Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943.</p></div>';
}
// Da de baja los códigos de la página de salud del paciente (alta): en el sitio y en la planilla
function psRevocarTodos_(dni, prof){
  const sh = psHoja_(); let n = 0, sitioFalla = 0;
  psFilas_().filter(function(x){ return x.dni===dni && x.codigo && x.estado!=='revocado'; }).forEach(function(x){
    const r = psApi_('revocar', {codigo:x.codigo});
    if(!r.ok) sitioFalla++;
    psActualizarFila_(sh, x.fila, {Estado:'revocado'});
    psGuardarHist_(sh, x.fila, {accion:'revocar', sitio:r.ok, profesional:prof});
    n++;
  });
  return {revocados:n, sitioFalla:sitioFalla};
}
// KINESIÓLOGO o MÉDICO: alta kinésica / médica / quirúrgica → bitácora + baja del código
function altaPaciente(d){
  const dni = kinDni_(d.dni), tipo = ['kinesica','medica','quirurgica'].indexOf(String(d.tipo))>-1 ? String(d.tipo) : '';
  if(!dni || !tipo) return {success:false, error:'Falta el DNI o el tipo de alta'};
  const hc = kinUltimoHC_(dni);
  if(tipo==='kinesica' && (!hc || String(hc.Estado)==='alta')) return {success:false, error:'El paciente no tiene un tratamiento kinésico activo'};
  const prof = kinTexto_(d.profesional, 80);
  const b = kinBitacora_(dni, tipo, {resumen:d.resumen, conclusion:d.conclusion});
  const html = kinBitacoraHtml_(b);
  if(hc && (tipo==='kinesica' || String(hc.Estado)!=='alta')){
    const sh = kinHojaHC_(), c = function(n){ return KIN_HDR_HC_.indexOf(n)+1; };
    if(tipo==='kinesica' || d.cerrarKine===true){
      sh.getRange(hc._row, c('Estado')).setValue('alta'); sh.getRange(hc._row, c('Fecha_Alta')).setValue(psHoy_()); sh.getRange(hc._row, c('Tipo_Alta')).setValue(tipo);
      try{ sh.getRange(hc._row, c('Bitacora_JSON')).setValue(JSON.stringify(b).slice(0,45000)); }catch(e){}
    }
  }
  try{
    const sg = kepHoja_('KEP_Seguimientos', KEP_HDR_SEG_), fs = kepFilas_(sg).filter(function(f){ return String(f.DNI)===dni; });
    fs.forEach(function(f){ sg.getRange(f._row, KEP_HDR_SEG_.indexOf('Activo')+1).setValue('no'); });
  }catch(e){}
  const rev = psRevocarTodos_(dni, prof);
  const pac = kepPaciente_(dni);
  const para = (pac && kinMailOk_(pac.email)) ? pac.email : '';
  let mailPac = false;
  const asunto = 'SIM-M · ' + ({kinesica:'Alta de kinesiología', medica:'Alta médica', quirurgica:'Alta quirúrgica'}[tipo]) + ' — bitácora de tu tratamiento';
  if(para){ try{ MailApp.sendEmail(para, asunto, 'Adjuntamos la bitácora de tu tratamiento (abrila en un cliente de mail que muestre formato).', {htmlBody:html, replyTo:MEDICO_MAIL_KINESIOLOGIA, name:'Consultorio Dr. Luis Alberto Muratori'}); mailPac = true; }catch(e){} }
  try{ MailApp.sendEmail(MEDICO_MAIL_KINESIOLOGIA + (hc && kinMailOk_(hc.Kinesiologo_Email) ? ',' + hc.Kinesiologo_Email : ''), 'SIM-M KIN · Alta registrada — ' + b.paciente, 'Alta ' + tipo + ' registrada por ' + (prof||'—') + '. Se dieron de baja ' + rev.revocados + ' código(s) de la página de salud' + (rev.sitioFalla ? ' (en ' + rev.sitioFalla + ' no se pudo avisar al sitio: revocalos desde el panel de la página)' : '') + '.\n\n(Versión con formato en el mail del paciente.)', {htmlBody:html, name:'S.I.M-M Kinesiología'}); }catch(e){}
  return {success:true, message:'Alta registrada. ' + (mailPac ? 'Se envió la bitácora al paciente. ' : 'El paciente no tiene mail cargado: la bitácora quedó en el seguimiento. ') + (rev.revocados ? 'Código(s) de la página de salud dados de baja: ' + rev.revocados + (rev.sitioFalla ? ' (revisá el sitio)' : '') + '.' : ''), bitacora:b, revocados:rev.revocados, sitioFalla:rev.sitioFalla, mailEnviado:mailPac};
}
// MÉDICO: consulta mensual → los códigos siguen activos (se suman días a todos los vigentes)
function renovarAccesosPaciente(d){
  const dni = kinDni_(d.dni); let dias = parseInt(d.dias); if(!(dias>=1 && dias<=365)) dias = 30;
  if(!dni) return {success:false, error:'Falta el DNI'};
  const sh = psHoja_(); let n = 0, fallas = 0, vence = '';
  psFilas_().filter(function(x){ return x.dni===dni && x.codigo && x.estado!=='revocado'; }).forEach(function(x){
    const r = psApi_('extender', {codigo:x.codigo, dias:dias});
    if(!r.ok){ fallas++; return; }
    vence = String(r.data.vence||'').slice(0,10);
    psActualizarFila_(sh, x.fila, {Vence:vence, Estado:'activo', Plan:'Consulta mensual +' + dias + ' días'});
    psGuardarHist_(sh, x.fila, {accion:'extender', dias:dias, vence:vence, profesional:String(d.profesional||'')});
    n++;
  });
  if(!n && !fallas) return {success:false, error:'El paciente no tiene códigos activos'};
  return {success:fallas===0, renovados:n, fallas:fallas, vence:vence, message: n ? ('Códigos renovados (' + n + ') hasta el ' + vence) : 'No se pudo renovar', error: fallas ? 'No se pudo renovar en el sitio' : undefined};
}
