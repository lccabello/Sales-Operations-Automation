function generarPropuestaParaWhatsapp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaContactos = ss.getSheetByName("CONTACTOS DIARIOS");
  const hojaValor = ss.getSheetByName("Valor");
  const ui = SpreadsheetApp.getUi();
  
  const celdaActiva = hojaContactos.getActiveCell();
  const filaActiva = celdaActiva.getRow();
  
  if (filaActiva < 2) return;

  // 1. Extraer datos (Rango A hasta Q)
  const rangoDatos = hojaContactos.getRange(filaActiva, 1, 1, 17).getValues()[0];

  const cliente   = rangoDatos[0];
  const cuit      = rangoDatos[1];
  const tel1      = rangoDatos[2];
  const tel2      = rangoDatos[3];
  const mail1     = rangoDatos[4];
  const mail2     = rangoDatos[5];
  const abonado   = rangoDatos[6];
  const direccionOriginal = rangoDatos[16];

  // 2. Formatear dirección (Capitalizar primera letra de cada palabra)
  const direccionMin = direccionOriginal.toString().toLowerCase()
                         .replace(/(^\w|\s\w)/g, m => m.toUpperCase());

  // 3. Preparar los textos de contacto (SOLO DATOS LIMPIOS)
  const telsString = [tel1, tel2].filter(String).join(" // ") || "S/D";
  const mailsString = [mail1, mail2].filter(String).join(" // ") || "S/D";

  // 4. Pegar en hoja Valor (Para la visual)
  if (hojaValor) {
    hojaValor.getRange("C11").setValue(abonado);
    hojaValor.getRange("C12").setValue(cliente);
    hojaValor.getRange("C13").setValue(cuit);
    hojaValor.getRange("C14").setValue(direccionMin);
    
    // Pegar teléfonos y correos
    hojaValor.getRange("C15").setValue(telsString);
    hojaValor.getRange("C16").setValue(mailsString);
  }

  // --- SELECCIÓN DE ESCENARIO ---
  
  const respuesta = ui.alert(
    "Selecciona el tipo de Propuesta B2B/B2G",
    "¿Es un Cambio de Tecnología (Infraestructura)?\n\n" +
    "SÍ = Cambio de Tecnología (Cableado/Equipo obsoleto)\n" +
    "NO = Upgrade FTTH (Mejorar plan/velocidad corporativa)",
    ui.ButtonSet.YES_NO_CANCEL
  );

  if (respuesta == ui.Button.CANCEL) return;

  let mensaje = "";

  // OPCIÓN A: CAMBIO DE TECNOLOGÍA (SÍ) -> Tono suave: "Elegible" / "Bonificada"
  if (respuesta == ui.Button.YES) {
    mensaje = "Canal 4 - Soluciones Corporativas\n\n" +
    "Estimado/a, en relación al servicio N° " + abonado + " perteneciente a " + cliente + ", ubicado en " + direccionMin + ".\n\n" +
    "Hemos identificado que su infraestructura actual es elegible para una actualización tecnológica. Esto nos permitiría elevar el estándar de su conexión y prevenir inconvenientes técnicos a futuro.\n\n" +
    "Al tener esta calificación prioritaria, podemos migrar su servicio a nuestra red de Fibra Óptica de forma 100% bonificada (sin costo de instalación/adecuación).\n\n" +
    "Si le interesa aprovechar esta mejora técnica, quedo a la espera de su confirmación para coordinar.\n\n" +
    "Atte. Lucas Cabello | 3885027689 | LCABELLO@elcuatro.com";
  } 
  
  // OPCIÓN B: UPGRADE / PLAN SUPERIOR (NO) -> Tono Productividad + Dirección incluida
  else if (respuesta == ui.Button.NO) {
    mensaje = "Canal 4 - Soluciones Corporativas\n\n" +
    "Estimado/a, analizando el consumo y demanda del servicio N° " + abonado + " (" + cliente + ") ubicado en " + direccionMin + ".\n\n" +
    "Vemos una oportunidad para optimizar la conectividad de su organización. Actualmente contamos con planes de Fibra Óptica con mayor ancho de banda simétrico, diseñados para soportar alta carga de trabajo, transferencia de datos y uso de sistemas en la nube.\n\n" +
    "Le propongo evaluar un upgrade a un plan superior para mejorar la eficiencia operativa. ¿Le gustaría recibir las opciones vigentes?\n\n" +
    "Atte. Lucas Cabello | 3885027689 | LCABELLO@elcuatro.com";
  }

  // 6. Alerta final para copiar
  ui.alert("Copiá el siguiente mensaje:", mensaje, ui.ButtonSet.OK);

  // 7. Visualización
  ss.setActiveSheet(hojaValor);
  hojaValor.getRange("A3:G16").activate(); 
}