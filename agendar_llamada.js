function agendarLlamada(e) {
  const hoja = e.source.getActiveSheet();
  const fila = e.range.getRow();
  const columna = e.range.getColumn();
  
  // Se activa con datos (1-7), Observaciones (13/M), Fecha (14/N) u Hora (15/O)
  const columnasInteres = [1, 2, 3, 4, 5, 6, 7, 13, 14, 15];
  if (!columnasInteres.includes(columna) || fila <= 1) return;

  // Leemos un rango amplio (A hasta P) para no fallar
  const rangoCompleto = hoja.getRange(fila, 1, 1, 16).getValues()[0];
  
  const cliente = rangoCompleto[0];      // Columna A
  const tel1 = rangoCompleto[2];         // Columna C
  const tel2 = rangoCompleto[3];         // Columna D
  const correo1 = rangoCompleto[4];      // Columna E
  const correo2 = rangoCompleto[5];      // Columna F
  const abonado = rangoCompleto[6];      // Columna G
  const observaciones = rangoCompleto[12]; // Columna M (Posición 12)
  const fechaPura = rangoCompleto[13];    // Columna N (Posición 13)
  const horaPura = rangoCompleto[14];     // Columna O (Posición 14)
  const eventId = rangoCompleto[15];      // Columna P (Posición 15)
  
  const calendario = CalendarApp.getDefaultCalendar();

  // Borrado de evento si falta fecha u hora
  if ((columna == 14 || columna == 15) && (!fechaPura || !horaPura)) {
    if (eventId) {
      try {
        const ev = calendario.getEventById(eventId);
        if (ev) ev.deleteEvent();
        hoja.getRange(fila, 16).clearContent();
      } catch (err) {}
    }
    return;
  }

  if (!(fechaPura instanceof Date) || !(horaPura instanceof Date)) return;
  
  const fechaFinal = new Date(
    fechaPura.getFullYear(), fechaPura.getMonth(), fechaPura.getDate(),
    horaPura.getHours(), horaPura.getMinutes(), horaPura.getSeconds()
  );

  const titulo = "📞 Llamar a: " + (cliente || "Cliente") + " (AB: " + (abonado || "S/N") + ")";
  
  // Descripción formateada
  const descripcion = "📝 NOTAS: " + (observaciones || "Sin observaciones") + 
                      "\n\n📱 Tel: " + tel1 + " / " + tel2 + 
                      "\n📧 Correo: " + correo1 + " / " + correo2;
                      
  const fechaFin = new Date(fechaFinal.getTime() + (20 * 60 * 1000));

  try {
    let evento;
    if (eventId) {
      evento = calendario.getEventById(eventId);
      if (evento) {
        evento.setTitle(titulo).setDescription(descripcion).setTime(fechaFinal, fechaFin);
      } else {
        evento = calendario.createEvent(titulo, fechaFinal, fechaFin, {description: descripcion});
      }
    } else {
      evento = calendario.createEvent(titulo, fechaFinal, fechaFin, {description: descripcion});
    }

    // Configura recordatorio de 10 minutos
    evento.removeAllReminders();
    evento.addPopupReminder(10);
    
    // Guarda el ID y da feedback visual
    hoja.getRange(fila, 16).setValue(evento.getId());
    hoja.getRange(fila, 14, 1, 2).setBackground("#d9ead3"); 
  } catch (error) {
    console.error("Error: " + error.toString());
  }
}