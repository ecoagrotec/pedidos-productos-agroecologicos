function onOpen() { 

  // agregar la funci�n al men�
  var libro = SpreadsheetApp.getActiveSpreadsheet(); 
  var menu = [{name: "Procesar pedidos", functionName: "procesar_pedidos"},{name: "M�s info...", functionName: "mas_info"}]; 
   libro.addMenu("Pedidos", menu); 

}

function mas_info() {
  var ui = SpreadsheetApp.getUi();
  var mensaje = ui.alert(
      'Pedidos Che Verde',
      'https://github.com/ecoagrotec/pedidos-productos-agroecologicos',
      ui.ButtonSet.OK);
}

function procesar_pedidos() {
  
  // Leer par�metros de entrada de la hoja 'Configuraci�n'
  var libro = SpreadsheetApp.getActiveSpreadsheet()
  var hojaConfiguracion = libro.getSheetByName("Configuraci�n");
  var nombreHojaPedidos = hojaConfiguracion.getRange(1,2,1,1).getValue().toString();
  var rotuloLugarEntrega = hojaConfiguracion.getRange(2,2,1,1).getValue().toString();
  var rotuloOrden = hojaConfiguracion.getRange(3,2,1,1).getValue().toString();
  var nombrePlanillas = hojaConfiguracion.getRange(4,2,1,1).getValue().toString();
  var totales = [];   // buscar estas palabras clave en los t�tulos y reportar los totales en la planilla Resumen
  var i=0;
  do {
    var total=hojaConfiguracion.getRange(5,i+2,1,1).getValue().toString();
    if (total!="") totales[i]=total;
    i++;
  } while (total!="")
  
  // Datos de la hoja con los datos de pedidos
  var hojaPedidos = libro.getSheetByName(nombreHojaPedidos);
  var cantidadColumnas = hojaPedidos.getLastColumn();
  var cantidadFilas = hojaPedidos.getLastRow();

  // Buscar cu�les son las columnas correspondientes a lugar de entrega (rotuloLugarEntrega) y criterio para ordenar (rotuloOrden)
  var columnaLugarEntrega = -1;
  var columnaOrden = -1;
  for (i=1;i<=cantidadColumnas;++i) 
  {
    var celda = hojaPedidos.getRange(1,i,1,1).getValues();  // fila 1, columna i, rango 1x1
    if (celda==rotuloLugarEntrega) columnaLugarEntrega = i;
    if (celda==rotuloOrden) columnaOrden = i;
  }
  if (columnaLugarEntrega==-1) throw new Error("No hay columna con el r�tulo " + rotuloLugarEntrega);   // marcar error si no se encuentra la columna
  if (columnaOrden==-1) throw new Error("No hay columna con el r�tulo " + rotuloOrden);   // marcar error si no se encuentra la columna

  // Crear una nueva columna con el lugar de entrega en formato num�rico (transformar "12) Lugar x" por "12")
  var columnaNumeroEntrega = cantidadColumnas+1;
  hojaPedidos.getRange(2,columnaNumeroEntrega,cantidadFilas-1,1).setFormula('=value(split('+hojaPedidos.getRange(2,columnaLugarEntrega,1,1).getA1Notation()+';")"))');
  hojaPedidos.getRange(1,columnaNumeroEntrega,1,1).setFormula('=max('+hojaPedidos.getRange(2,columnaNumeroEntrega,cantidadFilas-1,1).getA1Notation()+')');
  var cantidadLugaresEntrega = +hojaPedidos.getRange(1,columnaNumeroEntrega,1,1).getValue();
  
  // Ordenar planilla de pedidos (para que luego queden ordenadas las dem�s)
  hojaPedidos.getRange(2,1,cantidadFilas-1,columnaNumeroEntrega).sort(columnaOrden);
  hojaPedidos.getRange(2,1,cantidadFilas-1,columnaNumeroEntrega).sort(columnaNumeroEntrega);

  // Crear hoja resumen
  var hojaResumen = libro.getSheetByName("Resumen"); 
  if (!hojaResumen) {    // crear la hoja si no existe
    libro.insertSheet("Resumen");   
    var hojaResumen = libro.getSheetByName("Resumen"); 
  }
  hojaResumen.clear();
  libro.setActiveSheet(hojaResumen);

  // Borrar hojas de lugar de entrega existentes
  var numeroLugarEntrega=1;
  do {
    var nombreHojaLugarEntrega = nombrePlanillas+numeroLugarEntrega;
    var hojaLugarEntrega = libro.getSheetByName(nombreHojaLugarEntrega); 
    if (hojaLugarEntrega) {    // borrar la hoja si existe
      libro.deleteSheet(hojaLugarEntrega);   
    }
    numeroLugarEntrega++;
  } while (hojaLugarEntrega) // continuar mientras haya hojas
  
  // Crear nuevas hojas para cada lugar de entrega
  for (var i=1;i<=cantidadLugaresEntrega;++i) 
  {
    var numeroLugarEntrega = i;
    var nombreHojaLugarEntrega = nombrePlanillas+numeroLugarEntrega;
    libro.insertSheet(nombreHojaLugarEntrega);   
    var hojaLugarEntrega = libro.getSheetByName(nombreHojaLugarEntrega); 
    var datosPrimeraFila = hojaPedidos.getRange(1,1,1,cantidadColumnas).getValues();     // t�tulos primera fila
    hojaLugarEntrega.getRange(1,1,1,cantidadColumnas).setValues(datosPrimeraFila);   // copiar a primera fila de cada nueva hoja
  }  
  
  // Leer fila por fila en la hoja de pedidos y copiar cada una a distintas hojas seg�n el lugar de entrega
  for (i=2;i<=cantidadFilas;++i) 
  {
    numeroLugarEntrega = hojaPedidos.getRange(i,columnaNumeroEntrega,1,1).getValues();   // leer n�mero de lugar de entrega de la �ltima columna
    nombreHojaLugarEntrega = nombrePlanillas+numeroLugarEntrega;
    hojaLugarEntrega = libro.getSheetByName(nombreHojaLugarEntrega); 
    var ultimaFila = hojaLugarEntrega.getLastRow();

    var datosFilaCompleta = hojaPedidos.getRange(i,1,1,cantidadColumnas).getValues();
    hojaLugarEntrega.getRange(ultimaFila+1,1,1,cantidadColumnas).setValues(datosFilaCompleta);
  }

  // Buscar cu�les son las columnas para las que hay que reportar totales
  var cantidadTotales = totales.length;
  var columnasTotales = new Array(cantidadTotales);
  for (var i=1;i<=cantidadColumnas;++i) 
  {
    var celda = hojaPedidos.getRange(1,i,1,1).getValues();
    for (var j=0;j<cantidadTotales;++j)
    {
      if (celda.toString().toLowerCase().includes(totales[j].toLowerCase())) columnasTotales[j] = i;
    }
  }
  
  // Agregar t�tulos a planilla resumen
  for (var j=0;j<cantidadTotales;++j)
  {
    hojaResumen.getRange(1,j+2,1,1).setValue(totales[j]);
  }

  // Agregar totales a cada planilla y a planilla resumen
  for (var i=1;i<=cantidadLugaresEntrega;++i) 
  {
    numeroLugarEntrega = i;
    nombreHojaLugarEntrega = nombrePlanillas+numeroLugarEntrega;
    hojaLugarEntrega = libro.getSheetByName(nombreHojaLugarEntrega); 
    var ultimaFila = hojaLugarEntrega.getLastRow();
    
    // agregar totales
    for (j=0;j<cantidadTotales;++j)
    {
      if (ultimaFila>2) {   // si hay datos incluir f�rmula para totales
        hojaLugarEntrega.getRange(ultimaFila+1,columnasTotales[j],1,1).setFormula('=sum('+hojaLugarEntrega.getRange(2,columnasTotales[j],ultimaFila-1,1).getA1Notation()+')');
      }
      else {   // si no hay datos incluir ceros
        hojaLugarEntrega.getRange(ultimaFila+1,columnasTotales[j],1,1).setValue(0);
      }
      var totalCalculadoFormula = hojaLugarEntrega.getRange(ultimaFila+1,columnasTotales[j],1,1).getValue();   // leer valor calculado por la planilla
      hojaResumen.getRange(i+1,j+2,1,1).setValue(totalCalculadoFormula);   // copiar ese valor a la hoja resumen
    }
    
    // poner lugar de entrega en la primera columna de la hoja resumen
    if (ultimaFila>2) {   // si hay datos incluir el nombre completo del lugar de entrega
      var tituloLugarEntrega = nombreHojaLugarEntrega+" - "+hojaLugarEntrega.getRange(ultimaFila,columnaLugarEntrega,1,1).getValue();
    }
    else {   // si no hay datos s�lo el n�mero
      var tituloLugarEntrega = nombreHojaLugarEntrega;
    }
    hojaResumen.getRange(i+1,1,1,1).setValue(tituloLugarEntrega);  
  }  
  
  // agregar f�rmula para calcular totales generales de la planilla resumen
  hojaResumen.getRange(cantidadLugaresEntrega+2,1,1,1).setValue('Total hoja resumen');
  for (j=0;j<cantidadTotales;++j) {
    hojaResumen.getRange(cantidadLugaresEntrega+2,j+2,1,1).setFormula('=sum('+hojaResumen.getRange(2,j+2,cantidadLugaresEntrega,1).getA1Notation()+')');
  }
  
  // agregar totales de la planilla de pedidos
  hojaResumen.getRange(cantidadLugaresEntrega+3,1,1,1).setValue('Total planilla de Pedidos');
  ultimaFila = hojaPedidos.getLastRow();
  for (j=0;j<cantidadTotales;++j) {
    hojaPedidos.getRange(ultimaFila+1,columnasTotales[j],1,1).setFormula('=sum('+hojaPedidos.getRange(2,columnasTotales[j],ultimaFila-1,1).getA1Notation()+')');
    totalCalculadoFormula = hojaPedidos.getRange(ultimaFila+1,columnasTotales[j],1,1).getValue();   // leer valor calculado por la planilla
    hojaResumen.getRange(cantidadLugaresEntrega+3,j+2,1,1).setValue(totalCalculadoFormula);   // copiar ese valor a la hoja Resumen
    hojaPedidos.getRange(ultimaFila+1,columnasTotales[j],1,1).clear();   // borrar valor calculado por la planilla
  }
  
  // Borrar columna con el lugar de entrega en formato num�rico de la hoja de pedidos
  hojaPedidos.getRange(1,columnaNumeroEntrega,cantidadFilas,1).clear();
  
  // Mostrar la hoja resumen al final
  libro.setActiveSheet(hojaResumen);
  
}

