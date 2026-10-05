// =============================================================
// MENÚ PRINCIPAL
// =============================================================

function onOpen() {

  const ui = SpreadsheetApp.getUi();

  // Menú principal para generar documentos
  ui.createMenu('📄 Generar PDF')
    .addItem('Generar carta Acreditación', 'generarCartaPDF')
    .addItem('Código de Ética', 'abrirMenuCodigoEtica')
    .addSeparator()
    .addItem('Buscar logo de organización', 'probarBusquedaLogo')
    .addToUi();

  // Menú independiente para el manual
  ui.createMenu('📘 Manual de llenado')
    .addItem('Abrir manual', 'mostrarManualLlenado')
    .addToUi();
}


// =============================================================
// ACTUALIZAR LOGO AUTOMÁTICAMENTE AL CAMBIAR B2
// =============================================================

function onEdit(e) {

  if (!e || !e.range) return;

  const hoja = e.range.getSheet();

  if (hoja.getName() === 'Logos') return;

  if (e.range.getA1Notation() !== 'B2') return;

  const organizacion = hoja
    .getRange('B2')
    .getDisplayValue()
    .trim();

  if (!organizacion) {
    hoja.getRange('J1').clearContent();
    return;
  }

  buscarYColocarLogo(hoja, organizacion);
}


// =============================================================
// GENERAR CARTA PDF
// =============================================================

function generarCartaPDF() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  // ===========================================================
  // LOGO FIJO TEC
  // ===========================================================

  const ID_LOGO_IZQUIERDO =
    '16UKagcpT4pUaE8F0TLl36oIxrBfPDgl9';

  // ===========================================================
  // DATOS GENERALES
  // ===========================================================

  const asunto = hoja
    .getRange('B1')
    .getDisplayValue()
    .trim();

  const organizacion = hoja
    .getRange('B2')
    .getDisplayValue()
    .trim();

  const tipoPeriodo = hoja
    .getRange('A3')
    .getDisplayValue()
    .trim();

  const periodo = hoja
    .getRange('B3')
    .getDisplayValue()
    .trim();

  const anio = hoja
    .getRange('C3')
    .getDisplayValue()
    .trim();

  const periodoCompleto = [
    tipoPeriodo,
    periodo,
    anio
  ].filter(Boolean).join(' ');

  // ===========================================================
  // FIRMAS
  // ===========================================================

  const firmaIzquierda = hoja
    .getRange('G1')
    .getDisplayValue()
    .trim();

  const correoFirmaIzquierda = hoja
    .getRange('H1')
    .getDisplayValue()
    .trim();

  const firmaDerecha = hoja
    .getRange('G2')
    .getDisplayValue()
    .trim();

  // ===========================================================
  // VALIDACIONES
  // ===========================================================

  if (!asunto) {
    ui.alert('Falta capturar el proyecto/asunto en B1.');
    return;
  }

  if (!organizacion) {
    ui.alert('Falta capturar la organización en B2.');
    return;
  }

  if (!tipoPeriodo) {
    ui.alert('Falta capturar el tipo de periodo en A3.');
    return;
  }

  if (!periodo) {
    ui.alert('Falta capturar el periodo en B3.');
    return;
  }

  if (!anio) {
    ui.alert('Falta capturar el año en C3.');
    return;
  }

  if (!firmaIzquierda) {
    ui.alert(
      'Falta capturar la persona principal que firma en G1.'
    );
    return;
  }

  if (!correoFirmaIzquierda) {
    ui.alert(
      'Falta capturar el correo de la persona de G1 en H1.'
    );
    return;
  }

  if (!firmaDerecha) {
    ui.alert(
      'Falta capturar la segunda persona que firma en G2.'
    );
    return;
  }

  // ===========================================================
  // BUSCAR LOGO DE ORGANIZACIÓN
  // ===========================================================

  const logoEncontrado =
    buscarYColocarLogo(hoja, organizacion);

  if (!logoEncontrado) {

    ui.alert(
      'No encontré el logo de la organización:\n\n' +
      organizacion +
      '\n\n' +
      'Revisa la pestaña "Logos".\n\n' +
      'Columna A = Organización\n' +
      'Columna B = Liga del logo'
    );

    return;
  }

  SpreadsheetApp.flush();

  const ligaLogoDerecho = hoja
    .getRange('J1')
    .getDisplayValue()
    .trim();

  // ===========================================================
  // OBTENER LOGO TEC
  // ===========================================================

  let logoIzquierdo;

  try {

    logoIzquierdo = DriveApp
      .getFileById(ID_LOGO_IZQUIERDO)
      .getBlob();

  } catch (error) {

    ui.alert(
      'No se pudo acceder al logo fijo del Tec.\n\n' +
      'Verifica que tu cuenta tenga acceso al archivo.'
    );

    return;
  }

  // ===========================================================
  // OBTENER LOGO ORGANIZACIÓN
  // ===========================================================

  let logoDerecho;

  try {

    const idLogoDerecho =
      obtenerIdDrive(ligaLogoDerecho);

    if (!idLogoDerecho) {

      ui.alert(
        'La liga del logo de la organización no es válida.'
      );

      return;
    }

    logoDerecho = DriveApp
      .getFileById(idLogoDerecho)
      .getBlob();

  } catch (error) {

    ui.alert(
      'No se pudo acceder al logo de:\n\n' +
      organizacion +
      '\n\n' +
      'Verifica la liga y los permisos.'
    );

    return;
  }

  // ===========================================================
  // ENCABEZADOS TABLA A7:F7
  // ===========================================================

  const encabezadosTabla = hoja
    .getRange('A7:F7')
    .getDisplayValues()[0];

  // ===========================================================
  // ESTUDIANTES DESDE FILA 8
  // ===========================================================

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 8) {

    ui.alert(
      'No hay estudiantes capturados.'
    );

    return;
  }

  const datos = hoja
    .getRange(
      8,
      1,
      ultimaFila - 7,
      6
    )
    .getDisplayValues();

  const estudiantes = datos.filter(
    fila => fila.some(valor => valor !== '')
  );

  if (estudiantes.length === 0) {

    ui.alert(
      'No hay estudiantes capturados.'
    );

    return;
  }

  // ===========================================================
  // FECHA Y HORA
  // ===========================================================

  const ahora = new Date();

  const zonaHoraria =
    ss.getSpreadsheetTimeZone();

  const fechaCarta =
    Utilities.formatDate(
      ahora,
      zonaHoraria,
      'dd-MM-yyyy'
    );

  const horaCarta =
    Utilities.formatDate(
      ahora,
      zonaHoraria,
      'HH:mm'
    );

  // ===========================================================
  // FECHA SUPERIOR EN ESPAÑOL
  // ===========================================================

  const meses = [
    'enero',
    'febrero',
    'marzo',
    'abril',
    'mayo',
    'junio',
    'julio',
    'agosto',
    'septiembre',
    'octubre',
    'noviembre',
    'diciembre'
  ];

  const dia =
    Utilities.formatDate(
      ahora,
      zonaHoraria,
      'dd'
    );

  const mesNumero =
    Number(
      Utilities.formatDate(
        ahora,
        zonaHoraria,
        'M'
      )
    );

  const anioActual =
    Utilities.formatDate(
      ahora,
      zonaHoraria,
      'yyyy'
    );

  const fechaLarga =
    dia +
    ' de ' +
    meses[mesNumero - 1] +
    ' de ' +
    anioActual;

  // ===========================================================
  // CREAR DOCUMENTO TEMPORAL
  // ===========================================================

  const documento =
    DocumentApp.create(
      'TEMP_' + new Date().getTime()
    );

  const body =
    documento.getBody();

  body.clear();

  // ===========================================================
  // TAMAÑO CARTA
  // ===========================================================

  body.setPageWidth(612);
  body.setPageHeight(792);
  body.setMarginTop(25);
  body.setMarginBottom(65);
  body.setMarginLeft(40);
  body.setMarginRight(40);

  // ===========================================================
  // LOGOS
  // ===========================================================

  const encabezado =
    body.appendTable([
      ['', '']
    ]);

  encabezado.setBorderWidth(0);

  const celdaLogoIzq =
    encabezado.getCell(0, 0);

  const celdaLogoDer =
    encabezado.getCell(0, 1);

  celdaLogoIzq.setWidth(260);
  celdaLogoDer.setWidth(260);

  // ===========================================================
  // LOGO TEC IZQUIERDO
  // ===========================================================

  const imagenIzquierda =
    celdaLogoIzq
      .getChild(0)
      .asParagraph()
      .appendInlineImage(
        logoIzquierdo
      );

  ajustarImagen(
    imagenIzquierda,
    180,
    120
  );

  celdaLogoIzq
    .getChild(0)
    .asParagraph()
    .setAlignment(
      DocumentApp.HorizontalAlignment.LEFT
    );

  // ===========================================================
  // LOGO ORGANIZACIÓN DERECHO
  // ===========================================================

  const imagenDerecha =
    celdaLogoDer
      .getChild(0)
      .asParagraph()
      .appendInlineImage(
        logoDerecho
      );

  ajustarImagen(
    imagenDerecha,
    180,
    120
  );

  celdaLogoDer
    .getChild(0)
    .asParagraph()
    .setAlignment(
      DocumentApp.HorizontalAlignment.RIGHT
    );

  // ===========================================================
  // PIE DE PÁGINA
  // ===========================================================

  const footer =
    documento.addFooter();

  const textoFooter =
    'Carta firmada electrónicamente por ' +
    firmaIzquierda +
    ' a través de su cuenta de correo electrónico ' +
    correoFirmaIzquierda +
    ', en el campus Campus Monterrey el día ' +
    fechaCarta +
    ' a las ' +
    horaCarta;

  const parrafoFooter =
    footer.appendParagraph(
      textoFooter
    );

  parrafoFooter
    .setFontFamily('Arial')
    .setFontSize(7)
    .setBold(false);

  parrafoFooter.setAlignment(
    DocumentApp.HorizontalAlignment.LEFT
  );

  // ===========================================================
  // FECHA SUPERIOR DERECHA
  // ===========================================================

  let p =
    body.appendParagraph(
      fechaLarga
    );

  formatoNormal(p);

  p.setBold(false);

  p.setAlignment(
    DocumentApp.HorizontalAlignment.RIGHT
  );

  // ===========================================================
  // DESTINATARIO
  // ===========================================================

  p =
    body.appendParagraph(
      'A quién corresponda'
    );

  formatoNormal(p);
  p.setBold(false);

  // ===========================================================
  // CAMPUS MONTERREY
  // ===========================================================

  p =
    body.appendParagraph(
      'Campus Monterrey'
    );

  formatoNormal(p);
  p.setBold(true);

  // ===========================================================
  // PRESENTE
  // ===========================================================

  p =
    body.appendParagraph(
      'PRESENTE'
    );

  formatoNormal(p);
  p.setBold(false);

  body.appendParagraph('');

  // ===========================================================
  // ASUNTO
  // ===========================================================

  const textoAsunto =
    'Asunto: Carta de acreditación grupal ' +
    asunto;

  p =
    body.appendParagraph(
      textoAsunto
    );

  formatoNormal(p);
  p.setBold(false);

  ponerNegrita(
    p.editAsText(),
    textoAsunto,
    asunto
  );

  body.appendParagraph('');

  // ===========================================================
  // TEXTO PRINCIPAL
  // ===========================================================

  const texto =
    'Por medio de la presente, hago constar que los estudiantes ' +
    'del Tecnológico de Monterrey enlistados en este documento ' +
    'participaron en la Experiencia de servicio social ' +
    asunto +
    ' en colaboración con la Organización Socio Formadora ' +
    organizacion +
    ', y entregaron un producto o servicio que generó un impacto ' +
    'en beneficio de la sociedad. Esta experiencia se llevó a cabo ' +
    'en el periodo ' +
    periodoCompleto +
    '.';

  p =
    body.appendParagraph(
      texto
    );

  formatoNormal(p);

  p.setBold(false);

  p.setAlignment(
    DocumentApp.HorizontalAlignment.JUSTIFY
  );

  const textoObjeto =
    p.editAsText();

  ponerNegrita(
    textoObjeto,
    texto,
    periodoCompleto
  );

  body.appendParagraph('');

  // ===========================================================
  // SUBTÍTULO
  // ===========================================================

  p =
    body.appendParagraph(
      'De los siguientes estudiantes'
    );

  formatoNormal(p);
  p.setBold(false);

  body.appendParagraph('');

  // ===========================================================
  // TABLA ESTUDIANTES
  // ===========================================================

  const tablaDatos = [];

  tablaDatos.push(
    encabezadosTabla
  );

  estudiantes.forEach(
    (fila, index) => {

      tablaDatos.push([
        fila[0] ||
          String(index + 1),
        fila[1],
        fila[2],
        fila[3],
        fila[4],
        fila[5]
      ]);

    }
  );

  const tabla =
    body.appendTable(
      tablaDatos
    );

  // ===========================================================
  // FORMATO TABLA
  // ===========================================================

  for (
    let i = 0;
    i < tabla.getNumRows();
    i++
  ) {

    const fila =
      tabla.getRow(i);

    for (
      let j = 0;
      j < fila.getNumCells();
      j++
    ) {

      const celda =
        fila.getCell(j);

      celda
        .editAsText()
        .setFontFamily('Arial')
        .setFontSize(7.5)
        .setBold(false);

      if (i === 0) {

        celda
          .editAsText()
          .setBold(true);

      }

    }

  }

  // ===========================================================
  // ANCHO COLUMNAS
  // ===========================================================

  const anchos = [
    42,
    55,
    72,
    165,
    65,
    95
  ];

  for (
    let i = 0;
    i < tabla.getNumRows();
    i++
  ) {

    const fila =
      tabla.getRow(i);

    for (
      let j = 0;
      j < fila.getNumCells();
      j++
    ) {

      fila
        .getCell(j)
        .setWidth(
          anchos[j]
        );

    }

  }

  body.appendParagraph('');
  body.appendParagraph('');

  // ===========================================================
  // TEXTO FINAL
  // ===========================================================

  p =
    body.appendParagraph(
      'Sin más por el momento, quedo a sus órdenes para cualquier aclaración al respecto'
    );

  formatoNormal(p);
  p.setBold(false);

  body.appendParagraph('');
  body.appendParagraph('');

  // ===========================================================
  // FIRMAS
  // ===========================================================

  const tablaFirmas =
    body.appendTable([
      [
        '____________________________________________',
        '____________________________________________'
      ],
      [
        firmaIzquierda,
        firmaDerecha
      ]
    ]);

  tablaFirmas.setBorderWidth(0);

  tablaFirmas
    .getCell(0, 0)
    .setWidth(250);

  tablaFirmas
    .getCell(0, 1)
    .setWidth(250);

  // ===========================================================
  // CENTRAR LÍNEAS
  // ===========================================================

  tablaFirmas
    .getCell(0, 0)
    .getChild(0)
    .asParagraph()
    .setAlignment(
      DocumentApp.HorizontalAlignment.CENTER
    );

  tablaFirmas
    .getCell(0, 1)
    .getChild(0)
    .asParagraph()
    .setAlignment(
      DocumentApp.HorizontalAlignment.CENTER
    );

  // ===========================================================
  // CENTRAR NOMBRES
  // ===========================================================

  tablaFirmas
    .getCell(1, 0)
    .getChild(0)
    .asParagraph()
    .setAlignment(
      DocumentApp.HorizontalAlignment.CENTER
    );

  tablaFirmas
    .getCell(1, 1)
    .getChild(0)
    .asParagraph()
    .setAlignment(
      DocumentApp.HorizontalAlignment.CENTER
    );

  // ===========================================================
  // FORMATO LÍNEAS
  // ===========================================================

  tablaFirmas
    .getCell(0, 0)
    .editAsText()
    .setFontFamily('Arial')
    .setFontSize(9)
    .setBold(false);

  tablaFirmas
    .getCell(0, 1)
    .editAsText()
    .setFontFamily('Arial')
    .setFontSize(9)
    .setBold(false);

  // ===========================================================
  // NOMBRES EN NEGRITAS
  // ===========================================================

  tablaFirmas
    .getCell(1, 0)
    .editAsText()
    .setFontFamily('Arial')
    .setFontSize(9)
    .setBold(true);

  tablaFirmas
    .getCell(1, 1)
    .editAsText()
    .setFontFamily('Arial')
    .setFontSize(9)
    .setBold(true);

  // ===========================================================
  // GUARDAR DOCUMENTO
  // ===========================================================

  documento.saveAndClose();

  Utilities.sleep(1500);

  // ===========================================================
  // CONVERTIR A PDF
  // ===========================================================

  const archivoDoc =
    DriveApp.getFileById(
      documento.getId()
    );

  const pdfBlob =
    archivoDoc.getAs(
      MimeType.PDF
    );

  // ===========================================================
  // NOMBRE DEL PDF
  // ===========================================================

  let nombrePDF =
    asunto
      .replace(
        /[\\\/:\*?"<>|]/g,
        '-'
      )
      .trim();

  nombrePDF += '.pdf';

  pdfBlob.setName(
    nombrePDF
  );

  // ===========================================================
  // CREAR PDF EN DRIVE
  // ===========================================================

  const archivoPDF =
    DriveApp.createFile(
      pdfBlob
    );

  const idPDF =
    archivoPDF.getId();

  // ===========================================================
  // BORRAR GOOGLE DOC TEMPORAL
  // ===========================================================

  archivoDoc.setTrashed(true);

  // ===========================================================
  // LIGAS
  // ===========================================================

  const urlAbrir =
    'https://drive.google.com/file/d/' +
    idPDF +
    '/view';

  const urlDescargar =
    'https://drive.google.com/uc?export=download&id=' +
    idPDF;

  // ===========================================================
  // VENTANA PEQUEÑA
  // ===========================================================

  const html = `
<!DOCTYPE html>

<html>

<head>

<base target="_blank">

<style>

body {
  margin: 0;
  padding: 24px;
  font-family: Arial, sans-serif;
  background: white;
  text-align: center;
  color: #202124;
}

.titulo {
  font-size: 18px;
  font-weight: bold;
  margin-bottom: 7px;
}

.archivo {
  font-size: 13px;
  color: #5f6368;
  margin-bottom: 22px;
  overflow-wrap: anywhere;
}

.botones {
  display: flex;
  justify-content: center;
  gap: 10px;
}

.boton {
  display: inline-block;
  text-decoration: none;
  border-radius: 6px;
  padding: 13px 20px;
  font-size: 13px;
  font-weight: bold;
  cursor: pointer;
}

.abrir {
  background: #e8eaed;
  color: #202124;
}

.descargar {
  background: #0F9D58;
  color: white;
}

.nota {
  margin-top: 18px;
  font-size: 11px;
  line-height: 1.5;
  color: #777;
}

.cerrar {
  margin-top: 14px;
  border: none;
  background: transparent;
  color: #5f6368;
  font-size: 12px;
  cursor: pointer;
  text-decoration: underline;
}

</style>

</head>

<body>

<div class="titulo">
  Carta generada correctamente
</div>

<div class="archivo">
  ${escaparHtml(nombrePDF)}
</div>

<div class="botones">

<a
  class="boton abrir"
  href="${urlAbrir}"
  target="_blank"
>
  ABRIR PDF
</a>

<a
  class="boton descargar"
  href="${urlDescargar}"
  target="_blank"
>
  DESCARGAR PDF
</a>

</div>

<div class="nota">

<b>ABRIR PDF:</b>
permite revisar la carta antes de descargarla.

<br>

<b>DESCARGAR PDF:</b>
descarga directamente el archivo.

</div>

<button
  class="cerrar"
  onclick="google.script.host.close()"
>
Cerrar
</button>

</body>

</html>
`;

  const salida =
    HtmlService
      .createHtmlOutput(html)
      .setWidth(500)
      .setHeight(245);

  ui.showModalDialog(
    salida,
    'Carta generada'
  );
}


// =============================================================
// MANUAL DE LLENADO
// =============================================================

function mostrarManualLlenado() {

  const html =
    HtmlService.createHtmlOutput(`

<!DOCTYPE html>

<html>

<head>

<base target="_top">

<style>

body {
  font-family: Arial, sans-serif;
  margin: 0;
  padding: 24px;
  color: #202124;
  background: #ffffff;
}

.titulo {
  font-size: 21px;
  font-weight: bold;
  margin-bottom: 6px;
}

.subtitulo {
  font-size: 13px;
  color: #5f6368;
  margin-bottom: 18px;
}

.alerta {
  background: #fff4cc;
  border: 1px solid #f4b400;
  border-left: 7px solid #f4b400;
  border-radius: 7px;
  padding: 16px;
  margin-bottom: 20px;
  line-height: 1.55;
  font-size: 13px;
}

.alerta-titulo {
  font-size: 15px;
  font-weight: bold;
  margin-bottom: 6px;
}

.alerta-roja {
  color: #b31412;
  font-weight: bold;
}

.seccion {
  font-size: 15px;
  font-weight: bold;
  margin-top: 20px;
  margin-bottom: 10px;
}

table {
  width: 100%;
  border-collapse: collapse;
  font-size: 12px;
}

th {
  background: #f1f3f4;
  font-weight: bold;
  text-align: left;
}

th,
td {
  border: 1px solid #dadce0;
  padding: 9px;
  vertical-align: top;
}

.celda {
  font-weight: bold;
  white-space: nowrap;
}

.formato {
  border: 1px solid #dadce0;
  border-radius: 7px;
  padding: 12px;
  background: #fafafa;
  margin-top: 12px;
}

.fila {
  display: flex;
  gap: 6px;
  margin-bottom: 6px;
}

.bloque {
  padding: 9px;
  border: 1px solid #d0d0d0;
  background: white;
  font-size: 11px;
  border-radius: 3px;
}

.amarillo {
  background: #fff4cc;
  border-color: #e6c200;
}

.grande {
  flex: 3;
}

.mediano {
  flex: 2;
}

.pequeno {
  flex: 1;
}

.etiqueta {
  font-size: 10px;
  color: #5f6368;
  margin-top: 3px;
}

.pasos {
  font-size: 13px;
  line-height: 1.6;
  padding-left: 20px;
}

.nota {
  margin-top: 18px;
  background: #f1f8ff;
  border-left: 4px solid #1a73e8;
  padding: 12px;
  border-radius: 5px;
  font-size: 12px;
  line-height: 1.5;
}

.cerrar {
  display: block;
  margin: 22px auto 0 auto;
  padding: 10px 25px;
  border: none;
  border-radius: 5px;
  background: #1a73e8;
  color: white;
  font-size: 13px;
  font-weight: bold;
  cursor: pointer;
}

</style>

</head>

<body>

<div class="titulo">
  📘 Manual de llenado
</div>

<div class="subtitulo">
  Carta de acreditación grupal
</div>

<div class="alerta">

<div class="alerta-titulo">
  ⚠️ IMPORTANTE: NO MODIFICAR EL FORMATO
</div>

Las celdas que aparecen en
<b>color amarillo</b>
forman parte del formato de la carta.

<br><br>

<span class="alerta-roja">

NO mover, eliminar, borrar ni modificar las celdas amarillas,
encabezados o estructura del formato.

</span>

<br><br>

Únicamente captura la información solicitada
en los espacios correspondientes.

<br><br>

<b>
Si una celda está en amarillo, déjala como está.
</b>

</div>

<div class="seccion">
  👀 ¿Dónde se captura la información?
</div>

<div class="formato">

<div class="fila">

<div class="bloque amarillo grande">

<b>B1</b>
<br>
Proyecto / Experiencia

</div>

<div class="bloque amarillo grande">

<b>B2</b>
<br>
Organización

</div>

</div>

<div class="fila">

<div class="bloque amarillo pequeno">

<b>A3</b>
<br>
Tipo

</div>

<div class="bloque amarillo mediano">

<b>B3</b>
<br>
Periodo

</div>

<div class="bloque amarillo pequeno">

<b>C3</b>
<br>
Año

</div>

<div class="bloque mediano">

<b>G1 / H1 / G2</b>
<br>
Firmas

</div>

</div>

<div class="fila">

<div class="bloque amarillo grande">

<b>A7:F7</b>
<br>
Encabezados de estudiantes

</div>

</div>

<div class="fila">

<div class="bloque grande">

<b>A8:F</b>
<br>
Información de estudiantes

</div>

</div>

<div class="etiqueta">

El color amarillo indica campos que forman parte del formato.
No deben moverse ni eliminarse.

</div>

</div>

<div class="seccion">
  1. Información general
</div>

<table>

<tr>
<th>Celda</th>
<th>¿Qué debes colocar?</th>
</tr>

<tr>

<td class="celda">
B1
</td>

<td>

Nombre de la experiencia o proyecto.

<br><br>

Ejemplo:
<b>
Semana TEC de Inducción al Servicio Social
</b>

</td>

</tr>

<tr>

<td class="celda">
B2
</td>

<td>

Nombre de la Organización Socioformadora (OSF).

<br><br>

<b>Importante:</b>
utiliza el nombre correspondiente a la organización
registrada en la pestaña "Logos".

</td>

</tr>

<tr>

<td class="celda">
A3
</td>

<td>
Tipo de periodo correspondiente.
</td>

</tr>

<tr>

<td class="celda">
B3
</td>

<td>
Periodo correspondiente.
</td>

</tr>

<tr>

<td class="celda">
C3
</td>

<td>
Año correspondiente.
</td>

</tr>

</table>

<div class="seccion">
  2. Información de las personas que firman
</div>

<table>

<tr>
<th>Celda</th>
<th>¿Qué debes colocar?</th>
</tr>

<tr>

<td class="celda">
G1
</td>

<td>
Nombre de la persona principal que firma la carta.
</td>

</tr>

<tr>

<td class="celda">
H1
</td>

<td>
Correo electrónico de la persona que firma.
</td>

</tr>

<tr>

<td class="celda">
G2
</td>

<td>
Nombre de la segunda persona que firma la carta.
</td>

</tr>

</table>

<div class="seccion">
  3. Información de los estudiantes
</div>

<table>

<tr>

<th>Rango</th>
<th>¿Qué debes colocar?</th>

</tr>

<tr>

<td class="celda">
A7:F7
</td>

<td>

Estos son los encabezados de la tabla.
<b>No modificarlos.</b>

</td>

</tr>

<tr>

<td class="celda">
A8:F
</td>

<td>

Captura la información de los estudiantes
debajo de los encabezados correspondientes.

</td>

</tr>

</table>

<div class="seccion">
  4. Logo de la organización
</div>

<div class="nota">

<b>
🖼️ El logo se coloca automáticamente.
</b>

<br><br>

Cuando selecciones o modifiques la organización en
<b>B2</b>,
el sistema buscará automáticamente el logo correspondiente
en la pestaña <b>"Logos"</b>.

<br><br>

La celda <b>J1</b>
se utiliza para colocar automáticamente la liga del logo.

<br><br>

<b>
⚠️ No modificar J1 manualmente.
</b>

</div>

<div class="seccion">
  5. Antes de generar la carta
</div>

<ol class="pasos">

<li>
Verifica que la información general esté completa.
</li>

<li>
Revisa que el nombre de la Organización Socioformadora sea correcto.
</li>

<li>
Revisa los nombres de las personas que firman.
</li>

<li>
Verifica la información de todos los estudiantes.
</li>

<li>
Confirma que las horas correspondan al proyecto.
</li>

<li>
<b>
No muevas, elimines ni modifiques las celdas amarillas.
</b>
</li>

<li>

Cuando todo esté listo, utiliza:

<br><br>

<b>
📄 Generar PDF → Generar carta Acreditación
</b>

</li>

</ol>

<div class="alerta">

<div class="alerta-titulo">
  ⚠️ Recuerda
</div>

<b>

El color amarillo indica información que pertenece al formato
y que NO debe modificarse.

</b>

<br><br>

Si necesitas agregar estudiantes,
utiliza las filas disponibles debajo de los encabezados.

<br><br>

<span class="alerta-roja">

No cambies el orden de las columnas,
no borres los encabezados y no muevas las celdas amarillas.

</span>

</div>

<button
  class="cerrar"
  onclick="google.script.host.close()"
>
Cerrar manual
</button>

</body>

</html>

`)
    .setWidth(700)
    .setHeight(720);

  SpreadsheetApp
    .getUi()
    .showModalDialog(
      html,
      '📘 Manual de llenado'
    );
}


// =============================================================
// BUSCAR LOGO EN PESTAÑA "Logos"
// =============================================================

function buscarYColocarLogo(
  hojaDestino,
  organizacion
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const hojaLogos =
    ss.getSheetByName('Logos');

  if (!hojaLogos) {

    SpreadsheetApp
      .getUi()
      .alert(
        'No existe una pestaña llamada "Logos".'
      );

    return false;
  }

  const ultimaFila =
    hojaLogos.getLastRow();

  if (ultimaFila < 2) {

    hojaDestino
      .getRange('J1')
      .clearContent();

    return false;
  }

  const datos =
    hojaLogos
      .getRange(
        2,
        1,
        ultimaFila - 1,
        2
      )
      .getDisplayValues();

  const nombreBuscado =
    normalizarTexto(
      organizacion
    );

  let ligaEncontrada = '';

  for (
    let i = 0;
    i < datos.length;
    i++
  ) {

    const nombreOrganizacion =
      normalizarTexto(
        datos[i][0]
      );

    const liga =
      datos[i][1].trim();

    if (
      nombreOrganizacion ===
      nombreBuscado
    ) {

      ligaEncontrada =
        liga;

      break;
    }

  }

  if (ligaEncontrada) {

    hojaDestino
      .getRange('J1')
      .setValue(
        ligaEncontrada
      );

    return true;
  }

  hojaDestino
    .getRange('J1')
    .clearContent();

  return false;
}


// =============================================================
// PRUEBA MANUAL DEL LOGO
// =============================================================

function probarBusquedaLogo() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const hoja =
    ss.getActiveSheet();

  const ui =
    SpreadsheetApp.getUi();

  if (
    hoja.getName() === 'Logos'
  ) {

    ui.alert(
      'Ve a la hoja donde generas la carta.'
    );

    return;
  }

  const organizacion =
    hoja
      .getRange('B2')
      .getDisplayValue()
      .trim();

  if (!organizacion) {

    ui.alert(
      'No hay una organización capturada en B2.'
    );

    return;
  }

  const encontrado =
    buscarYColocarLogo(
      hoja,
      organizacion
    );

  if (encontrado) {

    const liga =
      hoja
        .getRange('J1')
        .getDisplayValue();

    ui.alert(
      'Logo encontrado correctamente.\n\n' +
      'Organización:\n' +
      organizacion +
      '\n\n' +
      'Liga colocada en J1:\n\n' +
      liga
    );

  } else {

    ui.alert(
      'No encontré la organización:\n\n' +
      organizacion +
      '\n\n' +
      'en la pestaña "Logos".'
    );

  }
}


// =============================================================
// NORMALIZAR TEXTO
// =============================================================

function normalizarTexto(texto) {

  return String(texto || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(
      /[\u0300-\u036f]/g,
      ''
    )
    .replace(
      /\s+/g,
      ' '
    );
}


// =============================================================
// OBTENER ID DE GOOGLE DRIVE
// =============================================================

function obtenerIdDrive(liga) {

  if (!liga) return null;

  liga = liga.trim();

  let coincidencia =
    liga.match(
      /\/d\/([a-zA-Z0-9_-]+)/
    );

  if (coincidencia) {
    return coincidencia[1];
  }

  coincidencia =
    liga.match(
      /[?&]id=([a-zA-Z0-9_-]+)/
    );

  if (coincidencia) {
    return coincidencia[1];
  }

  if (
    /^[a-zA-Z0-9_-]{20,}$/.test(liga)
  ) {

    return liga;
  }

  return null;
}


// =============================================================
// AJUSTAR IMAGEN SIN DEFORMAR
// =============================================================

function ajustarImagen(
  imagen,
  anchoMaximo,
  altoMaximo
) {

  const anchoOriginal =
    imagen.getWidth();

  const altoOriginal =
    imagen.getHeight();

  const escala =
    Math.min(
      anchoMaximo / anchoOriginal,
      altoMaximo / altoOriginal
    );

  imagen.setWidth(
    Math.round(
      anchoOriginal * escala
    )
  );

  imagen.setHeight(
    Math.round(
      altoOriginal * escala
    )
  );
}


// =============================================================
// FORMATO NORMAL
// =============================================================

function formatoNormal(parrafo) {

  parrafo
    .setFontFamily('Arial')
    .setFontSize(10);
}


// =============================================================
// PONER TEXTO ESPECÍFICO EN NEGRITAS
// =============================================================

function ponerNegrita(
  objetoTexto,
  textoCompleto,
  buscar
) {

  if (!buscar) return;

  let inicio = 0;

  while (
    (
      inicio =
      textoCompleto.indexOf(
        buscar,
        inicio
      )
    ) !== -1
  ) {

    const fin =
      inicio +
      buscar.length -
      1;

    objetoTexto.setBold(
      inicio,
      fin,
      true
    );

    inicio =
      fin + 1;
  }
}


// =============================================================
// ESCAPAR HTML
// =============================================================

function escaparHtml(texto) {

  return String(texto || '')
    .replace(
      /&/g,
      '&amp;'
    )
    .replace(
      /</g,
      '&lt;'
    )
    .replace(
      />/g,
      '&gt;'
    )
    .replace(
      /"/g,
      '&quot;'
    )
    .replace(
      /'/g,
      '&#039;'
    );
}


// ################################################################
// ################################################################
//                 CÓDIGO DE ÉTICA
// ################################################################
// ################################################################


// =============================================================
// ABRIR MENÚ DEL CÓDIGO DE ÉTICA
// =============================================================

function abrirMenuCodigoEtica() {

  const html = `
<!DOCTYPE html>

<html>

<head>

<base target="_top">

<style>

body {
  font-family: Arial, sans-serif;
  padding: 24px;
  margin: 0;
  color: #202124;
  text-align: center;
}

.titulo {
  font-size: 20px;
  font-weight: bold;
  margin-bottom: 8px;
}

.descripcion {
  font-size: 13px;
  color: #5f6368;
  line-height: 1.5;
  margin-bottom: 22px;
}

.boton {
  width: 100%;
  border: none;
  border-radius: 7px;
  padding: 14px;
  margin-bottom: 12px;
  font-size: 13px;
  font-weight: bold;
  cursor: pointer;
}

.individual {
  background: #1a73e8;
  color: white;
}

.masivo {
  background: #0F9D58;
  color: white;
}

.cerrar {
  background: transparent;
  color: #5f6368;
  text-decoration: underline;
}

.nota {
  margin-top: 10px;
  font-size: 11px;
  color: #777;
  line-height: 1.5;
}

</style>

</head>

<body>

<div class="titulo">
  📘 Código de Ética
</div>

<div class="descripcion">

En la pestaña <b>"Código de Ética"</b>:

<br><br>

<b>Columna A:</b> Persona que firma

<br>

<b>Columna B:</b> Organización Socio Formadora

<br><br>

Selecciona una fila para generar una carta individual
o genera todas las cartas de una sola vez.

</div>

<button
  class="boton individual"
  onclick="generarIndividual()"
>
  GENERAR CARTA DE LA FILA SELECCIONADA
</button>

<button
  class="boton masivo"
  onclick="generarMasivo()"
>
  DESCARGAR TODAS LAS CARTAS
</button>

<div class="nota">

Si una organización aparece varias veces,
se generará solamente una carta para esa organización.

</div>

<button
  class="boton cerrar"
  onclick="google.script.host.close()"
>
  Cerrar
</button>

<script>

function generarIndividual() {

  google.script.run
    .withSuccessHandler(function() {
      google.script.host.close();
    })
    .generarCodigoEticaPDF();

}

function generarMasivo() {

  google.script.run
    .withSuccessHandler(function() {
      google.script.host.close();
    })
    .generarCodigoEticaMasivo();

}

</script>

</body>

</html>
`;

  SpreadsheetApp
    .getUi()
    .showModalDialog(
      HtmlService
        .createHtmlOutput(html)
        .setWidth(430)
        .setHeight(370),
      'Código de Ética'
    );
}


// =============================================================
// GENERAR CÓDIGO DE ÉTICA INDIVIDUAL
// =============================================================
//
// Hoja: Código de Ética
//
// A = Persona que firma
// B = Organización Socio Formadora
//
// La fila seleccionada determina qué carta se genera.
// =============================================================

function generarCodigoEticaPDF() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const ui =
    SpreadsheetApp.getUi();

  const hoja =
    ss.getSheetByName(
      'Código de Ética'
    );

  if (!hoja) {

    ui.alert(
      'No existe una pestaña llamada "Código de Ética".'
    );

    return;
  }

  // ===========================================================
  // FILA SELECCIONADA
  // ===========================================================

  const rango =
    hoja.getActiveRange();

  if (!rango) {

    ui.alert(
      'Selecciona una fila en la pestaña "Código de Ética".'
    );

    return;
  }

  const fila =
    rango.getRow();

  if (fila < 2) {

    ui.alert(
      'Selecciona una fila de datos a partir de la fila 2.'
    );

    return;
  }

  // ===========================================================
  // DATOS A Y B
  // ===========================================================

  const nombre =
    hoja
      .getRange(fila, 1)
      .getDisplayValue()
      .trim();

  const organizacion =
    hoja
      .getRange(fila, 2)
      .getDisplayValue()
      .trim();

  if (!nombre) {

    ui.alert(
      'Falta capturar la persona que firma en la columna A de la fila ' +
      fila +
      '.'
    );

    return;
  }

  if (!organizacion) {

    ui.alert(
      'Falta capturar la Organización Socio Formadora en la columna B de la fila ' +
      fila +
      '.'
    );

    return;
  }

  // ===========================================================
  // CREAR PDF
  // ===========================================================

  const resultado =
    crearCodigoEticaPDF_(
      ss,
      nombre,
      organizacion
    );

  if (!resultado) {
    return;
  }

  const nombrePDF =
    resultado.nombrePDF;

  const idPDF =
    resultado.idPDF;

  const urlAbrir =
    resultado.urlAbrir;

  const urlDescargar =
    resultado.urlDescargar;

  // ===========================================================
  // VENTANA FINAL
  // ===========================================================

  const htmlSalida = `

<!DOCTYPE html>

<html>

<head>

<base target="_blank">

<style>

body {
  margin: 0;
  padding: 24px;
  font-family: Arial, sans-serif;
  background: #fff;
  text-align: center;
  color: #202124;
}

.titulo {
  font-size: 18px;
  font-weight: bold;
  margin-bottom: 7px;
}

.archivo {
  font-size: 13px;
  color: #5f6368;
  margin-bottom: 22px;
  overflow-wrap: anywhere;
}

.botones {
  display: flex;
  justify-content: center;
  gap: 10px;
}

.boton {
  display: inline-block;
  text-decoration: none;
  border-radius: 6px;
  padding: 13px 20px;
  font-size: 13px;
  font-weight: bold;
  cursor: pointer;
}

.abrir {
  background: #e8eaed;
  color: #202124;
}

.descargar {
  background: #0F9D58;
  color: white;
}

.nota {
  margin-top: 18px;
  font-size: 11px;
  line-height: 1.5;
  color: #777;
}

.cerrar {
  margin-top: 14px;
  border: none;
  background: transparent;
  color: #5f6368;
  font-size: 12px;
  cursor: pointer;
  text-decoration: underline;
}

</style>

</head>

<body>

<div class="titulo">
  Código de Ética generado correctamente
</div>

<div class="archivo">
  ${escaparHtml(nombrePDF)}
</div>

<div class="botones">

<a
  class="boton abrir"
  href="${urlAbrir}"
  target="_blank"
>
  ABRIR PDF
</a>

<a
  class="boton descargar"
  href="${urlDescargar}"
  target="_blank"
>
  DESCARGAR PDF
</a>

</div>

<div class="nota">

Nombre:
${escaparHtml(nombre)}

<br>

Organización:
${escaparHtml(organizacion)}

<br>

El archivo se guardó también en Google Drive.

</div>

<button
  class="cerrar"
  onclick="google.script.host.close()"
>
Cerrar
</button>

</body>

</html>
`;

  ui.showModalDialog(
    HtmlService
      .createHtmlOutput(
        htmlSalida
      )
      .setWidth(540)
      .setHeight(300),
    'Código de Ética'
  );
}


// =============================================================
// GENERAR TODAS LAS CARTAS DE CÓDIGO DE ÉTICA
// =============================================================
//
// Hoja: Código de Ética
//
// A = Persona que firma
// B = Organización Socio Formadora
//
// IMPORTANTE:
// Si una organización aparece varias veces,
// solamente se genera UNA carta.
//
// Se toma la primera persona registrada
// en columna A para esa organización.
//
// Al finalizar:
// - Se crean los PDFs individuales.
// - Se crea un ZIP.
// - El ZIP contiene un PDF por organización.
// =============================================================

function generarCodigoEticaMasivo() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const ui =
    SpreadsheetApp.getUi();

  const hoja =
    ss.getSheetByName(
      'Código de Ética'
    );

  if (!hoja) {

    ui.alert(
      'No existe una pestaña llamada "Código de Ética".'
    );

    return;
  }

  // ===========================================================
  // OBTENER DATOS
  // ===========================================================

  const ultimaFila =
    hoja.getLastRow();

  if (ultimaFila < 2) {

    ui.alert(
      'No hay organizaciones capturadas en la pestaña "Código de Ética".'
    );

    return;
  }

  const datos =
    hoja
      .getRange(
        2,
        1,
        ultimaFila - 1,
        2
      )
      .getDisplayValues();

  // ===========================================================
  // CREAR LISTA ÚNICA DE ORGANIZACIONES
  // ===========================================================

  const organizaciones =
    new Map();

  let filasValidas = 0;
  let duplicados = 0;

  datos.forEach(
    function(fila) {

      const nombre =
        String(
          fila[0] || ''
        ).trim();

      const organizacion =
        String(
          fila[1] || ''
        ).trim();

      // Saltar filas completamente vacías
      if (
        !nombre &&
        !organizacion
      ) {
        return;
      }

      // Si falta uno de los dos datos,
      // no se genera la carta.
      if (
        !nombre ||
        !organizacion
      ) {
        return;
      }

      filasValidas++;

      const clave =
        normalizarTexto(
          organizacion
        );

      if (
        organizaciones.has(clave)
      ) {

        duplicados++;

        return;
      }

      organizaciones.set(
        clave,
        {
          nombre: nombre,
          organizacion: organizacion
        }
      );

    }
  );

  const lista =
    Array.from(
      organizaciones.values()
    );

  if (lista.length === 0) {

    ui.alert(
      'No se encontraron registros completos.\n\n' +
      'Verifica que:\n' +
      '• Columna A tenga la persona que firma.\n' +
      '• Columna B tenga la Organización Socio Formadora.'
    );

    return;
  }

  // ===========================================================
  // CONFIRMACIÓN
  // ===========================================================

  const mensajeConfirmacion =

    'Se generarán ' +
    lista.length +
    ' cartas de Código de Ética.\n\n' +

    'Se creará UNA carta por organización.\n\n' +

    'Registros completos encontrados: ' +
    filasValidas +
    '\n' +

    'Organizaciones únicas: ' +
    lista.length +
    '\n' +

    'Registros duplicados: ' +
    duplicados +
    '\n\n' +

    'Los archivos individuales se guardarán en Google Drive ' +
    'y además se creará un ZIP para descargar todas las cartas juntas.\n\n' +

    '¿Deseas continuar?';

  const respuesta =
    ui.alert(
      'Descargar todas las cartas',
      mensajeConfirmacion,
      ui.ButtonSet.YES_NO
    );

  if (
    respuesta !==
    ui.Button.YES
  ) {
    return;
  }

  // ===========================================================
  // AVISO EN LA HOJA
  // ===========================================================

  ss.toast(
    'Generando ' +
    lista.length +
    ' cartas. Este proceso puede tardar algunos minutos.',
    'Código de Ética',
    10
  );

  // ===========================================================
  // GENERAR PDF POR ORGANIZACIÓN
  // ===========================================================

  const blobsPDF = [];

  let generadas = 0;
  let errores = [];

  lista.forEach(
    function(item) {

      try {

        const resultado =
          crearCodigoEticaPDF_(
            ss,
            item.nombre,
            item.organizacion
          );

        if (
          resultado &&
          resultado.pdfBlob
        ) {

          const blob =
            resultado.pdfBlob.copyBlob();

          blob.setName(
            resultado.nombrePDF
          );

          blobsPDF.push(
            blob
          );

          generadas++;

        }

        // Pequeña pausa para dar tiempo
        // a Drive a procesar cada archivo.
        Utilities.sleep(500);

      } catch (error) {

        errores.push(
          item.organizacion +
          ': ' +
          error.message
        );

      }

    }
  );

  // ===========================================================
  // VALIDAR QUE HAYA PDFs
  // ===========================================================

  if (
    blobsPDF.length === 0
  ) {

    ui.alert(
      'No fue posible generar ninguna carta.\n\n' +
      errores.join('\n')
    );

    return;
  }

  // ===========================================================
  // CREAR ZIP
  // ===========================================================

  const anio =
    obtenerAnioActual_(
      ss
    );

  const nombreZIP =
    'CE_' +
    anio +
    '_Cartas_Codigo_de_Etica.zip';

  let archivoZIP;

  try {

    const zipBlob =
      Utilities.zip(
        blobsPDF,
        nombreZIP
      );

    archivoZIP =
      DriveApp.createFile(
        zipBlob
      );

  } catch (error) {

    ui.alert(
      'Las cartas fueron generadas, pero no fue posible crear el ZIP.\n\n' +
      'Puedes encontrar los PDFs individuales en Google Drive.\n\n' +
      'Detalle:\n' +
      error.message
    );

    return;
  }

  // ===========================================================
  // URL DEL ZIP
  // ===========================================================

  const idZIP =
    archivoZIP.getId();

  const urlZIP =
    'https://drive.google.com/uc?export=download&id=' +
    idZIP;

  const urlAbrirZIP =
    'https://drive.google.com/file/d/' +
    idZIP +
    '/view';

  // ===========================================================
  // RESULTADO FINAL
  // ===========================================================

  let textoErrores = '';

  if (
    errores.length > 0
  ) {

    textoErrores =
      '<div class="errores">' +
      '<b>Algunas cartas no pudieron generarse:</b><br><br>' +
      errores
        .map(
          e => escaparHtml(e)
        )
        .join('<br>') +
      '</div>';

  }

  const htmlResultado = `

<!DOCTYPE html>

<html>

<head>

<base target="_blank">

<style>

body {
  margin: 0;
  padding: 25px;
  font-family: Arial, sans-serif;
  background: white;
  color: #202124;
  text-align: center;
}

.icono {
  font-size: 38px;
  margin-bottom: 8px;
}

.titulo {
  font-size: 20px;
  font-weight: bold;
  margin-bottom: 10px;
}

.descripcion {
  font-size: 13px;
  color: #5f6368;
  line-height: 1.5;
  margin-bottom: 20px;
}

.resumen {
  background: #f1f8f4;
  border-radius: 8px;
  padding: 14px;
  margin-bottom: 20px;
  font-size: 13px;
  line-height: 1.7;
}

.descargar {
  display: inline-block;
  text-decoration: none;
  background: #0F9D58;
  color: white;
  border-radius: 7px;
  padding: 15px 25px;
  font-size: 14px;
  font-weight: bold;
  margin-bottom: 10px;
}

.abrir {
  display: inline-block;
  text-decoration: none;
  background: #e8eaed;
  color: #202124;
  border-radius: 7px;
  padding: 11px 20px;
  font-size: 12px;
  font-weight: bold;
}

.nota {
  margin-top: 18px;
  font-size: 11px;
  line-height: 1.5;
  color: #777;
}

.errores {
  margin-top: 18px;
  background: #fce8e6;
  color: #b31412;
  border-radius: 6px;
  padding: 12px;
  text-align: left;
  font-size: 11px;
  max-height: 100px;
  overflow-y: auto;
}

.cerrar {
  display: block;
  margin: 18px auto 0 auto;
  border: none;
  background: transparent;
  color: #5f6368;
  font-size: 12px;
  cursor: pointer;
  text-decoration: underline;
}

</style>

</head>

<body>

<div class="icono">
  ✓
</div>

<div class="titulo">
  Cartas generadas correctamente
</div>

<div class="descripcion">

Se creó una carta de Código de Ética
por cada Organización Socio Formadora.

</div>

<div class="resumen">

<b>Organizaciones únicas:</b>
${lista.length}

<br>

<b>Cartas generadas:</b>
${generadas}

<br>

<b>Registros duplicados:</b>
${duplicados}

</div>

<a
  class="descargar"
  href="${urlZIP}"
  target="_blank"
>
  ⬇ DESCARGAR TODAS LAS CARTAS
</a>

<br>

<a
  class="abrir"
  href="${urlAbrirZIP}"
  target="_blank"
>
  ABRIR ZIP
</a>

<div class="nota">

El archivo ZIP contiene un PDF independiente
para cada organización.

<br><br>

Los PDFs individuales también fueron guardados
en Google Drive.

</div>

${textoErrores}

<button
  class="cerrar"
  onclick="google.script.host.close()"
>
Cerrar
</button>

</body>

</html>
`;

  ui.showModalDialog(
    HtmlService
      .createHtmlOutput(
        htmlResultado
      )
      .setWidth(560)
      .setHeight(
        errores.length > 0
          ? 500
          : 390
      ),
    'Cartas de Código de Ética'
  );
}


// =============================================================
// CREAR PDF DE CÓDIGO DE ÉTICA
// =============================================================
//
// Esta función es utilizada tanto para:
// 1. Generación individual.
// 2. Generación masiva.
//
// NO debe llamarse directamente desde el menú.
// =============================================================

function crearCodigoEticaPDF_(
  ss,
  nombre,
  organizacion
) {

  // ===========================================================
  // LOGO TEC
  // ===========================================================

  const ID_LOGO_TEC =
    '16UKagcpT4pUaE8F0TLl36oIxrBfPDgl9';

  // ===========================================================
  // FECHA
  // ===========================================================

  const zonaHoraria =
    ss.getSpreadsheetTimeZone();

  const ahora =
    new Date();

  const meses = [
    'enero',
    'febrero',
    'marzo',
    'abril',
    'mayo',
    'junio',
    'julio',
    'agosto',
    'septiembre',
    'octubre',
    'noviembre',
    'diciembre'
  ];

  const dia =
    Utilities.formatDate(
      ahora,
      zonaHoraria,
      'd'
    );

  const mes =
    meses[
      Number(
        Utilities.formatDate(
          ahora,
          zonaHoraria,
          'M'
        )
      ) - 1
    ];

  const anio =
    Utilities.formatDate(
      ahora,
      zonaHoraria,
      'yyyy'
    );

  const fechaLarga =
    dia +
    ' de ' +
    mes +
    ' de ' +
    anio;

  // ===========================================================
  // CREAR DOCUMENTO TEMPORAL
  // ===========================================================

  const documento =
    DocumentApp.create(
      'TEMP_CODIGO_ETICA_' +
      new Date().getTime()
    );

  const body =
    documento.getBody();

  body.clear();

  // ===========================================================
  // CONFIGURACIÓN DE PÁGINA
  // ===========================================================

  body.setPageWidth(612);
  body.setPageHeight(792);

  body.setMarginTop(28);
  body.setMarginBottom(28);
  body.setMarginLeft(42);
  body.setMarginRight(42);

  // ===========================================================
  // LOGO TEC
  // ===========================================================

  try {

    const logoBlob =
      DriveApp
        .getFileById(
          ID_LOGO_TEC
        )
        .getBlob();

    const tablaEncabezado =
      body.appendTable([
        ['', '']
      ]);

    tablaEncabezado
      .setBorderWidth(0);

    tablaEncabezado
      .getCell(0, 0)
      .setWidth(250);

    tablaEncabezado
      .getCell(0, 1)
      .setWidth(250);

    const pLogo =
      tablaEncabezado
        .getCell(0, 1)
        .getChild(0)
        .asParagraph();

    pLogo.setAlignment(
      DocumentApp.HorizontalAlignment.RIGHT
    );

    const img =
      pLogo.appendInlineImage(
        logoBlob
      );

    ajustarImagen(
      img,
      190,
      100
    );

  } catch (e) {

    // El PDF puede continuar sin logo
  }

  // ===========================================================
  // TÍTULO
  // ===========================================================

  let p =
    body.appendParagraph(
      'Código de Ética'
    );

  p
    .setFontFamily('Arial')
    .setFontSize(20)
    .setBold(true)
    .setForegroundColor('#D81B60')
    .setAlignment(
      DocumentApp.HorizontalAlignment.LEFT
    )
    .setSpacingBefore(5)
    .setSpacingAfter(12);

  // ===========================================================
  // PRINCIPIOS
  // ===========================================================

  const principios = [

    'Reconocemos la dignidad de las personas y las tratamos con respeto y justicia.',

    'Tratamos a todos con equidad y evitamos cualquier tipo de discriminación.',

    'Actuamos de manera íntegra, honesta, responsable, objetiva, congruente e imparcial.',

    'Reconocemos y respetamos la propiedad intelectual y el mérito de los demás.',

    'Evitamos conflictos de interés de cualquier tipo y, en caso de presentarse, los informamos ante las instancias correspondientes.',

    'Asumimos como compromiso la transparencia en la información, y respetamos la confidencialidad de los asuntos que la Institución así determina.',

    'Utilizamos los recursos de manera responsable, austera y eficiente.',

    'Protegemos el medio ambiente.',

    'Buscamos el beneficio de la Institución por encima del beneficio personal.',

    'Cumplimos con las leyes, reglamentos y políticas que rigen nuestra actividad a nivel institucional, nacional e internacional.'

  ];

  principios.forEach(
    function(texto, i) {

      const numero =
        String(i + 1);

      const par =
        body.appendParagraph(
          numero +
          ' ' +
          texto
        );

      par
        .setFontFamily('Arial')
        .setFontSize(9.5)
        .setBold(false)
        .setForegroundColor('#000000')
        .setAlignment(
          DocumentApp.HorizontalAlignment.JUSTIFY
        )
        .setSpacingAfter(5)
        .setLineSpacing(1);

      const edit =
        par.editAsText();

      edit.setForegroundColor(
        0,
        numero.length - 1,
        '#D81B60'
      );

      edit.setBold(
        0,
        numero.length - 1,
        true
      );

      edit.setFontSize(
        0,
        numero.length - 1,
        20
      );

    }
  );

  // ===========================================================
  // CARTA COMPROMISO
  // ===========================================================

  p =
    body.appendParagraph(
      'Carta compromiso'
    );

  p
    .setFontFamily('Arial')
    .setFontSize(20)
    .setBold(true)
    .setForegroundColor('#D81B60')
    .setAlignment(
      DocumentApp.HorizontalAlignment.LEFT
    )
    .setSpacingBefore(14)
    .setSpacingAfter(12);

  // ===========================================================
  // TEXTO COMPROMISO
  // ===========================================================

  const compromiso =
    'Yo ' +
    nombre +
    ' de la Organización Socio Formadora ' +
    organizacion +
    ' del Tecnológico de Monterrey, durante el ' +
    anio +
    ' me comprometo a cumplir el Código de Ética y convertirlo en parte de mi vida y quehacer cotidiano.';

  p =
    body.appendParagraph(
      compromiso
    );

  p
    .setFontFamily('Arial')
    .setFontSize(10)
    .setBold(false)
    .setForegroundColor('#000000')
    .setAlignment(
      DocumentApp.HorizontalAlignment.JUSTIFY
    )
    .setSpacingAfter(24)
    .setLineSpacing(1);

  // ===========================================================
  // SUBRAYAR NOMBRE
  // ===========================================================

  const textoComp =
    p.editAsText();

  const inicioNombre =
    compromiso.indexOf(
      nombre
    );

  if (
    inicioNombre >= 0
  ) {

    textoComp.setUnderline(
      inicioNombre,
      inicioNombre +
      nombre.length -
      1,
      true
    );

  }

  // ===========================================================
  // SUBRAYAR ORGANIZACIÓN
  // ===========================================================

  const inicioOrg =
    compromiso.indexOf(
      organizacion
    );

  if (
    inicioOrg >= 0
  ) {

    textoComp.setUnderline(
      inicioOrg,
      inicioOrg +
      organizacion.length -
      1,
      true
    );

  }

  // ===========================================================
  // FIRMA Y FECHA
  // ===========================================================

  const tablaFirma =
    body.appendTable([

      [
        '',
        fechaLarga
      ],

      [
        '_______________________________________________',
        '_______________________________________________'
      ],

      [
        'Firma',
        'Fecha'
      ]

    ]);

  tablaFirma
    .setBorderWidth(0);

  for (
    let r = 0;
    r < tablaFirma.getNumRows();
    r++
  ) {

    for (
      let c = 0;
      c < tablaFirma
        .getRow(r)
        .getNumCells();
      c++
    ) {

      const celda =
        tablaFirma
          .getCell(r, c);

      celda.setWidth(250);

      const t =
        celda.editAsText();

      t
        .setFontFamily('Arial')
        .setBold(false);

      if (r === 0) {

        t.setFontSize(10);

      } else {

        t.setFontSize(9);

      }

      celda
        .getChild(0)
        .asParagraph()
        .setAlignment(
          DocumentApp.HorizontalAlignment.CENTER
        );

    }

  }

  // ===========================================================
  // GUARDAR DOCUMENTO
  // ===========================================================

  documento.saveAndClose();

  Utilities.sleep(1000);

  // ===========================================================
  // CONVERTIR A PDF
  // ===========================================================

  const archivoDoc =
    DriveApp.getFileById(
      documento.getId()
    );

  const pdfBlob =
    archivoDoc.getAs(
      MimeType.PDF
    );

  // ===========================================================
  // NOMBRE DEL PDF
  // ===========================================================

  const nombrePDF =
    'CE_' +
    limpiarNombreArchivo_(
      anio
    ) +
    '_' +
    limpiarNombreArchivo_(
      organizacion
    ) +
    '.pdf';

  pdfBlob.setName(
    nombrePDF
  );

  // ===========================================================
  // CREAR PDF EN DRIVE
  // ===========================================================

  const archivoPDF =
    DriveApp.createFile(
      pdfBlob
    );

  const idPDF =
    archivoPDF.getId();

  // ===========================================================
  // BORRAR DOCUMENTO TEMPORAL
  // ===========================================================

  archivoDoc.setTrashed(
    true
  );

  // ===========================================================
  // LIGAS
  // ===========================================================

  const urlAbrir =
    'https://drive.google.com/file/d/' +
    idPDF +
    '/view';

  const urlDescargar =
    'https://drive.google.com/uc?export=download&id=' +
    idPDF;

  // ===========================================================
  // RETORNAR INFORMACIÓN
  // ===========================================================

  return {

    archivoPDF:
      archivoPDF,

    pdfBlob:
      pdfBlob,

    idPDF:
      idPDF,

    nombrePDF:
      nombrePDF,

    urlAbrir:
      urlAbrir,

    urlDescargar:
      urlDescargar,

    nombre:
      nombre,

    organizacion:
      organizacion,

    fechaLarga:
      fechaLarga

  };
}


// =============================================================
// LIMPIAR NOMBRE DE ARCHIVO
// =============================================================

function limpiarNombreArchivo_(valor) {

  return String(
    valor || ''
  )
    .replace(
      /[\\\/:\*?"<>|]/g,
      '-'
    )
    .replace(
      /\s+/g,
      ' '
    )
    .trim();
}


// =============================================================
// OBTENER AÑO ACTUAL
// =============================================================

function obtenerAnioActual_(ss) {

  const zonaHoraria =
    ss.getSpreadsheetTimeZone();

  return Utilities.formatDate(
    new Date(),
    zonaHoraria,
    'yyyy'
  );
}
