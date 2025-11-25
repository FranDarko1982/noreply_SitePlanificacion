function consolidarUltimoDocEnHistorico() {
  const props = PropertiesService.getScriptProperties();
  const FOLDER_ID = props.getProperty('FOLDER_ID');
  const HISTORICO_DOC_ID = props.getProperty('HISTORICO_DOC_ID');
  const lastProcessedId = props.getProperty('LAST_DOC_ID');
  const TIMEZONE = 'Europe/Madrid';
  const PREFIX = 'Resumen Week'; // prefijo obligatorio en el nombre

  // === 1) Buscar el último Google Doc por fecha de creación,
  //        cuyo nombre empiece por "Resumen Week"
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();

  let ultimoFile = null;
  let ultimaFecha = null;

  while (files.hasNext()) {
    const file = files.next();

    // Solo Docs
    if (file.getMimeType() !== MimeType.GOOGLE_DOCS) continue;

    const name = file.getName();
    // Solo los que empiezan por "Resumen Week"
    if (!name || !name.startsWith(PREFIX)) continue;

    const created = file.getDateCreated();
    if (!ultimaFecha || created > ultimaFecha) {
      ultimaFecha = created;
      ultimoFile = file;
    }
  }

  if (!ultimoFile) {
    throw new Error('No se ha encontrado ningún Google Doc en la carpeta cuyo nombre empiece por "Resumen Week".');
  }

  const sourceDocId = ultimoFile.getId();
  const sourceDocName = ultimoFile.getName();
  const sourceCreatedDate = ultimoFile.getDateCreated();

  // === 2) Evitar duplicados ===
  if (sourceDocId === lastProcessedId) {
    Logger.log('El último documento ("' + sourceDocName + '") ya fue consolidado. No se repite.');
    return;
  }

  // === 3) Abrir documentos ===
  const sourceDoc = DocumentApp.openById(sourceDocId);
  const targetDoc = DocumentApp.openById(HISTORICO_DOC_ID);

  const sourceBody = sourceDoc.getBody();
  const targetBody = targetDoc.getBody();

  // === 4) Añadir cabecera de consolidación ===
  const fechaTexto = Utilities.formatDate(
    sourceCreatedDate,
    TIMEZONE,
    'yyyy-MM-dd HH:mm:ss'
  );

  targetBody.appendParagraph('--------------------------------------------------')
            .setBold(true);

  targetBody.appendParagraph('Consolidación de documento')
            .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  targetBody.appendParagraph('Documento origen: ' + sourceDocName);
  targetBody.appendParagraph('ID origen: ' + sourceDocId);
  targetBody.appendParagraph('Fecha de creación del archivo: ' + fechaTexto);

  targetBody.appendParagraph('');
  targetBody.appendParagraph('Contenido:')
            .setHeading(DocumentApp.ParagraphHeading.HEADING3);
  targetBody.appendParagraph('');

  // === 5) Copiar contenido del documento origen ===
  const numChildren = sourceBody.getNumChildren();
  for (let i = 0; i < numChildren; i++) {
    const element = sourceBody.getChild(i).copy();
    const type = element.getType();

    switch (type) {
      case DocumentApp.ElementType.PARAGRAPH:
        targetBody.appendParagraph(element);
        break;
      case DocumentApp.ElementType.LIST_ITEM:
        targetBody.appendListItem(element);
        break;
      case DocumentApp.ElementType.TABLE:
        targetBody.appendTable(element);
        break;
      default:
        try {
          targetBody.appendParagraph(element.asText().getText());
        } catch (e) {
          targetBody.appendParagraph('[Elemento no soportado]');
        }
    }
  }

  targetBody.appendPageBreak();
  targetDoc.saveAndClose();

  // === 6) Guardar el último ID procesado para evitar duplicados ===
  props.setProperty('LAST_DOC_ID', sourceDocId);
  props.setProperty('LAST_DOC_CREATED', sourceCreatedDate.toISOString());

  Logger.log('Consolidado: ' + sourceDocName + ' (' + sourceDocId + ')');
}
