/**
 * Google Apps Script — Upload de documents Budget Construction
 *
 * Rôle :
 * - afficher une page HTML d’upload ;
 * - recevoir un fichier encodé en base64 ;
 * - enregistrer le fichier dans Google Drive ;
 * - écrire les métadonnées du fichier dans une feuille Google Sheets ;
 * - retourner l’URL du fichier au classeur Excel.
 *
 * Configuration :
 * Remplacer les placeholders ci-dessous par vos propres valeurs.
 */

const CONFIG = {
  DRIVE_FOLDER_ID: "<GOOGLE_DRIVE_FOLDER_ID>",
  SPREADSHEET_ID: "<GOOGLE_SPREADSHEET_ID>",
  SHEET_NAME: "input_staging"
};

function doGet(e) {
  const template = HtmlService.createTemplateFromFile("Index");

  template.docType = e.parameter.docType || "";
  template.fournisseur = e.parameter.fournisseur || "";
  template.ref = e.parameter.ref || "";
  template.sous_produit = e.parameter.sous_produit || "";
  template.input_id = e.parameter.input_id || "";

  return template
    .evaluate()
    .setTitle("Upload fichier");
}

function doPost(e) {
  try {
    const result = uploadFile(
      e.parameter.file,
      e.parameter.fileName,
      e.parameter.docType,
      e.parameter.fournisseur,
      e.parameter.ref,
      e.parameter.sous_produit,
      e.parameter.input_id
    );

    return createJsonResponse(result);
  } catch (err) {
    return createJsonResponse({
      success: false,
      error: err.message
    });
  }
}

function uploadFile(
  base64,
  fileName,
  docType,
  fournisseur,
  ref,
  sous_produit,
  input_id
) {
  try {
    validateUploadData(base64, fileName, docType, fournisseur, ref, sous_produit, input_id);

    const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    const sheet = getTargetSheet();

    const blob = createBlobFromBase64(base64, fileName);
    const newFileName = generateFileName(docType, fournisseur, ref, sous_produit);

    const file = folder.createFile(blob).setName(newFileName);
    const fileUrl = file.getUrl();

    const fichier_id = getNextFileId(sheet);

    sheet.appendRow([
      input_id,
      newFileName,
      fileUrl,
      fichier_id,
      new Date()
    ]);

    forcePlainTextUrl(sheet, sheet.getLastRow(), 3, fileUrl);

    return {
      success: true,
      input_id: input_id,
      fichier_id: fichier_id,
      fileName: newFileName,
      fileUrl: fileUrl
    };
  } catch (err) {
    return {
      success: false,
      error: err.message
    };
  }
}

function validateUploadData(
  base64,
  fileName,
  docType,
  fournisseur,
  ref,
  sous_produit,
  input_id
) {
  if (!base64 || !fileName) {
    throw new Error("Fichier manquant.");
  }

  if (!input_id) {
    throw new Error("input_id manquant.");
  }

  if (!docType || !fournisseur || !ref || !sous_produit) {
    throw new Error(
      "Métadonnées manquantes : type, fournisseur, référence ou sous-produit."
    );
  }
}

function createBlobFromBase64(base64, fileName) {
  const bytes = Utilities.base64Decode(base64);

  return Utilities.newBlob(
    bytes,
    "application/octet-stream",
    fileName
  );
}

function getTargetSheet() {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error("Feuille introuvable : " + CONFIG.SHEET_NAME);
  }

  return sheet;
}

function getNextFileId(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return 1;
  }

  const lastValue = sheet.getRange(lastRow, 4).getValue();

  if (!isNaN(lastValue) && lastValue !== "") {
    return Number(lastValue) + 1;
  }

  return 1;
}

function forcePlainTextUrl(sheet, row, column, url) {
  const plainUrl = SpreadsheetApp
    .newRichTextValue()
    .setText(url)
    .build();

  sheet
    .getRange(row, column)
    .clearContent()
    .setNumberFormat("@")
    .setRichTextValue(plainUrl)
    .setShowHyperlink(false);
}

function generateFileName(docType, fournisseur, ref, sous_produit) {
  const now = new Date();

  const formattedDate = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    "yyyyMMdd"
  );

  const shortId = Utilities
    .getUuid()
    .replace(/-/g, "")
    .substring(0, 6)
    .toUpperCase();

  const docId = "DOC" + shortId;

  return [
    formattedDate,
    cleanFileName(docType),
    cleanFileName(fournisseur),
    cleanFileName(ref),
    cleanFileName(sous_produit),
    docId
  ].join("_") + ".pdf";
}

function cleanFileName(value) {
  if (!value) {
    return "";
  }

  return value
    .toString()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "")
    .toUpperCase();
}

function createJsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Utilitaire de réparation.
 *
 * À utiliser uniquement si la colonne source_drive contient des noms de fichiers
 * au lieu des URLs Drive.
 */
function repairSourceDriveFromFileNames() {
  const sourceDriveCol = 3;
  const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  const sheet = getTargetSheet();

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return;
  }

  for (let row = 2; row <= lastRow; row++) {
    const cell = sheet.getRange(row, sourceDriveCol);
    const currentText = String(cell.getDisplayValue()).trim();

    if (!currentText || currentText.startsWith("http")) {
      continue;
    }

    const files = folder.getFilesByName(currentText);

    if (files.hasNext()) {
      const file = files.next();
      const url = file.getUrl();

      forcePlainTextUrl(sheet, row, sourceDriveCol, url);
    }
  }
}
