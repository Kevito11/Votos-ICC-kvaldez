import { useState, useRef, useEffect } from 'react';
import ExcelJS from 'exceljs/dist/exceljs.min.js';
import { 
  addCandidateToSheets, 
  deleteCandidateFromSheets, 
  clearCandidatesInSheets,
  updateCandidateInSheets,
  updateVotersInSheets, 
  addVoteToSheets,
  resetVoterVotesInSheets,
  clearVotesInSheets,
  resetMultipleVotersVotesInSheets
} from '../utils/api';
import { MOCK_CANDIDATES, MOCK_VOTERS, MOCK_VOTES } from '../utils/mockData';
import ImageCropper from './ImageCropper';
import CandidatePhoto from './CandidatePhoto';
import Tooltip from './Tooltip';


export default function AdminPanel({
  config,
  candidates,
  voters,
  votes,
  isLoading,
  isConnected,
  refreshData,
  showToast,
  setCandidates,
  setVoters,
  setVotes
}) {
  const [activeTab, setActiveTab] = useState('results');
  const [isExporting, setIsExporting] = useState(false);

  // (Estados locales de formulario de configuración de conexiones removidos)

  // Estados del Formulario de Candidato
  const [candFirstName, setCandFirstName] = useState('');
  const [candLastName, setCandLastName] = useState('');
  const [candTestimony, setCandTestimony] = useState('');
  const [candPhotoFile, setCandPhotoFile] = useState(null);
  const [cropImageFile, setCropImageFile] = useState(null);
  const [isUploading, setIsUploading] = useState(false);
  const fileInputRef = useRef(null);
  const [editingCandidateId, setEditingCandidateId] = useState(null);
  const [expandedCandidateId, setExpandedCandidateId] = useState(null);

  // Estados de Gestión de Votantes
  const [voterFirstName, setVoterFirstName] = useState('');
  const [voterLastName, setVoterLastName] = useState('');
  const [bulkVotersText, setBulkVotersText] = useState('');
  const [voterSearch, setVoterSearch] = useState('');
  const [voterPresenceFilter, setVoterPresenceFilter] = useState('all'); // 'all' | 'present' | 'absent'
  const [voterProgressFilter, setVoterProgressFilter] = useState('all'); // 'all' | 'pending' | 'partial' | 'completed'
  const [selectedVoterIds, setSelectedVoterIds] = useState([]);
  const [activeHeaderFilter, setActiveHeaderFilter] = useState(null); // null | 'name' | 'presence' | 'progress'
  const [voterSortKey, setVoterSortKey] = useState('name'); // 'name' | 'presence' | 'progress'
  const [voterSortDirection, setVoterSortDirection] = useState('asc'); // 'asc' | 'desc'

  // Cerrar los desplegables de filtros al hacer clic fuera de ellos (Excel-like)
  useEffect(() => {
    const handleClickOutside = (e) => {
      // Si el clic fue en un disparador o en un menú del filtro, no hacer nada
      if (e.target.closest('.excel-filter-trigger') || e.target.closest('.excel-filter-dropdown')) {
        return;
      }
      setActiveHeaderFilter(null);
    };

    if (activeHeaderFilter) {
      document.addEventListener('click', handleClickOutside);
    }
    return () => {
      document.removeEventListener('click', handleClickOutside);
    };
  }, [activeHeaderFilter]);

  // Filtro de resultados
  const [resultsFilter, setResultsFilter] = useState('all');
  const [votesPerPage, setVotesPerPage] = useState(10);
  const [votesCurrentPage, setVotesCurrentPage] = useState(1);

  // Secciones colapsables en Resultados
  const [showCandidateResults, setShowCandidateResults] = useState(true);
  const [showVotesList, setShowVotesList] = useState(false);

  // Código de Google Apps Script único y consolidado para copiar
  const appsScriptCodeConsolidated = `// Código único y consolidado para Google Sheets: Miembros, Candidatos y Votos
function doGet(e) {
  if (typeof e === "undefined") {
    return ContentService.createTextOutput("Servicio de Votaciones ICC activo. Realiza una 'Nueva implementación' como 'Aplicación web'.").setMimeType(ContentService.MimeType.TEXT);
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Inicializar hojas si no existen
  var membersSheet = sheet.getSheetByName("Miembros") || sheet.getSheetByName("Votantes");
  if (!membersSheet) {
    membersSheet = sheet.insertSheet("Miembros");
    membersSheet.appendRow(["ID", "Nombre", "Apellido", "Has Voted", "Fecha Voto", "Presente"]);
  }
  
  var candidatesSheet = sheet.getSheetByName("Candidatos");
  if (!candidatesSheet) {
    candidatesSheet = sheet.insertSheet("Candidatos");
    candidatesSheet.appendRow(["ID", "Nombre", "Apellido", "Testimonio", "Foto"]);
  } else {
    // Asegurar columna Foto
    var headers = candidatesSheet.getRange(1, 1, 1, candidatesSheet.getLastColumn()).getValues()[0];
    if (headers.indexOf("Foto") === -1) {
      candidatesSheet.getRange(1, candidatesSheet.getLastColumn() + 1).setValue("Foto");
    }
  }
  
  if (!sheet.getSheetByName("Votos")) {
    var newSheet = sheet.insertSheet("Votos");
    newSheet.appendRow(["ID Candidato", "Nombre Candidato", "Apellido Candidato", "Estado", "Motivo", "ID Votante", "Timestamp"]);
  }

  // Asegurar columnas Has Voted, Fecha Voto y Presente en hoja Miembros
  var membersHeaders = membersSheet.getRange(1, 1, 1, membersSheet.getLastColumn()).getValues()[0];
  if (membersHeaders.indexOf("Has Voted") === -1) {
    membersSheet.getRange(1, membersSheet.getLastColumn() + 1).setValue("Has Voted");
  }
  if (membersHeaders.indexOf("Fecha Voto") === -1) {
    membersSheet.getRange(1, membersSheet.getLastColumn() + 1).setValue("Fecha Voto");
  }
  if (membersHeaders.indexOf("Presente") === -1) {
    membersSheet.getRange(1, membersSheet.getLastColumn() + 1).setValue("Presente");
  }

  var action = e.parameter.action;
  
  if (action === "getData") {
    var updatedMembersSheet = sheet.getSheetByName("Miembros") || sheet.getSheetByName("Votantes");
    return ContentService.createTextOutput(JSON.stringify({
      voters: getSheetData(updatedMembersSheet),
      candidates: getSheetData(sheet.getSheetByName("Candidatos")),
      votes: getSheetData(sheet.getSheetByName("Votos"))
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({status: "success", msg: "Apps Script Conectado"})).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  if (typeof e === "undefined" || !e.postData || !e.postData.contents) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", msg: "No se recibieron datos"})).setMimeType(ContentService.MimeType.JSON);
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var membersSheet = sheet.getSheetByName("Miembros") || sheet.getSheetByName("Votantes");
  
  // Asegurar hojas creadas
  if (!membersSheet) {
    membersSheet = sheet.insertSheet("Miembros");
    membersSheet.appendRow(["ID", "Nombre", "Apellido", "Has Voted", "Fecha Voto", "Presente"]);
  }
  
  var candidatesSheet = sheet.getSheetByName("Candidatos");
  if (!candidatesSheet) {
    candidatesSheet = sheet.insertSheet("Candidatos");
    candidatesSheet.appendRow(["ID", "Nombre", "Apellido", "Testimonio", "Foto"]);
  } else {
    // Asegurar columna Foto
    var headers = candidatesSheet.getRange(1, 1, 1, candidatesSheet.getLastColumn()).getValues()[0];
    if (headers.indexOf("Foto") === -1) {
      candidatesSheet.getRange(1, candidatesSheet.getLastColumn() + 1).setValue("Foto");
    }
  }
  
  if (!sheet.getSheetByName("Votos")) {
    var newSheet = sheet.insertSheet("Votos");
    newSheet.appendRow(["ID Candidato", "Nombre Candidato", "Apellido Candidato", "Estado", "Motivo", "ID Votante", "Timestamp"]);
  }

  // Asegurar columnas Has Voted, Fecha Voto y Presente en hoja Miembros
  var membersSheet2 = sheet.getSheetByName("Miembros") || sheet.getSheetByName("Votantes");
  if (membersSheet2) {
    var mHeaders = membersSheet2.getRange(1, 1, 1, membersSheet2.getLastColumn()).getValues()[0];
    if (mHeaders.indexOf("Has Voted") === -1) {
      membersSheet2.getRange(1, membersSheet2.getLastColumn() + 1).setValue("Has Voted");
    }
    if (mHeaders.indexOf("Fecha Voto") === -1) {
      membersSheet2.getRange(1, membersSheet2.getLastColumn() + 1).setValue("Fecha Voto");
    }
    if (mHeaders.indexOf("Presente") === -1) {
      membersSheet2.getRange(1, membersSheet2.getLastColumn() + 1).setValue("Presente");
    }
  }

  var data = JSON.parse(e.postData.contents);
  var action = data.action;

  if (action === "updateVoters") {
    membersSheet.clear();
    membersSheet.appendRow(["ID", "Nombre", "Apellido", "Has Voted", "Fecha Voto", "Presente"]);
    data.voters.forEach(function(v) {
      membersSheet.appendRow([
        v.id || v.ID || "", 
        v.name || v.Nombre || "", 
        v.lastName || v.Apellido || "",
        v.hasVoted === true ? true : false,
        v.votedAt || "",
        v.isPresent !== false ? true : false
      ]);
    });
    return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "addCandidate") {
    var candidatesSheet = sheet.getSheetByName("Candidatos");
    var headers = candidatesSheet.getRange(1, 1, 1, candidatesSheet.getLastColumn()).getValues()[0];
    var fotoColIdx = headers.indexOf("Foto") + 1;
    if (fotoColIdx === 0) {
      candidatesSheet.getRange(1, headers.length + 1).setValue("Foto");
      fotoColIdx = headers.length + 1;
    }

    var photoUrl = "";
    if (data.photoBase64 && data.photoName) {
      photoUrl = saveFileToDrive(data.photoBase64, data.photoName);
    }

    // Asegurar 5 columnas alineadas con el appendRow
    var newRow = [
      data.id || "",
      data.firstName || "",
      data.lastName || "",
      data.testimony || ""
    ];
    // Rellenar hasta llegar a la columna de Foto
    while (newRow.length < fotoColIdx - 1) {
      newRow.push("");
    }
    newRow[fotoColIdx - 1] = photoUrl;

    candidatesSheet.appendRow(newRow);
    return ContentService.createTextOutput(JSON.stringify({status: "success", photoUrl: photoUrl})).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "updateCandidate") {
    var candidatesSheet = sheet.getSheetByName("Candidatos");
    var headers = candidatesSheet.getRange(1, 1, 1, candidatesSheet.getLastColumn()).getValues()[0];
    var fotoColIdx = headers.indexOf("Foto") + 1;
    if (fotoColIdx === 0) {
      candidatesSheet.getRange(1, headers.length + 1).setValue("Foto");
      fotoColIdx = headers.length + 1;
    }

    var rows = candidatesSheet.getDataRange().getValues();
    var targetId = String(data.id).trim();
    var targetIdInt = parseInt(targetId, 10);
    for (var i = 1; i < rows.length; i++) {
      var rowId = String(rows[i][0]).trim();
      var rowIdInt = parseInt(rowId, 10);
      var isMatch = rowId === targetId || (!isNaN(rowIdInt) && !isNaN(targetIdInt) && rowIdInt === targetIdInt);
      if (isMatch) {
        candidatesSheet.getRange(i + 1, 2).setValue(data.firstName || "");
        candidatesSheet.getRange(i + 1, 3).setValue(data.lastName || "");
        candidatesSheet.getRange(i + 1, 4).setValue(data.testimony || "");
        
        if (data.photoBase64 && data.photoName) {
          // Eliminar foto anterior si existe
          var oldPhotoUrl = candidatesSheet.getRange(i + 1, fotoColIdx).getValue();
          deleteFileFromDrive(oldPhotoUrl);

          var photoUrl = saveFileToDrive(data.photoBase64, data.photoName);
          candidatesSheet.getRange(i + 1, fotoColIdx).setValue(photoUrl);
        }
        return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "error", msg: "Candidato no encontrado"})).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "clearCandidates") {
    var candidatesSheet = sheet.getSheetByName("Candidatos");
    var headers = candidatesSheet.getRange(1, 1, 1, candidatesSheet.getLastColumn()).getValues()[0];
    var fotoColIdx = headers.indexOf("Foto") + 1;
    if (fotoColIdx > 0) {
      var rows = candidatesSheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        var photoUrl = rows[i][fotoColIdx - 1];
        deleteFileFromDrive(photoUrl);
      }
    }
    candidatesSheet.clear();
    candidatesSheet.appendRow(["ID", "Nombre", "Apellido", "Testimonio", "Foto"]);
    return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "deleteCandidate") {
    var candidatesSheet = sheet.getSheetByName("Candidatos");
    var headers = candidatesSheet.getRange(1, 1, 1, candidatesSheet.getLastColumn()).getValues()[0];
    var fotoColIdx = headers.indexOf("Foto") + 1;
    var rows = candidatesSheet.getDataRange().getValues();
    var targetId = String(data.id).trim();
    var targetIdInt = parseInt(targetId, 10);
    for (var i = 1; i < rows.length; i++) {
      var rowId = String(rows[i][0]).trim();
      var rowIdInt = parseInt(rowId, 10);
      var isMatch = rowId === targetId || (!isNaN(rowIdInt) && !isNaN(targetIdInt) && rowIdInt === targetIdInt);
      if (isMatch) {
        if (fotoColIdx > 0) {
          var photoUrl = candidatesSheet.getRange(i + 1, fotoColIdx).getValue();
          deleteFileFromDrive(photoUrl);
        }
        candidatesSheet.deleteRow(i + 1);
        return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "error", msg: "Candidato no encontrado"})).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "clearVotes") {
    var votesSheet = sheet.getSheetByName("Votos");
    votesSheet.clear();
    votesSheet.appendRow(["ID Candidato", "Nombre Candidato", "Apellido Candidato", "Estado", "Motivo", "ID Votante", "Timestamp"]);
    // También limpiar Has Voted en todos los miembros
    var mSheet = sheet.getSheetByName("Miembros") || sheet.getSheetByName("Votantes");
    if (mSheet) {
      var mHeaders = mSheet.getRange(1, 1, 1, mSheet.getLastColumn()).getValues()[0];
      var hvCol = mHeaders.indexOf("Has Voted") + 1;
      if (hvCol > 0 && mSheet.getLastRow() > 1) {
        var numRows = mSheet.getLastRow() - 1;
        var hvRange = mSheet.getRange(2, hvCol, numRows, 1);
        hvRange.setValue(false);
        var fvCol = mHeaders.indexOf("Fecha Voto") + 1;
        if (fvCol > 0) mSheet.getRange(2, fvCol, numRows, 1).setValue("");
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "resetMultipleVotersVotes") {
    var votesSheet = sheet.getSheetByName("Votos");
    var mSheet = sheet.getSheetByName("Miembros") || sheet.getSheetByName("Votantes");
    var targetVoters = data.voters || [];
    
    // Construir set de IDs de votantes a resetear
    var targetIds = {};
    targetVoters.forEach(function(v) {
      if (v.voterId) targetIds[String(v.voterId).trim()] = true;
    });

    // Eliminar votos de la hoja Votos por voterId (columna 6)
    var vRows = votesSheet.getDataRange().getValues();
    for (var i = vRows.length - 1; i >= 1; i--) {
      var vid = String(vRows[i][5]).trim();
      if (targetIds[vid]) {
        votesSheet.deleteRow(i + 1);
      }
    }

    // Limpiar Has Voted en la hoja Miembros
    if (mSheet) {
      var mHeaders = mSheet.getRange(1, 1, 1, mSheet.getLastColumn()).getValues()[0];
      var idCol = mHeaders.indexOf("ID") + 1;
      var hvCol = mHeaders.indexOf("Has Voted") + 1;
      var fvCol = mHeaders.indexOf("Fecha Voto") + 1;
      if (hvCol > 0) {
        var mRows = mSheet.getDataRange().getValues();
        for (var j = 1; j < mRows.length; j++) {
          var rowId = String(mRows[j][idCol - 1]).trim();
          if (targetIds[rowId]) {
            mSheet.getRange(j + 1, hvCol).setValue(false);
            if (fvCol > 0) mSheet.getRange(j + 1, fvCol).setValue("");
          }
        }
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "addVote") {
    var votesSheet = sheet.getSheetByName("Votos");
    votesSheet.appendRow([
      data.candidateId,
      data.candidateFirstName,
      data.candidateLastName,
      data.status,
      data.reason || "",
      data.voterId || "",          // ID opaco del votante (VOTO SECRETO)
      new Date().toISOString()    // Timestamp
    ]);
    return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
  }

  // Marcar participación de un votante (VOTO SECRETO)
  if (action === "markParticipation") {
    var mSheet = sheet.getSheetByName("Miembros") || sheet.getSheetByName("Votantes");
    if (!mSheet) return ContentService.createTextOutput(JSON.stringify({status: "error", msg: "Hoja Miembros no encontrada"})).setMimeType(ContentService.MimeType.JSON);
    var mHeaders = mSheet.getRange(1, 1, 1, mSheet.getLastColumn()).getValues()[0];
    var idCol = mHeaders.indexOf("ID") + 1;
    var hvCol = mHeaders.indexOf("Has Voted") + 1;
    var fvCol = mHeaders.indexOf("Fecha Voto") + 1;
    if (hvCol === 0) {
      mSheet.getRange(1, mHeaders.length + 1).setValue("Has Voted");
      mSheet.getRange(1, mHeaders.length + 2).setValue("Fecha Voto");
      hvCol = mHeaders.length + 1;
      fvCol = mHeaders.length + 2;
    }
    var mRows = mSheet.getDataRange().getValues();
    var targetId = String(data.voterId).trim();
    for (var i = 1; i < mRows.length; i++) {
      if (String(mRows[i][idCol - 1]).trim() === targetId) {
        mSheet.getRange(i + 1, hvCol).setValue(true);
        if (fvCol > 0) mSheet.getRange(i + 1, fvCol).setValue(new Date().toISOString());
        return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "error", msg: "Votante no encontrado"})).setMimeType(ContentService.MimeType.JSON);
  }

  // Limpiar participación de un votante (admin)
  if (action === "resetParticipation") {
    var mSheet = sheet.getSheetByName("Miembros") || sheet.getSheetByName("Votantes");
    if (!mSheet) return ContentService.createTextOutput(JSON.stringify({status: "error", msg: "Hoja Miembros no encontrada"})).setMimeType(ContentService.MimeType.JSON);
    var mHeaders = mSheet.getRange(1, 1, 1, mSheet.getLastColumn()).getValues()[0];
    var idCol = mHeaders.indexOf("ID") + 1;
    var hvCol = mHeaders.indexOf("Has Voted") + 1;
    var fvCol = mHeaders.indexOf("Fecha Voto") + 1;
    var mRows = mSheet.getDataRange().getValues();
    var targetId = String(data.voterId).trim();
    for (var i = 1; i < mRows.length; i++) {
      if (String(mRows[i][idCol - 1]).trim() === targetId) {
        if (hvCol > 0) mSheet.getRange(i + 1, hvCol).setValue(false);
        if (fvCol > 0) mSheet.getRange(i + 1, fvCol).setValue("");
        return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({status: "error", msg: "Acción no reconocida"})).setMimeType(ContentService.MimeType.JSON);
}

function saveFileToDrive(base64Data, fileName) {
  try {
    var splitData = base64Data.split(",");
    var contentType = splitData[0].match(/:(.*?);/)[1];
    var rawData = splitData[1];
    
    var decoded = Utilities.base64Decode(rawData);
    var blob = Utilities.newBlob(decoded, contentType, fileName);
    
    var folderName = "Votaciones_ICC_Fotos";
    var folders = DriveApp.getFoldersByName(folderName);
    var folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }
    
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return "https://lh3.googleusercontent.com/d/" + file.getId();
  } catch (e) {
    Logger.log("Error saving file to Drive: " + e.toString());
    return "";
  }
}

function deleteFileFromDrive(fileUrl) {
  try {
    if (!fileUrl) return;
    var fileId = "";
    if (fileUrl.indexOf("id=") !== -1) {
      fileId = fileUrl.split("id=")[1].split("&")[0];
    } else if (fileUrl.indexOf("/d/") !== -1) {
      fileId = fileUrl.split("/d/")[1].split("/")[0];
    } else if (fileUrl.indexOf("googleusercontent.com/d/") !== -1) {
      fileId = fileUrl.split("googleusercontent.com/d/")[1];
    }
    
    if (fileId) {
      var file = DriveApp.getFileById(fileId);
      file.setTrashed(true);
    }
  } catch (e) {
    Logger.log("Error deleting file from Drive: " + e.toString());
  }
}

function getSheetData(sheet) {
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var headers = rows[0];
  var data = [];
  for (var i = 1; i < rows.length; i++) {
    var rowData = {};
    for (var j = 0; j < headers.length; j++) {
      rowData[headers[j]] = rows[i][j];
    }
    data.push(rowData);
  }
  return data;
}`;

  // URL dinámica del QR (ruta limpia amigable, eliminando /admin del path para apuntar a la ruta de votación)
  const cleanPathname = window.location.pathname.replace(/\/admin\/?$/, '').replace(/\/+$/, '');
  const hostUrl = window.location.origin + cleanPathname;
  const votingUrl = `${hostUrl}/voter`;
  
  const [qrCodeUrl, setQrCodeUrl] = useState(`https://api.qrserver.com/v1/create-qr-code/?size=250x250&data=${encodeURIComponent(votingUrl)}`);

  // Sincronizar el QR con la URL de votación cuando esta cambie
  useEffect(() => {
    setQrCodeUrl(`https://api.qrserver.com/v1/create-qr-code/?size=250x250&data=${encodeURIComponent(votingUrl)}`);
  }, [votingUrl]);

  // Si el servidor de QR primario (QRServer) falla (por bloqueos de DNS o adblockers), usar Google Charts o QuickChart como fallback
  const handleQrError = () => {
    if (qrCodeUrl.includes('qrserver.com')) {
      console.warn("Fallo al cargar QRServer. Intentando Google Charts API...");
      setQrCodeUrl(`https://chart.googleapis.com/chart?cht=qr&chs=250x250&chl=${encodeURIComponent(votingUrl)}`);
    } else if (qrCodeUrl.includes('googleapis.com')) {
      console.warn("Fallo al cargar Google Charts API. Intentando QuickChart API...");
      setQrCodeUrl(`https://quickchart.io/qr?size=250&text=${encodeURIComponent(votingUrl)}`);
    }
  };

  // (Manejadores de guardado y Apps Script para UI removidos de forma segura)

  // Convertir archivo de foto a Base64 con compresión para evitar exceder la cuota de localStorage
  const convertAndCompressPhoto = (file) => {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = (event) => {
        const img = new Image();
        img.src = event.target.result;
        img.onload = () => {
          try {
            const canvas = document.createElement('canvas');
            const maxDimension = 200; // Suficiente para un avatar de 140x140
            let width = img.width;
            let height = img.height;

            if (width > height) {
              if (width > maxDimension) {
                height = Math.round((height * maxDimension) / width);
                width = maxDimension;
              }
            } else {
              if (height > maxDimension) {
                width = Math.round((width * maxDimension) / height);
                height = maxDimension;
              }
            }

            canvas.width = width;
            canvas.height = height;
            const ctx = canvas.getContext('2d');
            ctx.drawImage(img, 0, 0, width, height);

            // Exportamos como JPEG con calidad 0.7 para obtener un tamaño ínfimo (~10-15KB)
            const compressedBase64 = canvas.toDataURL('image/jpeg', 0.7);
            resolve(compressedBase64);
          } catch (compressError) {
            console.error("Error al comprimir imagen, usando original:", compressError);
            resolve(event.target.result); // Fallback a la original si falla el canvas
          }
        };
        img.onerror = () => {
          resolve(event.target.result); // Fallback si no carga como imagen
        };
      };
      reader.onerror = () => {
        resolve(""); // Retornar vacío si no se puede leer
      };
    });
  };

  // Guardar en localStorage de forma segura controlando excepciones de cuota llena
  const safeSetLocalStorage = (key, value) => {
    try {
      localStorage.setItem(key, value);
      return true;
    } catch (e) {
      console.error(`Error al guardar en localStorage para la clave ${key}:`, e);
      return false;
    }
  };

  const handleStartEditCandidate = (cand) => {
    setEditingCandidateId(cand.id);
    setCandFirstName(cand.firstName || '');
    setCandLastName(cand.lastName || '');
    setCandTestimony(cand.testimony || '');
    setCandPhotoFile(null);
    setExpandedCandidateId(null);
    if (fileInputRef.current) fileInputRef.current.value = '';
    window.scrollTo({ top: 0, behavior: 'smooth' });
    showToast(`Editando a: ${cand.firstName} ${cand.lastName}`, "info");
  };

  const handleCancelEdit = () => {
    setEditingCandidateId(null);
    setCandFirstName('');
    setCandLastName('');
    setCandTestimony('');
    setCandPhotoFile(null);
    setExpandedCandidateId(null);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  // Enviar Nuevo Candidato o Guardar Cambios de Edición
  const handleAddCandidate = async (e) => {
    e.preventDefault();
    const fName = candFirstName.trim();
    const lName = candLastName.trim();
    if (!fName || !lName || !candTestimony.trim()) {
      showToast("Ingresa nombre, apellido y testimonio del candidato", "error");
      return;
    }

    setIsUploading(true);
    try {
      let photoUrl = '';
      let compressedPhotoBase64 = undefined;

      if (candPhotoFile) {
        compressedPhotoBase64 = await convertAndCompressPhoto(candPhotoFile);
      }
      
      if (editingCandidateId) {
        // Estamos EDITANDO un candidato existente
        const existingCand = candidates.find(c => c.id === editingCandidateId);
        
        if (candPhotoFile) {
          // Nueva foto
          photoUrl = compressedPhotoBase64;
          safeSetLocalStorage(`icc_photo_${editingCandidateId}`, photoUrl);
          showToast("Fotografía actualizada localmente para el candidato", "info");
        } else {
          // Mantener foto anterior
          photoUrl = existingCand?.photo || '';
        }

        const updatedCand = {
          id: editingCandidateId,
          firstName: fName,
          lastName: lName,
          photo: photoUrl,
          testimony: candTestimony.trim(),
          createdAt: existingCand?.createdAt || new Date().toISOString()
        };

        if (config.sheetUrlVoters) {
          if (!isConnected) {
            throw new Error("No hay conexión con Google Sheets. No se puede actualizar el candidato.");
          }
          await updateCandidateInSheets(config.sheetUrlVoters, {
            id: editingCandidateId,
            firstName: fName,
            lastName: lName,
            testimony: candTestimony.trim(),
            photoBase64: compressedPhotoBase64,
            photoName: candPhotoFile ? candPhotoFile.name : undefined
          });
          
          showToast("Candidato actualizado en Google Sheets", "success");
          refreshData();
        } else {
          const updatedList = candidates.map(c => c.id === editingCandidateId ? updatedCand : c);
          setCandidates(updatedList);
          safeSetLocalStorage('icc_local_candidates', JSON.stringify(updatedList));
          showToast("Candidato actualizado localmente con éxito", "success");
        }
        
        // Limpiar estado de edición
        setEditingCandidateId(null);
      } else {
        // Creando NUEVO candidato
        const numericIds = candidates.map(c => {
          const num = parseInt(c.id, 10);
          return isNaN(num) ? 0 : num;
        });
        const nextNum = numericIds.length > 0 ? Math.max(...numericIds) + 1 : 1;
        const nextIdStr = String(nextNum).padStart(5, '0');

        if (candPhotoFile) {
          photoUrl = compressedPhotoBase64;
          safeSetLocalStorage(`icc_photo_${nextIdStr}`, photoUrl);
          showToast("Fotografía guardada localmente asociada al ID del candidato", "info");
        } else {
          photoUrl = '';
        }

        localStorage.removeItem('icc_candidates_cleared_by_user');
        if (config.sheetUrlVoters) {
          if (!isConnected) {
            throw new Error("No hay conexión con Google Sheets. No se puede agregar el candidato.");
          }
          await addCandidateToSheets(config.sheetUrlVoters, {
            id: nextIdStr,
            firstName: fName,
            lastName: lName,
            testimony: candTestimony.trim(),
            photoBase64: compressedPhotoBase64,
            photoName: candPhotoFile ? candPhotoFile.name : undefined
          });
          showToast("Candidato registrado en Google Sheets", "success");
          refreshData();
        } else {
          if (candPhotoFile) {
            photoUrl = compressedPhotoBase64;
            safeSetLocalStorage(`icc_photo_${nextIdStr}`, photoUrl);
          }
          const newCand = {
            id: nextIdStr,
            firstName: fName,
            lastName: lName,
            photo: photoUrl,
            testimony: candTestimony.trim(),
            createdAt: new Date().toISOString()
          };
          const updated = [...candidates, newCand];
          setCandidates(updated);
          safeSetLocalStorage('icc_local_candidates', JSON.stringify(updated));
          showToast("Candidato registrado localmente con éxito", "success");
        }
      }

      // Limpiar Formulario
      setCandFirstName('');
      setCandLastName('');
      setCandTestimony('');
      setCandPhotoFile(null);
      if (fileInputRef.current) fileInputRef.current.value = '';

    } catch (error) {
      console.error(error);
      showToast(`Error al guardar candidato: ${error.message}`, "error");
    } finally {
      setIsUploading(false);
    }
  };

  // Eliminar Candidato
  const handleDeleteCandidate = async (id) => {
    if (!window.confirm("¿Estás seguro de eliminar a este candidato? Esto eliminará sus registros.")) return;
    setExpandedCandidateId(null);
    try {
      if (config.sheetUrlVoters) {
        if (!isConnected) {
          throw new Error("No hay conexión con Google Sheets. No se puede eliminar el candidato.");
        }
        await deleteCandidateFromSheets(config.sheetUrlVoters, id);
        localStorage.removeItem(`icc_photo_${id}`);
        showToast("Candidato eliminado de Google Sheets", "success");
        refreshData();
      } else {
        const updated = candidates.filter(c => c.id !== id);
        setCandidates(updated);
        safeSetLocalStorage('icc_local_candidates', JSON.stringify(updated));
        localStorage.removeItem(`icc_photo_${id}`);
        // Limpiar también los votos asociados localmente
        const updatedVotes = votes.filter(v => v.candidateId !== id);
        setVotes(updatedVotes);
        safeSetLocalStorage('icc_local_votes', JSON.stringify(updatedVotes));
        showToast("Candidato eliminado localmente", "success");
      }
    } catch (error) {
      console.error(error);
      showToast(`Error al eliminar: ${error.message}`, "error");
    }
  };

  // Limpiar lista de candidatos
  const handleClearAllCandidates = async () => {
    if (!window.confirm("¿Seguro que deseas eliminar a TODOS los candidatos de la base de datos?")) return;
    setIsUploading(true);
    try {
      // Eliminar fotos asociadas a los candidatos en localStorage siempre
      candidates.forEach(c => {
        localStorage.removeItem(`icc_photo_${c.id}`);
      });

      if (config.sheetUrlVoters) {
        if (!isConnected) {
          throw new Error("No hay conexión con Google Sheets. No se pueden eliminar los candidatos.");
        }
        showToast("Eliminando candidatos de Google Sheets...", "info");
        await clearCandidatesInSheets(config.sheetUrlVoters);
        safeSetLocalStorage('icc_candidates_cleared_by_user', 'true');
        showToast("Todos los candidatos eliminados de Google Sheets", "success");
        refreshData();
      } else {
        setCandidates([]);
        safeSetLocalStorage('icc_local_candidates', JSON.stringify([]));
        safeSetLocalStorage('icc_candidates_cleared_by_user', 'true');
        
        // Limpiar también los votos asociados localmente
        const updatedVotes = votes.filter(v => !candidates.some(c => c.id === v.candidateId));
        setVotes(updatedVotes);
        safeSetLocalStorage('icc_local_votes', JSON.stringify(updatedVotes));
        showToast("Lista de candidatos vaciada localmente", "success");
      }
    } catch (error) {
      console.error(error);
      showToast("Error al vaciar candidatos: " + error.message, "error");
    } finally {
      setIsUploading(false);
    }
  };

  // Agregar un solo Votante
  const handleAddSingleVoter = async (e) => {
    e.preventDefault();
    const fName = voterFirstName.trim();
    const lName = voterLastName.trim();
    if (!fName || !lName) {
      showToast("Ingresa nombre y apellido del votante", "error");
      return;
    }

    const voterExists = voters.some(v => 
      v.name?.toLowerCase() === fName.toLowerCase() && 
      v.lastName?.toLowerCase() === lName.toLowerCase()
    );

    if (voterExists) {
      showToast("Este miembro ya está en la lista de votantes", "error");
      return;
    }

    const numericIds = voters.map(v => {
      const num = parseInt(v.id, 10);
      return isNaN(num) ? 0 : num;
    });
    const nextNum = numericIds.length > 0 ? Math.max(...numericIds) + 1 : 1;
    const nextIdStr = String(nextNum).padStart(5, '0');

    const newVoter = {
      id: nextIdStr,
      name: fName,
      lastName: lName,
      isPresent: true
    };

    const updatedVoters = [...voters, newVoter];

    try {
      if (config.sheetUrlVoters) {
        if (!isConnected) {
          throw new Error("No hay conexión con Google Sheets. No se puede agregar el votante.");
        }
        await updateVotersInSheets(config.sheetUrlVoters, updatedVoters);
        setVoters(updatedVoters);
        showToast("Votante registrado en Google Sheets", "success");
      } else {
        setVoters(updatedVoters);
        localStorage.setItem('icc_local_voters', JSON.stringify(updatedVoters));
        showToast("Votante registrado localmente", "success");
      }
      setVoterFirstName('');
      setVoterLastName('');
    } catch (error) {
      console.error(error);
      showToast(`Error: ${error.message}`, "error");
    }
  };

  // Agregar Votantes en lote
  const handleAddBulkVoters = async (e) => {
    e.preventDefault();
    if (!bulkVotersText.trim()) return;

    // Dividir por saltos de línea o comas y limpiar espacios vacíos
    const newNames = bulkVotersText
      .split(/[\n,]+/)
      .map(name => name.trim())
      .filter(name => name.length > 0);

    if (newNames.length === 0) return;

    const uniqueNewVoters = [];
    newNames.forEach((fullName) => {
      const parts = fullName.split(' ');
      const name = parts[0] || '';
      const lastName = parts.slice(1).join(' ') || '';

      const voterExists = voters.some(v => 
        v.name?.toLowerCase() === name.toLowerCase() && 
        v.lastName?.toLowerCase() === lastName.toLowerCase()
      ) || uniqueNewVoters.some(v => 
        v.name?.toLowerCase() === name.toLowerCase() && 
        v.lastName?.toLowerCase() === lastName.toLowerCase()
      );

      if (!voterExists && name) {
        const currentList = [...voters, ...uniqueNewVoters];
        const numericIds = currentList.map(v => {
          const num = parseInt(v.id, 10);
          return isNaN(num) ? 0 : num;
        });
        const nextNum = numericIds.length > 0 ? Math.max(...numericIds) + 1 : 1;
        const nextIdStr = String(nextNum).padStart(5, '0');

        uniqueNewVoters.push({
          id: nextIdStr,
          name,
          lastName,
          isPresent: true
        });
      }
    });

    if (uniqueNewVoters.length === 0) {
      showToast("Todos los nombres ingresados ya existen en la lista", "error");
      return;
    }

    const updatedVoters = [...voters, ...uniqueNewVoters];

    try {
      if (config.sheetUrlVoters) {
        if (!isConnected) {
          throw new Error("No hay conexión con Google Sheets. No se pueden guardar los votantes.");
        }
        await updateVotersInSheets(config.sheetUrlVoters, updatedVoters);
        showToast(`${uniqueNewVoters.length} votantes guardados en Google Sheets`, "success");
        refreshData();
      } else {
        setVoters(updatedVoters);
        localStorage.setItem('icc_local_voters', JSON.stringify(updatedVoters));
        showToast(`${uniqueNewVoters.length} votantes guardados localmente`, "success");
      }
      setBulkVotersText('');
    } catch (error) {
      console.error(error);
      showToast(`Error al guardar lote: ${error.message}`, "error");
    }
  };

  // Eliminar un votante
  const handleDeleteVoter = async (voterId) => {
    const updatedVoters = voters.filter(v => v.id !== voterId);
    try {
      if (config.sheetUrlVoters) {
        if (!isConnected) {
          throw new Error("No hay conexión con Google Sheets. No se puede eliminar el votante.");
        }
        await updateVotersInSheets(config.sheetUrlVoters, updatedVoters);
        setVoters(updatedVoters);
        showToast("Votante eliminado", "success");
      } else {
        setVoters(updatedVoters);
        localStorage.setItem('icc_local_voters', JSON.stringify(updatedVoters));
        showToast("Votante eliminado localmente", "success");
      }
    } catch (error) {
      console.error(error);
      showToast(`Error al eliminar votante: ${error.message}`, "error");
    }
  };

  const handleToggleVoterPresence = async (voter, isPresent) => {
    // 1. Guardar en el fallback local usando el nombre completo como clave
    try {
      const stored = localStorage.getItem('icc_voter_attendance');
      const attendanceMap = stored ? JSON.parse(stored) : {};
      const nameKey = `${voter.name || ''} ${voter.lastName || ''}`.trim().toLowerCase();
      attendanceMap[nameKey] = isPresent;
      localStorage.setItem('icc_voter_attendance', JSON.stringify(attendanceMap));
    } catch (e) {
      console.error("Error saving attendance to localStorage fallback:", e);
    }

    // 2. Actualizar el estado de React inmediatamente de forma optimista
    const updatedVoters = voters.map(v => 
      String(v.id) === String(voter.id) ? { ...v, isPresent } : v
    );
    setVoters(updatedVoters);
    
    // 3. Persistir en servidor o almacenamiento local
    try {
      if (config.sheetUrlVoters) {
        if (!isConnected) {
          throw new Error("No hay conexión con Google Sheets. No se puede guardar la asistencia.");
        }
        await updateVotersInSheets(config.sheetUrlVoters, updatedVoters);
        showToast(`Asistencia de ${voter.name} actualizada`, "success");
      } else {
        localStorage.setItem('icc_local_voters', JSON.stringify(updatedVoters));
        showToast(`Asistencia de ${voter.name} actualizada localmente`, "success");
      }
    } catch (error) {
      console.error(error);
      showToast(`Error al guardar asistencia: ${error.message}`, "error");
      // Revertir en caso de error
      setVoters(voters);
    }
  };

  const handleBulkTogglePresence = async (isPresent) => {
    if (selectedVoterIds.length === 0) return;
    
    // 1. Guardar en el fallback local usando el nombre completo de cada votante como clave
    try {
      const stored = localStorage.getItem('icc_voter_attendance');
      const attendanceMap = stored ? JSON.parse(stored) : {};
      selectedVoterIds.forEach(id => {
        const voter = voters.find(v => String(v.id) === String(id));
        if (voter) {
          const nameKey = `${voter.name || ''} ${voter.lastName || ''}`.trim().toLowerCase();
          attendanceMap[nameKey] = isPresent;
        }
      });
      localStorage.setItem('icc_voter_attendance', JSON.stringify(attendanceMap));
    } catch (e) {
      console.error("Error saving bulk attendance to localStorage fallback:", e);
    }

    // 2. Actualizar el estado de React inmediatamente de forma optimista
    const updatedVoters = voters.map(v => 
      selectedVoterIds.some(id => String(id) === String(v.id)) ? { ...v, isPresent } : v
    );
    setVoters(updatedVoters);
    
    setIsUploading(true);
    try {
      if (config.sheetUrlVoters) {
        if (!isConnected) {
          throw new Error("No hay conexión con Google Sheets.");
        }
        await updateVotersInSheets(config.sheetUrlVoters, updatedVoters);
        showToast("Asistencia masiva registrada en Google Sheets", "success");
      } else {
        localStorage.setItem('icc_local_voters', JSON.stringify(updatedVoters));
        showToast("Asistencia masiva registrada localmente", "success");
      }
      setSelectedVoterIds([]);
    } catch (error) {
      console.error(error);
      showToast(`Error al actualizar asistencia: ${error.message}`, "error");
      // Revertir en caso de error
      setVoters(voters);
    } finally {
      setIsUploading(false);
    }
  };

  // Limpiar lista de votantes
  const handleClearAllVoters = async () => {
    if (!window.confirm("¿Seguro que deseas eliminar a TODOS los votantes?")) return;
    try {
      if (config.sheetUrlVoters) {
        if (!isConnected) {
          throw new Error("No hay conexión con Google Sheets. No se puede vaciar la lista de votantes.");
        }
        await updateVotersInSheets(config.sheetUrlVoters, []);
        setVoters([]);
        showToast("Lista de votantes vaciada", "success");
      } else {
        setVoters([]);
        safeSetLocalStorage('icc_local_voters', JSON.stringify([]));
        showToast("Lista de votantes vaciada localmente", "success");
      }
    } catch (error) {
      console.error(error);
      showToast("Error al vaciar lista", "error");
    }
  };

  // Obtener el progreso de votación de un miembro (VOTO SECRETO)
  const getVoterProgress = (voter) => {
    if (!voter) {
      return {
        votedCount: 0,
        totalCount: candidates.length,
        hasVoted: false,
        isComplete: false,
        pendingCandidates: [...candidates]
      };
    }

    // Filtrar los votos que coinciden con el voterId del miembro (soportando llaves locales y de Google Sheets)
    const voterVotes = votes.filter(v => {
      const vId = String(v.voterId || v.ID_Votante || v["ID Votante"] || '').trim();
      const voterId = String(voter.id).trim();
      return vId === voterId;
    });
    const votedCount = voterVotes.length;
    const totalCount = candidates.length;
    
    // Obtener los IDs de candidatos por los que ya votó (soportando llaves locales y de Google Sheets)
    const votedCandIds = new Set(voterVotes.map(v => String(v.candidateId || v.ID_Candidato || v["ID Candidato"] || '').trim()));
    const pendingCandidates = candidates.filter(c => !votedCandIds.has(String(c.id).trim()));

    // Consideramos que ya votó por completo si los votos coinciden con el total de candidatos
    const isComplete = votedCount >= totalCount && totalCount > 0;
    
    // Consideramos que ha participado si su hasVoted es true o si ya tiene al menos un voto registrado
    const participated = voter.hasVoted === true || votedCount > 0;

    return {
      votedCount,
      totalCount,
      hasVoted: participated,
      isComplete,
      pendingCandidates
    };
  };

  // Restaurar los votos de un votante para permitirle volver a votar (VOTO SECRETO)
  const handleResetVoterVotes = async (voter) => {
    const fullName = `${voter.name} ${voter.lastName}`;
    if (!window.confirm(`¿Estás seguro de que deseas restablecer los votos de ${fullName}? Esto le permitirá volver a votar.`)) {
      return;
    }

    setIsUploading(true);
    try {
      if (config.sheetUrlCandidates) {
        if (!isConnected) {
          throw new Error("No hay conexión con Google Sheets. No se pueden restablecer los votos.");
        }
        await resetVoterVotesInSheets(config.sheetUrlCandidates, { voterId: voter.id });
        
        // Actualizar localmente los votantes y los votos
        const targetVoterId = String(voter.id).trim();
        const updatedVoters = voters.map(v => 
          String(v.id).trim() === targetVoterId ? { ...v, hasVoted: false, votedAt: '' } : v
        );
        setVoters(updatedVoters);

        const updatedVotes = votes.filter(v => String(v.voterId || '').trim() !== targetVoterId);
        setVotes(updatedVotes);

        showToast(`Participación de ${fullName} restablecida en Google Sheets`, "success");
      } else {
        // Limpiar votos locales por voterId
        const updatedVotes = votes.filter(v => String(v.voterId || '').trim() !== String(voter.id).trim());
        setVotes(updatedVotes);
        safeSetLocalStorage('icc_local_votes', JSON.stringify(updatedVotes));
        // Limpiar hasVoted local
        const localVoters = localStorage.getItem('icc_local_voters');
        if (localVoters) {
          const parsed = JSON.parse(localVoters);
          const updated = parsed.map(v => v.id === voter.id ? { ...v, hasVoted: false, votedAt: '' } : v);
          localStorage.setItem('icc_local_voters', JSON.stringify(updated));
        }
        showToast(`Participación de ${fullName} restablecida localmente`, "success");
      }
    } catch (error) {
      console.error(error);
      showToast(`Error al restablecer votos: ${error.message}`, "error");
    } finally {
      setIsUploading(false);
    }
  };

  // Vaciar todos los votos registrados (Reiniciar la Votación general)
  const handleClearAllVotes = async () => {
    if (!window.confirm("¿Seguro que deseas eliminar TODOS los votos registrados? Esto reiniciará la votación para todos los votantes, pero conservará los candidatos y la lista de votantes intactos.")) {
      return;
    }

    setIsUploading(true);
    try {
      if (config.sheetUrlCandidates) {
        if (!isConnected) {
          throw new Error("No hay conexión con Google Sheets. No se pueden eliminar los votos.");
        }
        await clearVotesInSheets(config.sheetUrlCandidates);
        showToast("Todos los votos han sido eliminados de Google Sheets", "success");
        refreshData();
      } else {
        setVotes([]);
        safeSetLocalStorage('icc_local_votes', JSON.stringify([]));
        showToast("Todos los votos han sido eliminados localmente", "success");
      }
    } catch (error) {
      console.error(error);
      showToast("Error al vaciar los votos: " + error.message, "error");
    } finally {
      setIsUploading(false);
    }
  };

  // Alternar la selección de todos los votantes visibles
  const handleToggleSelectAll = (allFilteredSelected) => {
    if (allFilteredSelected) {
      const filteredIdsStr = new Set(filteredVoters.map(v => String(v.id)));
      setSelectedVoterIds(selectedVoterIds.filter(id => !filteredIdsStr.has(String(id))));
    } else {
      const filteredIdsStr = filteredVoters.map(v => String(v.id));
      const newSelected = Array.from(new Set([...selectedVoterIds.map(id => String(id)), ...filteredIdsStr]));
      setSelectedVoterIds(newSelected);
    }
  };

  // Restablecer votos de múltiples votantes en lote (VOTO SECRETO)
  const handleBulkResetVotes = async () => {
    const selectedVoterIdsSet = new Set(selectedVoterIds.map(id => String(id)));
    const selectedVoters = voters.filter(v => selectedVoterIdsSet.has(String(v.id)));
    if (selectedVoters.length === 0) return;

    if (!window.confirm(`¿Estás seguro de que deseas restablecer la participación de los ${selectedVoters.length} votantes seleccionados?`)) {
      return;
    }

    setIsUploading(true);
    try {
      if (config.sheetUrlCandidates) {
        if (!isConnected) {
          throw new Error("No hay conexión con Google Sheets. No se pueden restablecer los votos.");
        }
        const payloadList = selectedVoters.map(v => ({ voterId: v.id }));
        await resetMultipleVotersVotesInSheets(config.sheetUrlCandidates, payloadList);
        
        // Actualizar localmente
        const targetIds = new Set(selectedVoters.map(v => String(v.id).trim()));
        const updatedVoters = voters.map(v => 
          targetIds.has(String(v.id).trim()) ? { ...v, hasVoted: false, votedAt: '' } : v
        );
        setVoters(updatedVoters);

        const updatedVotes = votes.filter(v => !targetIds.has(String(v.voterId || '').trim()));
        setVotes(updatedVotes);

        showToast(`Participación de ${selectedVoters.length} miembros restablecida en Google Sheets`, "success");
      } else {
        const targetIds = new Set(selectedVoters.map(v => String(v.id).trim()));
        // Limpiar votos locales por voterId
        const updatedVotes = votes.filter(v => !targetIds.has(String(v.voterId || '').trim()));
        setVotes(updatedVotes);
        safeSetLocalStorage('icc_local_votes', JSON.stringify(updatedVotes));
        // Limpiar hasVoted local
        const localVoters = localStorage.getItem('icc_local_voters');
        if (localVoters) {
          const parsed = JSON.parse(localVoters);
          const updated = parsed.map(v => targetIds.has(String(v.id).trim()) ? { ...v, hasVoted: false, votedAt: '' } : v);
          localStorage.setItem('icc_local_voters', JSON.stringify(updated));
        }
        showToast(`Participación de ${selectedVoters.length} miembros restablecida localmente`, "success");
      }
      setSelectedVoterIds([]);
    } catch (error) {
      console.error(error);
      showToast(`Error al restablecer votos: ${error.message}`, "error");
    } finally {
      setIsUploading(false);
    }
  };

  // Eliminar múltiples votantes en lote
  const handleBulkDeleteVoters = async () => {
    const idsToDelete = new Set(selectedVoterIds.map(id => String(id)));
    if (idsToDelete.size === 0) return;

    if (!window.confirm(`¿Estás seguro de que deseas eliminar a los ${idsToDelete.size} votantes seleccionados del censo?`)) {
      return;
    }

    setIsUploading(true);
    const updatedVoters = voters.filter(v => !idsToDelete.has(String(v.id)));
    try {
      if (config.sheetUrlVoters) {
        if (!isConnected) {
          throw new Error("No hay conexión con Google Sheets. No se pueden eliminar los votantes.");
        }
        await updateVotersInSheets(config.sheetUrlVoters, updatedVoters);
        setVoters(updatedVoters);
        showToast(`${idsToDelete.size} votantes eliminados de Google Sheets`, "success");
      } else {
        setVoters(updatedVoters);
        safeSetLocalStorage('icc_local_voters', JSON.stringify(updatedVoters));
        showToast(`${idsToDelete.size} votantes eliminados localmente`, "success");
      }
      setSelectedVoterIds([]); // Limpiar selección
    } catch (error) {
      console.error(error);
      showToast(`Error al eliminar votantes: ${error.message}`, "error");
    } finally {
      setIsUploading(false);
    }
  };

  // Calcular estadísticas por candidato
  const getCandidateStats = (candId) => {
    const candVotes = votes.filter(v => {
      const vId = String(v.candidateId || v.ID_Candidato || v["ID Candidato"] || '').trim();
      const cId = String(candId || '').trim();
      const vIdInt = parseInt(vId, 10);
      const cIdInt = parseInt(cId, 10);
      return vId === cId || (!isNaN(vIdInt) && !isNaN(cIdInt) && vIdInt === cIdInt);
    });
    
    // Contar aprobados y desaprobados
    let approves = 0;
    let disapproves = 0;
    const details = [];

    candVotes.forEach(vote => {
      // Soportar campos tanto de Google Sheets (en español/mayúsculas) como de LocalStorage
      const status = vote.status || vote.Estado || vote.estado;
      const voter = vote.voterName || vote.Nombre_Votante || vote.Nombre_Votante || vote["Nombre Votante"] || vote.votante;
      const reason = vote.reason || vote.Motivo || vote.motivo || '';
      const date = vote.timestamp || vote.Fecha || vote.fecha || '';

      const isApprove = status?.toLowerCase() === 'approve' || status?.toLowerCase() === 'aprueba' || status?.toLowerCase() === 'aprobar' || status === 'Aprueba';
      
      if (isApprove) {
        approves++;
      } else {
        disapproves++;
      }

      details.push({
        voter,
        status: isApprove ? 'Aprueba' : 'No Aprueba',
        reason,
        date
      });
    });

    const total = approves + disapproves;
    const approvalRate = total > 0 ? Math.round((approves / total) * 100) : 0;

    return {
      approves,
      disapproves,
      total,
      approvalRate,
      details
    };
  };

  // Exportar reporte completo a Excel
  const handleExportExcel = async () => {
    setIsExporting(true);
    try {
      const workbook = new ExcelJS.Workbook();
      workbook.creator = 'Votaciones ICC Admin';
      workbook.lastModifiedBy = 'Votaciones ICC Admin';
      workbook.created = new Date();
      workbook.modified = new Date();

      // ----------------------------------------------------
      // HOJA 1: Dashboard y Resultados
      // ----------------------------------------------------
      const wsDashboard = workbook.addWorksheet('Dashboard');
      wsDashboard.views = [{ showGridLines: true }];

      // Configurar anchos de columna para Dashboard (Columnas A a H)
      wsDashboard.columns = [
        { key: 'spacing', width: 4 }, // Columna A
        { key: 'name', width: 28 },    // Columna B
        { key: 'total', width: 14 },   // Columna C
        { key: 'app_cnt', width: 14 }, // Columna D
        { key: 'app_pct', width: 14 }, // Columna E
        { key: 'dis_cnt', width: 14 }, // Columna F
        { key: 'dis_pct', width: 14 }, // Columna G
        { key: 'rate', width: 18 },    // Columna H
      ];

      // Banner del Título Principal
      wsDashboard.mergeCells('B2:H2');
      for (let col = 2; col <= 8; col++) {
        const cell = wsDashboard.getCell(2, col);
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF0F172A' } // Slate 900
        };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FF0F172A' } },
          left: { style: 'thin', color: { argb: 'FF0F172A' } },
          bottom: { style: 'thin', color: { argb: 'FF0F172A' } },
          right: { style: 'thin', color: { argb: 'FF0F172A' } }
        };
      }
      const titleCell = wsDashboard.getCell('B2');
      titleCell.value = 'DASHBOARD DE RESULTADOS - VOTACIONES ICC';
      titleCell.font = { name: 'Segoe UI', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
      titleCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
      wsDashboard.getRow(2).height = 40;

      // Fecha de Reporte
      wsDashboard.mergeCells('B3:H3');
      const dateCell = wsDashboard.getCell('B3');
      dateCell.value = `Generado el: ${new Date().toLocaleString('es-ES')}`;
      dateCell.font = { name: 'Segoe UI', size: 9, italic: true, color: { argb: 'FF64748B' } };
      wsDashboard.getRow(3).height = 20;

      // Calcular participación y pendientes usando compatibilidad de IDs
      const participatedCount = voters.filter(v => {
        const voterVotes = votes.filter(vote => {
          const vId = String(vote.voterId || vote.ID_Votante || vote["ID Votante"] || '').trim();
          const voterId = String(v.id).trim();
          return vId === voterId;
        });
        return v.hasVoted === true || voterVotes.length > 0;
      }).length;
      
      const pendingCount = voters.length - participatedCount;
      const participationPct = voters.length > 0 ? (participatedCount / voters.length) * 100 : 0;
      const pendingPct = voters.length > 0 ? (pendingCount / voters.length) * 100 : 0;

      // Cards de Estadísticas Globales alineadas a la tabla (Columnas B a H)
      styleAndMergeCard(wsDashboard, 'B5:B6', 'Candidatos Activos', `${candidates.length}`);
      styleAndMergeCard(wsDashboard, 'C5:D6', 'Votos Registrados', `${votes.length}`);
      styleAndMergeCard(wsDashboard, 'E5:F6', 'Han Participado', `${Math.round(participationPct)}% (${participatedCount}/${voters.length})`);
      styleAndMergeCard(wsDashboard, 'G5:H6', 'Pendientes de Votar', `${Math.round(pendingPct)}% (${pendingCount}/${voters.length})`);

      wsDashboard.getRow(5).height = 24;
      wsDashboard.getRow(6).height = 24;

      // Título de la Tabla de Candidatos
      wsDashboard.getCell('B8').value = 'Resumen por Candidato';
      wsDashboard.getCell('B8').font = { name: 'Segoe UI', size: 12, bold: true, color: { argb: 'FF334155' } };
      wsDashboard.getRow(8).height = 24;

      // Tabla de Cabeceras
      const headerRow = wsDashboard.getRow(9);
      headerRow.height = 30;
      const headers = [
        { col: 'B', val: 'Candidato' },
        { col: 'C', val: 'Votos Totales' },
        { col: 'D', val: 'Aprueba (Cant)' },
        { col: 'E', val: 'Aprueba (%)' },
        { col: 'F', val: 'No Aprueba (Cant)' },
        { col: 'G', val: 'No Aprueba (%)' },
        { col: 'H', val: 'Tasa Aprobación' }
      ];

      headers.forEach(h => {
        const cell = wsDashboard.getCell(`${h.col}9`);
        cell.value = h.val;
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF1E293B' } // Slate 800
        };
        cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FF94A3B8' } },
          left: { style: 'thin', color: { argb: 'FF94A3B8' } },
          bottom: { style: 'medium', color: { argb: 'FF0F172A' } },
          right: { style: 'thin', color: { argb: 'FF94A3B8' } }
        };
      });

      // Llenar tabla
      let currentRow = 10;
      candidates.forEach((cand, idx) => {
        const stats = getCandidateStats(cand.id);
        const row = wsDashboard.getRow(currentRow);
        row.height = 24;

        const isEven = idx % 2 === 1;
        const rowFillColor = isEven ? 'FFF8FAFC' : 'FFFFFFFF';

        const cellsData = [
          { col: 'B', val: `${cand.firstName || ''} ${cand.lastName || ''}`.trim(), align: 'left', fontBold: true, numFmt: null },
          { col: 'C', val: stats.total, align: 'right', fontBold: false, numFmt: '#,##0' },
          { col: 'D', val: stats.approves, align: 'right', fontBold: false, numFmt: '#,##0' },
          { col: 'E', val: stats.total > 0 ? stats.approves / stats.total : 0, align: 'right', fontBold: false, numFmt: '0.0%' },
          { col: 'F', val: stats.disapproves, align: 'right', fontBold: false, numFmt: '#,##0' },
          { col: 'G', val: stats.total > 0 ? stats.disapproves / stats.total : 0, align: 'right', fontBold: false, numFmt: '0.0%' },
          { col: 'H', val: stats.total > 0 ? stats.approvalRate / 100 : 0, align: 'center', fontBold: true, numFmt: '0%' }
        ];

        cellsData.forEach(c => {
          const cell = wsDashboard.getCell(`${c.col}${currentRow}`);
          cell.value = c.val;

          let cellFillColor = rowFillColor;
          let cellFontColor = 'FF334155';

          // Formateo de Status Pill para la tasa de aprobación
          if (c.col === 'H') {
            cellFillColor = stats.approvalRate >= 75 ? 'FFE6F4EA' : 'FFFCE8E6'; // Soft Green o Soft Red
            cellFontColor = stats.approvalRate >= 75 ? 'FF137333' : 'FFC5221F'; // Dark Green o Dark Red
          }

          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: cellFillColor }
          };

          cell.font = { name: 'Segoe UI', size: 10, bold: c.fontBold, color: { argb: cellFontColor } };
          cell.alignment = { vertical: 'middle', horizontal: c.align };
          
          if (c.numFmt) {
            cell.numFmt = c.numFmt;
          }

          cell.border = {
            top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
          };
        });

        currentRow++;
      });

      // Agregar Gráfica
      const chartBase64 = generateChartImage(candidates, getCandidateStats);
      const cleanBase64 = chartBase64.replace(/^data:image\/png;base64,/, '');

      const chartImageId = workbook.addImage({
        base64: cleanBase64,
        extension: 'png',
      });

      wsDashboard.addImage(chartImageId, {
        tl: { col: 1, row: currentRow + 2 },
        ext: { width: 620, height: 330 }
      });

      // ----------------------------------------------------
      // HOJA 2: Detalle de Votos
      // ----------------------------------------------------
      const wsVotes = workbook.addWorksheet('Detalle de Votos');
      wsVotes.views = [{ showGridLines: true }];

      wsVotes.columns = [
        { key: 'spacing', width: 4 }, // Columna A
        { key: 'candidate', width: 28 }, // Columna B
        { key: 'verdict', width: 16 },   // Columna C
        { key: 'comment', width: 55 },   // Columna D
        { key: 'date', width: 22 }       // Columna E
      ];

      // Banner del Título Principal Hoja 2
      wsVotes.mergeCells('B2:E2');
      for (let col = 2; col <= 5; col++) {
        const cell = wsVotes.getCell(2, col);
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF1E3A8A' } // Blue 900
        };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FF1E3A8A' } },
          left: { style: 'thin', color: { argb: 'FF1E3A8A' } },
          bottom: { style: 'thin', color: { argb: 'FF1E3A8A' } },
          right: { style: 'thin', color: { argb: 'FF1E3A8A' } }
        };
      }
      const titleCell2 = wsVotes.getCell('B2');
      titleCell2.value = 'DETALLE DE VOTOS REGISTRADOS';
      titleCell2.font = { name: 'Segoe UI', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
      titleCell2.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
      wsVotes.getRow(2).height = 36;

      // Subtítulo
      wsVotes.mergeCells('B3:E3');
      const subtitleCell2 = wsVotes.getCell('B3');
      subtitleCell2.value = 'Organizado por Candidato: Primero objeciones ("No Aprueba") y luego aprobaciones ("Aprueba")';
      subtitleCell2.font = { name: 'Segoe UI', size: 9, italic: true, color: { argb: 'FF64748B' } };
      wsVotes.getRow(3).height = 20;

      // Cabeceras Tabla
      const headerRow2 = wsVotes.getRow(5);
      headerRow2.height = 28;
      const headers2 = [
        { col: 'B', val: 'Candidato' },
        { col: 'C', val: 'Veredicto' },
        { col: 'D', val: 'Motivo de Objeción (Si aplica)' },
        { col: 'E', val: 'Fecha del Voto' }
      ];

      headers2.forEach(h => {
        const cell = wsVotes.getCell(`${h.col}5`);
        cell.value = h.val;
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF2B6CB0' } // Corporate Blue
        };
        cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.alignment = { vertical: 'middle', horizontal: h.col === 'D' ? 'left' : 'center' };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
          left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
          bottom: { style: 'medium', color: { argb: 'FF1A365D' } },
          right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
        };
      });

      // Procesar votos
      const parsedVotes = votes.map(v => {
        const candId = v.candidateId || v.ID_Candidato || v["ID Candidato"];
        const cand = candidates.find(c => {
          const cId = String(c.id).trim();
          const vId = String(candId || '').trim();
          return cId === vId || (parseInt(cId, 10) === parseInt(vId, 10));
        });
        const candNameVal = v.candidateName ||
          (cand ? `${cand.firstName || ''} ${cand.lastName || ''}`.trim() : '') ||
          v.Nombre_Candidato || v["Nombre Candidato"] || "Candidato Desconocido";
        const statusVal = v.status || v.Estado || v.estado;
        const isApp = statusVal?.toLowerCase() === 'approve' || statusVal?.toLowerCase() === 'aprueba' || statusVal?.toLowerCase() === 'aprobar' || statusVal === 'Aprueba';
        const reason = v.reason || v.Motivo || v.motivo || '';
        const date = v.timestamp || v.Fecha || v.fecha || '';

        return {
          candidateName: candNameVal,
          status: isApp ? 'Aprueba' : 'No Aprueba',
          reason,
          date
        };
      });

      // ORDENAMIENTO REQUERIDO: Organizado por Candidatos, y dentro de cada candidato primero los "No Aprueba"
      const sortedVotes = [...parsedVotes].sort((a, b) => {
        // 1. Agrupar por Candidato
        const nameA = a.candidateName.toLowerCase().trim();
        const nameB = b.candidateName.toLowerCase().trim();
        if (nameA < nameB) return -1;
        if (nameA > nameB) return 1;

        // 2. Dentro del mismo candidato, poner "No Aprueba" antes que "Aprueba"
        if (a.status === 'No Aprueba' && b.status === 'Aprueba') return -1;
        if (a.status === 'Aprueba' && b.status === 'No Aprueba') return 1;
        return 0;
      });

      let rowIdx = 6;
      sortedVotes.forEach((vote, idx) => {
        const row = wsVotes.getRow(rowIdx);
        row.height = 22;

        const isEven = idx % 2 === 1;
        const rowFillColor = isEven ? 'FFF8FAFC' : 'FFFFFFFF';

        const cCand = wsVotes.getCell(`B${rowIdx}`);
        cCand.value = vote.candidateName;
        cCand.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF334155' } };

        const cVerdict = wsVotes.getCell(`C${rowIdx}`);
        cVerdict.value = vote.status;
        cVerdict.alignment = { vertical: 'middle', horizontal: 'center' };
        if (vote.status === 'No Aprueba') {
          cVerdict.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFECEC' } };
          cVerdict.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF9B2C2C' } };
        } else {
          cVerdict.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6FFFA' } };
          cVerdict.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF234E52' } };
        }

        const cComment = wsVotes.getCell(`D${rowIdx}`);
        cComment.value = vote.reason || '—';
        cComment.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        cComment.font = { 
          name: 'Segoe UI', 
          size: 10, 
          italic: vote.status === 'No Aprueba' && !!vote.reason,
          color: vote.status === 'No Aprueba' && !!vote.reason ? { argb: 'FF475569' } : { argb: 'FF94A3B8' } 
        };

        const cDate = wsVotes.getCell(`E${rowIdx}`);
        cDate.value = vote.date ? new Date(vote.date).toLocaleString('es-ES') : '—';
        cDate.alignment = { vertical: 'middle', horizontal: 'center' };
        cDate.font = { name: 'Segoe UI', size: 9, color: { argb: 'FF64748B' } };

        [cCand, cComment, cDate].forEach(c => {
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowFillColor } };
        });

        // Separación visual gruesa al cambiar de candidato
        const isLastOfCandidate = idx === sortedVotes.length - 1 || 
          sortedVotes[idx + 1].candidateName !== vote.candidateName;

        const borderBottomStyle = isLastOfCandidate ? 'medium' : 'thin';
        const borderBottomColor = isLastOfCandidate ? 'FF475569' : 'FFE2E8F0';

        [cCand, cVerdict, cComment, cDate].forEach(c => {
          c.border = {
            top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            bottom: { style: borderBottomStyle, color: { argb: borderBottomColor } },
            right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
          };
        });

        rowIdx++;
      });

      // Escribir archivo y descargar en navegador
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = `Reporte_Votaciones_ICC_${new Date().toISOString().slice(0, 10)}.xlsx`;
      anchor.click();
      window.URL.revokeObjectURL(url);

      showToast('Reporte Excel exportado correctamente', 'success');
    } catch (error) {
      console.error('Error al exportar a Excel:', error);
      showToast(`Error al exportar Excel: ${error.message}`, 'error');
    } finally {
      setIsExporting(false);
    }
  };

  const handleResetLocalData = () => {
    if (window.confirm("¿Deseas restaurar la base de datos local a los valores iniciales de prueba? Esto borrará tus cambios locales.")) {
      localStorage.removeItem('icc_local_candidates');
      localStorage.removeItem('icc_local_voters');
      localStorage.removeItem('icc_local_votes');
      localStorage.removeItem('icc_candidates_cleared_by_user');
      localStorage.removeItem('icc_voter_attendance');
      refreshData();
      showToast("Datos locales restablecidos", "info");
    }
  };

  // Exportar datos locales a Google Sheets
  const handleExportLocalDataToSheets = async () => {
    if (!config.sheetUrlVoters || !config.sheetUrlCandidates) {
      showToast("Primero configura y guarda ambas URLs de Google Sheets (Votantes y Registro de Votos)", "error");
      return;
    }
    
    if (!window.confirm("¿Seguro que deseas exportar tus datos locales (votos y miembros) a tus Google Sheets?")) {
      return;
    }

    setIsExporting(true);
    try {
      // 1. Subir miembros
      showToast("Exportando lista de miembros a Google Sheets...", "info");
      await updateVotersInSheets(config.sheetUrlVoters, voters);

      // 2. Subir votos
      showToast("Exportando votos registrados...", "info");
      for (const vote of votes) {
        const cId = vote.candidateId;
        const cFirstName = vote.candidateFirstName || vote.candidateName?.split(' ')[0] || '';
        const cLastName = vote.candidateLastName || vote.candidateName?.split(' ').slice(1).join(' ') || '';
        const vFirstName = vote.voterFirstName || vote.voterName?.split(' ')[0] || '';
        const vLastName = vote.voterLastName || vote.voterName?.split(' ').slice(1).join(' ') || '';
        const vStatus = vote.status;
        const vReason = vote.reason || '';

        if (cId) {
          await addVoteToSheets(config.sheetUrlCandidates, {
            candidateId: cId,
            candidateFirstName: cFirstName,
            candidateLastName: cLastName,
            voterFirstName: vFirstName,
            voterLastName: vLastName,
            status: vStatus,
            reason: vReason
          });
        }
      }

      showToast("¡Votos y miembros exportados con éxito a Google Sheets!", "success");
      refreshData();
    } catch (error) {
      console.error(error);
      showToast(`Error al exportar datos: ${error.message}`, "error");
    } finally {
      setIsExporting(false);
    }
  };

  // Sembrar datos de prueba en Google Sheets
  const handleSeedSheetsWithMockData = async () => {
    if (!config.sheetUrlVoters || !config.sheetUrlCandidates) {
      showToast("Primero configura y guarda ambas URLs de Google Sheets (Votantes y Registro de Votos)", "error");
      return;
    }
    
    if (!window.confirm("¿Deseas poblar tus Google Sheets con los 5 miembros, 5 candidatos y 5 votos de prueba?")) {
      return;
    }

    setIsExporting(true);
    try {
      localStorage.removeItem('icc_candidates_cleared_by_user');
      // 1. Subir miembros
      showToast("Sembrando lista de miembros en Google Sheets...", "info");
      await updateVotersInSheets(config.sheetUrlVoters, MOCK_VOTERS);

      // 2. Subir candidatos
      showToast("Sembrando lista de candidatos en Google Sheets...", "info");
      for (const cand of MOCK_CANDIDATES) {
        await addCandidateToSheets(config.sheetUrlVoters, {
          id: cand.id,
          firstName: cand.firstName,
          lastName: cand.lastName,
          testimony: cand.testimony
        });
        // Registrar foto localmente asociada a ese ID
        if (!localStorage.getItem(`icc_photo_${cand.id}`)) {
          safeSetLocalStorage(`icc_photo_${cand.id}`, cand.photo);
        }
      }

      // 3. Subir votos
      showToast("Sembrando votos registrados...", "info");
      for (const vote of MOCK_VOTES) {
        // Encontrar nombres separados o extraerlos
        const cand = candidates.find(c => c.id === vote.candidateId) || {};
        const voterParts = vote.voterName.split(' ');
        const voterFirst = voterParts[0] || '';
        const voterLast = voterParts.slice(1).join(' ') || '';

        await addVoteToSheets(config.sheetUrlCandidates, {
          candidateId: vote.candidateId,
          candidateFirstName: cand.firstName || vote.candidateName.split(' ')[0] || '',
          candidateLastName: cand.lastName || vote.candidateName.split(' ').slice(1).join(' ') || '',
          voterFirstName: voterFirst,
          voterLastName: voterLast,
          status: vote.status === 'approve' || vote.status === 'Aprueba' ? 'Aprueba' : 'No Aprueba',
          reason: vote.reason || ''
        });
      }

      showToast("¡Datos de prueba subidos con éxito a Google Sheets!", "success");
      refreshData();
    } catch (error) {
      console.error(error);
      showToast(`Error al sembrar datos: ${error.message}`, "error");
    } finally {
      setIsExporting(false);
    }
  };

  // Copiar URL de votación al portapapeles
  const handleCopyVotingLink = () => {
    navigator.clipboard.writeText(votingUrl);
    showToast("Enlace de votación copiado", "success");
  };

  // Calcular la prioridad/coincidencia de la búsqueda (empieza con -> contiene)
  const getSearchScore = (fullName, query) => {
    if (!query) return 0;
    if (fullName.startsWith(query)) return 3;
    const words = fullName.split(/\s+/);
    if (words.some(w => w.startsWith(query))) return 2;
    if (fullName.includes(query)) return 1;
    return 0;
  };

  // Filtrar Votantes en lista considerando búsqueda, asistencia y progreso de votación (VOTO SECRETO)
  const filteredVoters = voters.filter(v => {
    const fullName = `${v.name || ''} ${v.lastName || ''}`.trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
    const searchNormalized = voterSearch.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
    const matchesSearch = fullName.includes(searchNormalized);
    if (!matchesSearch) return false;

    // Filtro por Asistencia (Presencia)
    if (voterPresenceFilter === 'present' && v.isPresent === false) return false;
    if (voterPresenceFilter === 'absent' && v.isPresent !== false) return false;

    // Filtro por Progreso de Votación
    const progress = getVoterProgress(v);
    if (voterProgressFilter === 'completed' && !progress.isComplete) return false;
    if (voterProgressFilter === 'pending' && progress.hasVoted) return false;
    if (voterProgressFilter === 'partial' && (!progress.hasVoted || progress.isComplete)) return false;

    return true;
  });

  // Ordenar Votantes filtrados según criterios seleccionados (Excel-like)
  const sortedVoters = [...filteredVoters].sort((a, b) => {
    // Si hay una búsqueda activa, priorizar palabras que empiecen con el término buscado
    if (voterSearch.trim() !== '') {
      const normA = `${a.name || ''} ${a.lastName || ''}`.trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
      const normB = `${b.name || ''} ${b.lastName || ''}`.trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
      const queryNorm = voterSearch.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();

      const scoreA = getSearchScore(normA, queryNorm);
      const scoreB = getSearchScore(normB, queryNorm);

      if (scoreA !== scoreB) {
        return scoreB - scoreA; // Mayor relevancia primero
      }
    }

    let valA = '';
    let valB = '';

    if (voterSortKey === 'name') {
      valA = `${a.name || ''} ${a.lastName || ''}`.trim().toLowerCase();
      valB = `${b.name || ''} ${b.lastName || ''}`.trim().toLowerCase();
    } else if (voterSortKey === 'presence') {
      valA = a.isPresent !== false ? 'present' : 'absent';
      valB = b.isPresent !== false ? 'present' : 'absent';
    } else if (voterSortKey === 'progress') {
      const progA = getVoterProgress(a);
      const progB = getVoterProgress(b);
      // Orden: completes (2), partials (1), pending (0)
      valA = progA.isComplete ? 2 : (progA.hasVoted ? 1 : 0);
      valB = progB.isComplete ? 2 : (progB.hasVoted ? 1 : 0);
    }

    if (valA < valB) return voterSortDirection === 'asc' ? -1 : 1;
    if (valA > valB) return voterSortDirection === 'asc' ? 1 : -1;
    return 0;
  });

  // Desglose de votos para la tabla completa (VOTO SECRETO: sin nombre del votante)
  const allVotesParsed = votes.map(v => {
    const candId = v.candidateId || v.ID_Candidato || v["ID Candidato"];
    const cand = candidates.find(c => {
      const cId = String(c.id).trim();
      const vId = String(candId || '').trim();
      return cId === vId || (parseInt(cId, 10) === parseInt(vId, 10));
    });
    const candNameVal = v.candidateName ||
      (cand ? `${cand.firstName || ''} ${cand.lastName || ''}`.trim() : '') ||
      v.Nombre_Candidato || v["Nombre Candidato"] || "Candidato Desconocido";
    const statusVal = v.status || v.Estado || v.estado;
    const isApp = statusVal?.toLowerCase() === 'approve' || statusVal?.toLowerCase() === 'aprueba' || statusVal?.toLowerCase() === 'aprobar' || statusVal === 'Aprueba';
    const reason = v.reason || v.Motivo || v.motivo || '';
    const date = v.timestamp || v.Fecha || v.fecha || '';

    return {
      candId,
      candidateName: candNameVal,
      // voterId opaco — no se muestra en UI (voto secreto)
      status: isApp ? 'Aprueba' : 'No Aprueba',
      reason,
      date
    };
  });

  const displayVotes = resultsFilter === 'all' 
    ? allVotesParsed 
    : allVotesParsed.filter(v => v.candId === resultsFilter);

  // Paginación y control de cantidad de votos en Dashboard
  const totalVotesCount = displayVotes.length;
  const maxPage = votesPerPage === 'all' ? 1 : Math.ceil(totalVotesCount / votesPerPage);
  const activePage = Math.min(votesCurrentPage, maxPage) || 1;
  const startIndex = votesPerPage === 'all' ? 0 : (activePage - 1) * votesPerPage;
  const endIndex = votesPerPage === 'all' ? totalVotesCount : Math.min(startIndex + Number(votesPerPage), totalVotesCount);
  const paginatedVotes = displayVotes.slice(startIndex, endIndex);

  return (
    <div className="card" style={{ flexGrow: 1, padding: '24px', textAlign: 'left' }}>

      {/* Tabs */}
      <div className="tabs-header">
        <Tooltip text="Ver el avance de los votos, gráficos de aprobación e historial de decisiones." position="bottom">
          <button 
            className={`tab-btn ${activeTab === 'results' ? 'active' : ''}`}
            onClick={() => setActiveTab('results')}
          >
            Resultados y Estadísticas
          </button>
        </Tooltip>
        <Tooltip text="Agregar, editar o eliminar los candidatos de la iglesia." position="bottom">
          <button 
            className={`tab-btn ${activeTab === 'candidates' ? 'active' : ''}`}
            onClick={() => setActiveTab('candidates')}
          >
            Registrar Candidatos ({candidates.length})
          </button>
        </Tooltip>
        <Tooltip text="Administrar la lista oficial de miembros autorizados para votar." position="bottom">
          <button 
            className={`tab-btn ${activeTab === 'voters' ? 'active' : ''}`}
            onClick={() => setActiveTab('voters')}
          >
            Lista de Votantes ({voters.length})
          </button>
        </Tooltip>
        <Tooltip text="Mostrar código QR y enlace para acceder a votar desde celulares." position="bottom">
          <button 
            className={`tab-btn ${activeTab === 'share' ? 'active' : ''}`}
            onClick={() => setActiveTab('share')}
          >
            Compartir QR
          </button>
        </Tooltip>
        {/* Pestaña de Ajustes/Configuración removida del menú */}
      </div>

      {/* Contenido de Tabs */}

      {/* TAB RESULTS */}
      {activeTab === 'results' && (
        <div>
          <div className="dashboard-header">
            <h2 style={{ fontFamily: 'var(--font-heading)', fontWeight: 700, margin: 0 }}>Resultados de la Votación</h2>
          </div>

          {/* Tarjetas de Estadísticas Globales */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '20px', marginBottom: '32px' }}>
            <div className="card stat-card">
              <div className="stat-label">Candidatos Activos</div>
              <div className="stat-val">{candidates.length}</div>
            </div>
            <div className="card stat-card">
              <div className="stat-label">Votos Registrados</div>
              <div className="stat-val">{votes.length}</div>
            </div>
            <div className="card stat-card">
              <div className="stat-label">Han Participado</div>
              <div className="stat-val">
                {(() => {
                  const participatedCount = voters.filter(v => {
                    const voterVotes = votes.filter(vote => {
                      const vId = String(vote.voterId || vote.ID_Votante || vote["ID Votante"] || '').trim();
                      const voterId = String(v.id).trim();
                      return vId === voterId;
                    });
                    return v.hasVoted === true || voterVotes.length > 0;
                  }).length;
                  const participationPct = voters.length > 0 ? (participatedCount / voters.length) * 100 : 0;
                  return `${participatedCount} / ${voters.length} (${Math.round(participationPct)}%)`;
                })()}
              </div>
            </div>
            <div className="card stat-card">
              <div className="stat-label">Pendientes de Votar</div>
              <div className="stat-val">
                {(() => {
                  const participatedCount = voters.filter(v => {
                    const voterVotes = votes.filter(vote => {
                      const vId = String(vote.voterId || vote.ID_Votante || vote["ID Votante"] || '').trim();
                      const voterId = String(v.id).trim();
                      return vId === voterId;
                    });
                    return v.hasVoted === true || voterVotes.length > 0;
                  }).length;
                  const pendingCount = voters.length - participatedCount;
                  const pendingPct = voters.length > 0 ? (pendingCount / voters.length) * 100 : 0;
                  return `${pendingCount} / ${voters.length} (${Math.round(pendingPct)}%)`;
                })()}
              </div>
            </div>
          </div>

          {/* Resultados por Candidato */}
          {isLoading && candidates.length === 0 ? (
            <div style={{ textAlign: 'center', padding: '48px 0', color: 'var(--text-secondary)' }}>
              <div className="spinner" style={{ display: 'inline-block', width: '40px', height: '40px', border: '4px solid var(--border)', borderTopColor: 'var(--primary)', borderRadius: '50%', animation: 'spin 1s linear infinite', marginBottom: '12px' }}></div>
              <div>Cargando datos desde Google Sheets...</div>
              <style>{`
                @keyframes spin {
                  0% { transform: rotate(0deg); }
                  100% { transform: rotate(360deg); }
                }
              `}</style>
            </div>
          ) : candidates.length === 0 ? (
            <div className="card" style={{ padding: '32px 24px', textAlign: 'center', border: '1px dashed var(--border)', background: 'var(--bg-input)', borderRadius: 'var(--radius-lg)', margin: '20px 0' }}>
              <div style={{ fontSize: '48px', marginBottom: '16px' }}>📂</div>
              <h4 style={{ fontFamily: 'var(--font-heading)', fontWeight: 700, fontSize: '18px', marginBottom: '8px' }}>
                {isConnected ? 'Hojas de Cálculo Conectadas y Vacías' : 'No Hay Datos Locales'}
              </h4>
              <p style={{ color: 'var(--text-secondary)', fontSize: '14px', maxWidth: '480px', margin: '0 auto 24px' }}>
                {isConnected 
                  ? 'Tus Google Sheets se conectaron correctamente, pero no tienen datos cargados. Inicialízalos con los 5 candidatos y 5 votantes de prueba para comenzar.'
                  : 'No se encontraron candidatos ni votantes locales en tu navegador.'
                }
              </p>
              <Tooltip text={isConnected ? "Subir los candidatos, votantes y votos de prueba iniciales a tu Google Sheets." : "Restaurar las listas de prueba por defecto en el almacenamiento local."} position="top">
                <button 
                  type="button" 
                  className="btn btn-primary"
                  onClick={isConnected ? handleSeedSheetsWithMockData : handleResetLocalData}
                  disabled={isExporting || isLoading}
                >
                  {isExporting ? 'Cargando datos...' : isConnected ? 'Sembrar Datos de Prueba en Google Sheets' : 'Restablecer Base de Datos Local'}
                </button>
              </Tooltip>
            </div>
          ) : (
            <div className="card collapsible-section" style={{ marginBottom: '24px', padding: 0, overflow: 'hidden', border: '1px solid var(--border)' }}>
              <button 
                type="button" 
                onClick={() => setShowCandidateResults(!showCandidateResults)}
                style={{
                  width: '100%',
                  padding: '16px 20px',
                  background: 'var(--bg-card)',
                  border: 'none',
                  display: 'flex',
                  justifyContent: 'space-between',
                  alignItems: 'center',
                  cursor: 'pointer',
                  textAlign: 'left',
                  outline: 'none',
                  fontFamily: 'var(--font-heading)',
                  fontWeight: 600,
                  fontSize: '16px',
                  color: 'var(--text-primary)'
                }}
              >
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <span>📊</span>
                  <span>Votación por Candidato (Resumen)</span>
                </div>
                <span style={{ 
                  transform: showCandidateResults ? 'rotate(180deg)' : 'rotate(0deg)', 
                  transition: 'transform 0.2s ease',
                  fontSize: '12px',
                  color: 'var(--text-secondary)'
                }}>
                  ▼
                </span>
              </button>
              
              {showCandidateResults && (
                <div style={{ padding: '20px', borderTop: '1px solid var(--border)' }}>
                  <div className="grid-2">
                    {candidates.map(cand => {
                      const stats = getCandidateStats(cand.id);
                      const fullName = `${cand.firstName || ''} ${cand.lastName || ''}`.trim();
                      return (
                        <div key={cand.id} className={`card result-card ${stats.approvalRate >= 75 ? 'high-approval' : ''}`} style={{ margin: 0 }}>
                          <div style={{ display: 'flex', gap: '16px', alignItems: 'center', marginBottom: '12px' }}>
                            <CandidatePhoto
                              photo={cand.photo}
                              firstName={cand.firstName}
                              lastName={cand.lastName}
                              className="candidate-avatar"
                              style={{ width: '60px', height: '60px' }}
                            />
                            <div>
                              <h4 className="candidate-name" style={{ margin: 0 }}>{fullName}</h4>
                              <div style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
                                Aprobación: <strong>{stats.approvalRate}%</strong> ({stats.total} votos totales)
                              </div>
                            </div>
                          </div>

                          <div className="progress-container">
                            <div className="progress-bar-bg">
                              <div className="progress-bar-fill-approve" style={{ width: `${stats.total > 0 ? (stats.approves / stats.total) * 100 : 0}%` }}></div>
                              <div className="progress-bar-fill-disapprove" style={{ width: `${stats.total > 0 ? (stats.disapproves / stats.total) * 100 : 0}%` }}></div>
                            </div>
                            <div className="progress-stats">
                              <span style={{ color: 'var(--success)' }}>Aprueba: {stats.approves}</span>
                              <span style={{ color: 'var(--danger)' }}>No Aprueba: {stats.disapproves}</span>
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>
          )}

          {/* Desglose Detallado de Votos (Colapsable) */}
          {candidates.length > 0 && (
            <div className="card collapsible-section" style={{ marginBottom: '24px', padding: 0, overflow: 'hidden', border: '1px solid var(--border)' }}>
              <button 
                type="button" 
                onClick={() => setShowVotesList(!showVotesList)}
                style={{
                  width: '100%',
                  padding: '16px 20px',
                  background: 'var(--bg-card)',
                  border: 'none',
                  display: 'flex',
                  justifyContent: 'space-between',
                  alignItems: 'center',
                  cursor: 'pointer',
                  textAlign: 'left',
                  outline: 'none',
                  fontFamily: 'var(--font-heading)',
                  fontWeight: 600,
                  fontSize: '16px',
                  color: 'var(--text-primary)'
                }}
              >
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <span>📋</span>
                  <span>Desglose Detallado de Votos ({displayVotes.length})</span>
                </div>
                <span style={{ 
                  transform: showVotesList ? 'rotate(180deg)' : 'rotate(0deg)', 
                  transition: 'transform 0.2s ease',
                  fontSize: '12px',
                  color: 'var(--text-secondary)'
                }}>
                  ▼
                </span>
              </button>

              {showVotesList && (
                <div style={{ padding: '20px', borderTop: '1px solid var(--border)' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', flexWrap: 'wrap', gap: '12px' }}>
                    <h3 style={{ fontFamily: 'var(--font-heading)', fontWeight: 600, margin: 0, fontSize: '15px' }}>Detalles de Votos Recibidos</h3>
                    <div style={{ display: 'flex', gap: '16px', alignItems: 'center', flexWrap: 'wrap' }}>
                      <div>
                        <label style={{ fontSize: '13px', marginRight: '8px', fontWeight: 600, color: 'var(--text-secondary)' }}>Mostrar:</label>
                        <select 
                          className="form-control" 
                          style={{ display: 'inline-block', width: 'auto', padding: '6px 12px', fontSize: '13px' }}
                          value={votesPerPage}
                          onChange={(e) => {
                            const val = e.target.value === 'all' ? 'all' : Number(e.target.value);
                            setVotesPerPage(val);
                            setVotesCurrentPage(1);
                          }}
                        >
                          <option value={10}>10 votos</option>
                          <option value={20}>20 votos</option>
                          <option value={30}>30 votos</option>
                          <option value={50}>50 votos</option>
                          <option value="all">Todos</option>
                        </select>
                      </div>
                      <div>
                        <label style={{ fontSize: '13px', marginRight: '8px', fontWeight: 600, color: 'var(--text-secondary)' }}>Filtrar candidato:</label>
                        <select 
                          className="form-control" 
                          style={{ display: 'inline-block', width: 'auto', padding: '6px 12px', fontSize: '13px' }}
                          value={resultsFilter}
                          onChange={(e) => {
                            setResultsFilter(e.target.value);
                            setVotesCurrentPage(1);
                          }}
                        >
                          <option value="all">Todos los candidatos</option>
                          {candidates.map(c => (
                            <option key={c.id} value={c.id}>{`${c.firstName || ''} ${c.lastName || ''}`.trim()}</option>
                          ))}
                        </select>
                      </div>
                    </div>
                  </div>

                  <div className="votos-table-container">
                    {displayVotes.length === 0 ? (
                      <div style={{ padding: '24px', textAlign: 'center', color: 'var(--text-muted)' }}>
                        No se han registrado votos todavía para este criterio.
                      </div>
                    ) : (
                      <>
                        <table className="votos-table">
                          <thead>
                            <tr>
                              <th>Candidato</th>
                              <th>Veredicto</th>
                              <th>Motivo de Objeción (Si aplica)</th>
                              <th>Fecha</th>
                            </tr>
                          </thead>
                          <tbody>
                            {paginatedVotes.map((vote, index) => (
                              <tr key={index}>
                                <td data-label="Candidato" style={{ fontWeight: 600 }}>{vote.candidateName}</td>
                                <td data-label="Veredicto">
                                  <span className={`badge ${vote.status === 'Aprueba' ? 'badge-success' : 'badge-danger'}`}>
                                    {vote.status}
                                  </span>
                                </td>
                                <td data-label="Motivo" style={{ color: vote.status === 'No Aprueba' ? 'var(--danger)' : 'var(--text-secondary)', fontStyle: vote.status === 'No Aprueba' ? 'italic' : 'normal' }}>
                                  {vote.reason ? vote.reason : '—'}
                                </td>
                                <td data-label="Fecha" style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
                                  {vote.date ? new Date(vote.date).toLocaleString('es-ES') : '—'}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>

                        {/* Controles de Paginación */}
                        {votesPerPage !== 'all' && totalVotesCount > votesPerPage && (
                          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '16px', flexWrap: 'wrap', gap: '12px' }}>
                            <div style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>
                              Mostrando <strong>{startIndex + 1}</strong> - <strong>{endIndex}</strong> de <strong>{totalVotesCount}</strong> votos
                            </div>
                            <div style={{ display: 'flex', gap: '6px' }}>
                              <button 
                                type="button"
                                className="btn btn-secondary" 
                                style={{ padding: '6px 12px', fontSize: '13px', height: '34px', display: 'flex', alignItems: 'center' }}
                                disabled={activePage === 1}
                                onClick={() => setVotesCurrentPage(prev => Math.max(prev - 1, 1))}
                              >
                                Anterior
                              </button>
                              {Array.from({ length: maxPage }, (_, i) => i + 1).map(pageNum => {
                                if (maxPage > 6 && Math.abs(pageNum - activePage) > 1 && pageNum !== 1 && pageNum !== maxPage) {
                                  if (pageNum === 2 || pageNum === maxPage - 1) {
                                    return <span key={pageNum} style={{ alignSelf: 'center', padding: '0 4px', color: 'var(--text-muted)' }}>...</span>;
                                  }
                                  return null;
                                }
                                return (
                                  <button 
                                    type="button"
                                    key={pageNum}
                                    className={`btn ${pageNum === activePage ? 'btn-primary' : 'btn-secondary'}`}
                                    style={{ padding: '6px 12px', fontSize: '13px', minWidth: '35px', height: '34px', display: 'flex', alignItems: 'center', justifyContent: 'center' }}
                                    onClick={() => setVotesCurrentPage(pageNum)}
                                  >
                                    {pageNum}
                                  </button>
                                );
                              })}
                              <button 
                                type="button"
                                className="btn btn-secondary" 
                                style={{ padding: '6px 12px', fontSize: '13px', height: '34px', display: 'flex', alignItems: 'center' }}
                                disabled={activePage === maxPage}
                                onClick={() => setVotesCurrentPage(prev => Math.min(prev + 1, maxPage))}
                              >
                                Siguiente
                              </button>
                            </div>
                          </div>
                        )}
                      </>
                    )}
                  </div>
                </div>
              )}
            </div>
          )}


          {/* Database actions panel at the bottom */}
          <div style={{ marginTop: '40px', paddingTop: '24px', borderTop: '1px solid var(--border)', display: 'flex', gap: '12px', justifyContent: 'flex-end', flexWrap: 'wrap' }}>
            {votes.length > 0 && (
              <>
                <Tooltip text="Exportar reporte detallado y gráfico de resultados a un archivo Excel." position="top">
                  <button 
                    type="button"
                    className="btn btn-success"
                    onClick={handleExportExcel}
                    disabled={isExporting}
                    style={{ padding: '8px 16px', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '6px' }}
                  >
                    {isExporting ? (
                      <>
                        <div className="spinner-small" style={{ display: 'inline-block', width: '12px', height: '12px', border: '2px solid rgba(255,255,255,0.3)', borderTopColor: 'white', borderRadius: '50%', animation: 'spin 1s linear infinite' }}></div>
                        <span>Exportando...</span>
                      </>
                    ) : (
                      <>
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                          <polyline points="7 10 12 15 17 10" />
                          <line x1="12" y1="15" x2="12" y2="3" />
                        </svg>
                        <span>Exportar a Excel</span>
                      </>
                    )}
                  </button>
                </Tooltip>
                <Tooltip text="Eliminar permanentemente todos los votos de la base de datos." position="top">
                  <button 
                    type="button"
                    className="btn btn-danger"
                    onClick={handleClearAllVotes}
                    disabled={isUploading}
                    style={{ padding: '8px 16px', fontSize: '13px' }}
                  >
                    {isUploading ? 'Reiniciando...' : 'Reiniciar Votación (Limpiar Votos)'}
                  </button>
                </Tooltip>
              </>
            )}
            <Tooltip text="Restaurar las listas de prueba por defecto en el almacenamiento local del navegador.">
              <button type="button" className="btn btn-secondary" style={{ fontSize: '13px' }} onClick={handleResetLocalData}>
                Restablecer Datos Locales
              </button>
            </Tooltip>
            {isConnected && (
              <>
                <Tooltip text="Subir los candidatos, votantes y votos de prueba iniciales a tu Google Sheets.">
                  <button 
                    type="button" 
                    className="btn btn-secondary" 
                    onClick={handleSeedSheetsWithMockData}
                    disabled={isExporting || isLoading}
                    style={{ borderColor: 'var(--accent)', color: 'var(--accent)', fontSize: '13px' }}
                  >
                    {isExporting ? 'Sembrando...' : 'Sembrar Datos de Prueba en Google Sheets'}
                  </button>
                </Tooltip>
                <Tooltip text="Subir toda la información y votos locales actuales a tus hojas de cálculo en la nube.">
                  <button 
                    type="button" 
                    className="btn btn-secondary" 
                    onClick={handleExportLocalDataToSheets}
                    disabled={isExporting || isLoading}
                    style={{ borderColor: 'var(--primary)', color: 'var(--primary)', fontSize: '13px' }}
                  >
                    {isExporting ? 'Exportando...' : 'Exportar Datos Locales a Google Sheets'}
                  </button>
                </Tooltip>
              </>
            )}
          </div>
        </div>
      )}

      {/* TAB REGISTER CANDIDATES */}
      {activeTab === 'candidates' && (
        <div>
          <h2 style={{ fontFamily: 'var(--font-heading)', fontWeight: 700, marginBottom: '20px' }}>Gestión de Candidatos</h2>
          
          <div style={{ padding: '12px 16px', backgroundColor: 'rgba(59, 130, 246, 0.05)', borderLeft: '4px solid var(--primary)', borderRadius: '0 var(--radius-md) var(--radius-md) 0', fontSize: '13px', color: 'var(--text-secondary)', marginBottom: '20px', display: 'flex', gap: '10px', alignItems: 'center', animation: 'fadeIn 0.2s ease-out' }}>
            <span style={{ fontSize: '16px' }}>ℹ️</span>
            <span><strong>Nota de Sincronización:</strong> Si agregas, editas o eliminas candidatos, los votantes que tengan su pantalla abierta deberán <strong>actualizar/recargar</strong> su navegador para visualizar los cambios.</span>
          </div>
          
          <form onSubmit={handleAddCandidate} className="card" style={{ marginBottom: '24px', border: '1px solid var(--border)' }}>
            <h3 style={{ fontFamily: 'var(--font-heading)', fontWeight: 600, fontSize: '18px', marginBottom: '16px' }}>
              {editingCandidateId ? `Editar Candidato (ID: ${editingCandidateId})` : 'Agregar Nuevo Candidato'}
            </h3>
            
            <div className="grid-3">
              <div className="form-group">
                <label className="form-label">Nombre</label>
                <input 
                  type="text" 
                  className="form-control" 
                  placeholder="Ej. Juan Marcos"
                  value={candFirstName}
                  onChange={(e) => setCandFirstName(e.target.value)}
                  required
                />
              </div>

              <div className="form-group">
                <label className="form-label">Apellido</label>
                <input 
                  type="text" 
                  className="form-control" 
                  placeholder="Ej. Delgado"
                  value={candLastName}
                  onChange={(e) => setCandLastName(e.target.value)}
                  required
                />
              </div>

              <div className="form-group">
                <label className="form-label">
                  {editingCandidateId ? "Nueva Fotografía (Opcional)" : "Fotografía del Candidato"}
                </label>
                <input 
                  type="file" 
                  accept="image/*" 
                  className="form-control" 
                  onChange={(e) => {
                    if (e.target.files && e.target.files[0]) {
                      setCropImageFile(e.target.files[0]);
                    }
                  }}
                  ref={fileInputRef}
                />
                <span style={{ fontSize: '11px', color: 'var(--text-muted)', marginTop: '4px', display: 'block' }}>
                  {editingCandidateId 
                    ? "Deja este campo vacío para conservar la foto actual." 
                    : config.supabaseUrl 
                      ? "✓ La foto se subirá automáticamente a Supabase Storage." 
                      : "⚡ Modo Local: La foto se procesará como imagen incrustada (Base64)."
                  }
                </span>

                {/* Previsualización de la foto */}
                {(candPhotoFile || (editingCandidateId && candidates.find(c => c.id === editingCandidateId)?.photo)) && (
                  <div style={{ marginTop: '12px', display: 'flex', alignItems: 'center', gap: '12px', backgroundColor: 'var(--bg-input)', padding: '10px', borderRadius: 'var(--radius-md)', border: '1px solid var(--border)' }}>
                    {candPhotoFile ? (
                      <img 
                        src={URL.createObjectURL(candPhotoFile)} 
                        alt="Vista previa" 
                        style={{ width: '60px', height: '60px', borderRadius: '50%', objectFit: 'cover', border: '2px solid var(--primary-light)' }} 
                      />
                    ) : (
                      <CandidatePhoto
                        photo={candidates.find(c => c.id === editingCandidateId)?.photo}
                        firstName={candFirstName}
                        lastName={candLastName}
                        className="candidate-avatar"
                        style={{ width: '60px', height: '60px', border: '2px solid var(--primary-light)' }}
                      />
                    )}
                    <div>
                      <span style={{ fontSize: '13px', fontWeight: '600', display: 'block', color: 'var(--text-primary)' }}>
                        {candPhotoFile ? "Nueva foto seleccionada" : "Foto actual del candidato"}
                      </span>
                      {candPhotoFile && (
                        <button 
                          type="button" 
                          className="btn-link" 
                          style={{ color: 'var(--danger)', fontSize: '12px', background: 'none', border: 'none', padding: 0, cursor: 'pointer', textDecoration: 'underline' }}
                          onClick={() => {
                            setCandPhotoFile(null);
                            if (fileInputRef.current) fileInputRef.current.value = '';
                          }}
                        >
                          Quitar foto
                        </button>
                      )}
                    </div>
                  </div>
                )}
              </div>
            </div>

            <div className="form-group">
              <label className="form-label">Testimonio (¿Por qué quiere unirse a la iglesia?)</label>
              <textarea 
                className="form-control" 
                rows="4" 
                placeholder="Escribe el testimonio personal del candidato..."
                value={candTestimony}
                onChange={(e) => setCandTestimony(e.target.value)}
                required
              ></textarea>
            </div>

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '12px' }}>
              {editingCandidateId && (
                <Tooltip text="Descartar los cambios actuales y volver al formulario de registro.">
                  <button type="button" className="btn btn-secondary" onClick={handleCancelEdit}>
                    Cancelar Edición
                  </button>
                </Tooltip>
              )}
              <Tooltip text={editingCandidateId ? "Guardar los cambios editados del candidato." : "Registrar al candidato en la base de datos de votación."}>
                <button type="submit" className="btn btn-primary" disabled={isUploading}>
                  {isUploading 
                    ? "Cargando fotografía..." 
                    : editingCandidateId 
                      ? "Guardar Cambios" 
                      : "Guardar Candidato"
                  }
                </button>
              </Tooltip>
            </div>
          </form>

          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
            <h3 style={{ fontFamily: 'var(--font-heading)', fontWeight: 600, margin: 0 }}>Lista de Candidatos Activos ({candidates.length})</h3>
            {candidates.length > 0 && (
              <Tooltip text="Eliminar permanentemente todos los candidatos de la lista y sus fotos." position="left">
                <button 
                  type="button"
                  className="btn btn-danger" 
                  onClick={handleClearAllCandidates}
                  disabled={isUploading}
                  style={{ padding: '6px 12px', fontSize: '12px' }}
                >
                  {isUploading ? 'Vaciando...' : 'Vaciar Candidatos'}
                </button>
              </Tooltip>
            )}
          </div>
          {candidates.length === 0 ? (
            <p style={{ color: 'var(--text-muted)' }}>No hay candidatos registrados.</p>
          ) : (
            <div>
              {candidates.map(cand => {
                const fullName = `${cand.firstName || ''} ${cand.lastName || ''}`.trim();
                const isExpanded = expandedCandidateId === cand.id;
                return (
                  <div 
                    key={cand.id} 
                    className={`candidate-row-card candidate-admin-item ${isExpanded ? 'expanded' : ''}`}
                    role="button"
                    tabIndex={0}
                    onClick={() => {
                      console.log("Card clicked for candidate:", cand.id, "fullName:", fullName);
                      setExpandedCandidateId(prev => prev === cand.id ? null : cand.id);
                    }}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter' || e.key === ' ') {
                        e.preventDefault();
                        setExpandedCandidateId(prev => prev === cand.id ? null : cand.id);
                      }
                    }}
                    style={{ cursor: 'pointer' }}
                  >
                    <div className="candidate-row-content">
                      <div className="candidate-info">
                        <CandidatePhoto
                          photo={cand.photo}
                          firstName={cand.firstName}
                          lastName={cand.lastName}
                          className="candidate-avatar"
                        />
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div className="candidate-name">{fullName}</div>
                          <div className="candidate-testimony-snippet-wrapper">
                            <div className="candidate-testimony-snippet">{cand.testimony}</div>
                          </div>
                          <div className="candidate-testimony-full-wrapper">
                            <div className="candidate-testimony-full">
                              "{cand.testimony}"
                            </div>
                          </div>
                        </div>
                      </div>
                      <div className="candidate-actions-desktop">
                        <Tooltip text="Modificar la información, foto o testimonio de este candidato.">
                          <button 
                            type="button"
                            className="btn btn-secondary" 
                            onClick={(e) => {
                              e.stopPropagation();
                              handleStartEditCandidate(cand);
                            }}
                            style={{ padding: '6px 12px', fontSize: '12px' }}
                          >
                            Editar
                          </button>
                        </Tooltip>
                        <Tooltip text="Dar de baja de forma permanente a este candidato de la votación.">
                          <button 
                            type="button"
                            className="btn btn-danger" 
                            onClick={(e) => {
                              e.stopPropagation();
                              handleDeleteCandidate(cand.id);
                            }}
                            style={{ padding: '6px 12px', fontSize: '12px' }}
                          >
                            Eliminar
                          </button>
                        </Tooltip>
                      </div>
                    </div>
                    <div className="candidate-actions-mobile-wrapper">
                      <div className="candidate-actions-mobile" onClick={(e) => e.stopPropagation()}>
                        <button 
                          type="button"
                          className="btn btn-secondary" 
                          onClick={() => {
                            handleStartEditCandidate(cand);
                          }}
                        >
                          Editar
                        </button>
                        <button 
                          type="button"
                          className="btn btn-danger" 
                          onClick={() => {
                            handleDeleteCandidate(cand.id);
                          }}
                        >
                          Eliminar
                        </button>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}

      {/* TAB VOTERS LIST */}
      {activeTab === 'voters' && (
        <div>
          <h2 style={{ fontFamily: 'var(--font-heading)', fontWeight: 700, marginBottom: '20px' }}>Gestión de Votantes Habilitados</h2>

          <div style={{ padding: '12px 16px', backgroundColor: 'rgba(59, 130, 246, 0.05)', borderLeft: '4px solid var(--primary)', borderRadius: '0 var(--radius-md) var(--radius-md) 0', fontSize: '13px', color: 'var(--text-secondary)', marginBottom: '20px', display: 'flex', gap: '10px', alignItems: 'center', animation: 'fadeIn 0.2s ease-out' }}>
            <span style={{ fontSize: '16px' }}>ℹ️</span>
            <span><strong>Nota de Sincronización:</strong> Si agregas votantes, cambias su asistencia o reinicias sus votos, los miembros que ya tengan la página abierta necesitarán <strong>actualizar/recargar</strong> su pantalla de votación para recibir los cambios.</span>
          </div>

          <div className="grid-2" style={{ marginBottom: '24px' }}>
            {/* Agregar Individual */}
            <form onSubmit={handleAddSingleVoter} className="card" style={{ border: '1px solid var(--border)' }}>
              <h3 style={{ fontFamily: 'var(--font-heading)', fontWeight: 600, fontSize: '17px', marginBottom: '16px' }}>Agregar un Votante</h3>
              <div className="form-group" style={{ marginBottom: '12px' }}>
                <label className="form-label">Nombre</label>
                <input 
                  type="text" 
                  className="form-control" 
                  placeholder="Nombre"
                  value={voterFirstName}
                  onChange={(e) => setVoterFirstName(e.target.value)}
                  required
                />
              </div>
              <div className="form-group" style={{ marginBottom: '16px' }}>
                <label className="form-label">Apellido</label>
                <input 
                  type="text" 
                  className="form-control" 
                  placeholder="Apellido"
                  value={voterLastName}
                  onChange={(e) => setVoterLastName(e.target.value)}
                  required
                />
              </div>
              <Tooltip text="Añadir a un nuevo miembro autorizado al censo de votación.">
                <button type="submit" className="btn btn-primary" style={{ width: '100%' }}>Agregar</button>
              </Tooltip>
            </form>

            {/* Agregar Masivo */}
            <form onSubmit={handleAddBulkVoters} className="card" style={{ border: '1px solid var(--border)' }}>
              <h3 style={{ fontFamily: 'var(--font-heading)', fontWeight: 600, fontSize: '17px', marginBottom: '16px' }}>Carga Masiva (Lote)</h3>
              <div className="form-group">
                <label className="form-label">Nombres (separados por comas o saltos de línea)</label>
                <textarea 
                  className="form-control" 
                  rows="3" 
                  placeholder="Juan Pérez&#10;Ana Gómez&#10;Pedro Rojas"
                  value={bulkVotersText}
                  onChange={(e) => setBulkVotersText(e.target.value)}
                ></textarea>
              </div>
              <Tooltip text="Cargar múltiples miembros al censo de forma masiva usando texto.">
                <button type="submit" className="btn btn-secondary" style={{ width: '100%' }}>Agregar Lista</button>
              </Tooltip>
            </form>
          </div>

          <div className="card" style={{ border: '1px solid var(--border)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
              <h3 style={{ fontFamily: 'var(--font-heading)', fontWeight: 600, fontSize: '18px', margin: 0 }}>Votantes Registrados ({voters.length})</h3>
              <Tooltip text="Eliminar a todos los miembros autorizados de la lista de votantes." position="left">
                <button className="btn btn-danger" onClick={handleClearAllVoters} style={{ padding: '6px 12px', fontSize: '12px' }}>
                  Vaciar Lista
                </button>
              </Tooltip>
            </div>

            {/* Buscador de Votantes Global (Fuera del contenedor scroll) */}
            <div style={{ marginBottom: '16px', position: 'relative' }}>
              <input 
                type="text" 
                className="form-control" 
                placeholder="🔍 Buscar votante por nombre o apellido..."
                value={voterSearch}
                onChange={(e) => setVoterSearch(e.target.value)}
                style={{ 
                  margin: 0, 
                  height: '42px', 
                  fontSize: '14px', 
                  padding: '8px 12px 8px 38px', 
                  borderRadius: 'var(--radius-md)', 
                  border: '1px solid var(--border)',
                  width: '100%',
                  boxSizing: 'border-box'
                }}
              />
              {voterSearch && (
                <button 
                  type="button"
                  onClick={() => setVoterSearch('')}
                  style={{ 
                    position: 'absolute', 
                    right: '12px', 
                    top: '50%', 
                    transform: 'translateY(-50%)', 
                    border: 'none', 
                    background: 'none', 
                    fontSize: '18px', 
                    color: 'var(--text-muted)', 
                    cursor: 'pointer',
                    padding: '4px'
                  }}
                >
                  ×
                </button>
              )}
            </div>



            {/* Filtros Activos e Informativos (Excel-like) */}
            {(voterSearch || voterPresenceFilter !== 'all' || voterProgressFilter !== 'all') && (
              <div style={{ display: 'flex', gap: '8px', alignItems: 'center', marginBottom: '16px', flexWrap: 'wrap', animation: 'fadeIn 0.2s ease-out' }}>
                <span style={{ fontSize: '13px', fontWeight: '600', color: 'var(--text-secondary)' }}>Filtros aplicados:</span>
                {voterSearch && <span className="badge badge-info" style={{ fontSize: '11px', padding: '4px 8px', backgroundColor: 'rgba(59, 130, 246, 0.1)', color: 'var(--primary)' }}>Nombre: "{voterSearch}"</span>}
                {voterPresenceFilter !== 'all' && <span className="badge badge-info" style={{ fontSize: '11px', padding: '4px 8px', backgroundColor: 'rgba(59, 130, 246, 0.1)', color: 'var(--primary)' }}>Asistencia: {voterPresenceFilter === 'present' ? 'Presentes' : 'Ausentes'}</span>}
                {voterProgressFilter !== 'all' && <span className="badge badge-info" style={{ fontSize: '11px', padding: '4px 8px', backgroundColor: 'rgba(59, 130, 246, 0.1)', color: 'var(--primary)' }}>Estado: {voterProgressFilter === 'pending' ? 'Sin Votar' : voterProgressFilter === 'partial' ? 'Parcial' : 'Ya Votó'}</span>}
                <button 
                  type="button" 
                  className="btn-link" 
                  onClick={() => {
                    setVoterSearch('');
                    setVoterPresenceFilter('all');
                    setVoterProgressFilter('all');
                  }}
                  style={{ fontSize: '12px', color: 'var(--danger)', cursor: 'pointer', background: 'none', border: 'none', textDecoration: 'underline', padding: 0 }}
                >
                  Limpiar todos los filtros
                </button>
              </div>
            )}

            {/* Barra de Acciones Masivas */}
            {selectedVoterIds.length > 0 && (
              <div className="bulk-actions-bar" style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: '12px', padding: '12px 16px', backgroundColor: 'var(--primary-light)', border: '1px solid var(--primary)', borderRadius: 'var(--radius-md)', marginBottom: '16px', animation: 'fadeIn 0.2s ease-out' }}>
                <span style={{ fontSize: '14px', fontWeight: '700', color: 'var(--primary)' }}>
                  {selectedVoterIds.length} seleccionados
                </span>
                <div style={{ display: 'flex', gap: '8px', marginLeft: 'auto' }}>
                  <Tooltip text="Borrar los votos de los miembros seleccionados para que puedan votar de nuevo.">
                    <button 
                      type="button" 
                      className="btn btn-secondary" 
                      onClick={handleBulkResetVotes}
                      disabled={isUploading}
                      style={{ padding: '6px 12px', fontSize: '12px', border: '1px solid var(--border)' }}
                    >
                      Restablecer Votos
                    </button>
                  </Tooltip>
                  <Tooltip text="Marcar como presentes (habilitar para votar) a todos los seleccionados.">
                    <button 
                      type="button" 
                      className="btn btn-secondary" 
                      onClick={() => handleBulkTogglePresence(true)}
                      disabled={isUploading}
                      style={{ padding: '6px 12px', fontSize: '12px', border: '1px solid var(--border)', borderColor: 'var(--success)', color: 'var(--success)' }}
                    >
                      Marcar Presentes
                    </button>
                  </Tooltip>
                  <Tooltip text="Marcar como ausentes (deshabilitar para votar) a todos los seleccionados.">
                    <button 
                      type="button" 
                      className="btn btn-secondary" 
                      onClick={() => handleBulkTogglePresence(false)}
                      disabled={isUploading}
                      style={{ padding: '6px 12px', fontSize: '12px', border: '1px solid var(--border)', borderColor: 'var(--danger)', color: 'var(--danger)' }}
                    >
                      Marcar Ausentes
                    </button>
                  </Tooltip>
                  <Tooltip text="Quitar permanentemente a los miembros seleccionados del censo de votación.">
                    <button 
                      type="button" 
                      className="btn btn-danger" 
                      onClick={handleBulkDeleteVoters}
                      disabled={isUploading}
                      style={{ padding: '6px 12px', fontSize: '12px' }}
                    >
                      Eliminar Votantes
                    </button>
                  </Tooltip>
                  <Tooltip text="Limpiar la selección de miembros actual.">
                    <button 
                      type="button" 
                      className="btn btn-secondary"
                      onClick={() => setSelectedVoterIds([])}
                      style={{ padding: '6px 12px', fontSize: '12px', border: '1px solid var(--border)' }}
                    >
                      Deseleccionar
                    </button>
                  </Tooltip>
                </div>
              </div>
            )}

            {filteredVoters.length === 0 ? (
              <p style={{ color: 'var(--text-muted)', textAlign: 'center', padding: '24px 0', border: '1px solid var(--border)', borderRadius: 'var(--radius-md)', margin: 0 }}>No se encontraron votantes.</p>
            ) : (
              <div className="votos-table-container" style={{ margin: 0, maxHeight: '450px', minHeight: '250px', overflowY: 'auto', border: '1px solid var(--border)', borderRadius: 'var(--radius-md)' }}>
                <table className="votos-table" style={{ margin: 0 }}>
                  <thead>
                    <tr style={{ position: 'sticky', top: 0, zIndex: 10, background: 'var(--bg-card)' }}>
                      <th style={{ width: '40px', padding: '12px 16px' }}>
                        <input 
                          type="checkbox" 
                          checked={filteredVoters.length > 0 && filteredVoters.every(v => selectedVoterIds.some(id => String(id) === String(v.id)))}
                          onChange={() => handleToggleSelectAll(filteredVoters.length > 0 && filteredVoters.every(v => selectedVoterIds.some(id => String(id) === String(v.id))))}
                          style={{ cursor: 'pointer', width: '16px', height: '16px' }}
                        />
                      </th>
                      
                      {/* Columna Nombre */}
                      <th style={{ padding: '12px 16px', position: 'relative' }}>
                        <div 
                          className="excel-filter-trigger"
                          style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', cursor: 'pointer', userSelect: 'none' }}
                          onClick={() => setActiveHeaderFilter(activeHeaderFilter === 'name' ? null : 'name')}
                        >
                          <span>Nombre y Apellido</span>
                          <span style={{ fontSize: '11px', color: voterSearch || voterSortKey === 'name' ? 'var(--primary)' : 'var(--text-muted)' }}>
                            {voterSortKey === 'name' ? (voterSortDirection === 'asc' ? '▲' : '▼') : '⛛'}
                          </span>
                        </div>
                        {activeHeaderFilter === 'name' && (
                          <div className="excel-filter-dropdown" style={{ position: 'absolute', top: '100%', left: '16px', background: 'var(--bg-card)', border: '1px solid var(--border)', borderRadius: 'var(--radius-md)', padding: '10px', zIndex: 999, width: '180px', boxShadow: '0 4px 16px rgba(0,0,0,0.15)', display: 'flex', flexDirection: 'column', gap: '8px' }} onClick={(e) => e.stopPropagation()}>
                            <button 
                              type="button"
                              className="btn btn-secondary"
                              onClick={() => {
                                setVoterSortKey('name');
                                setVoterSortDirection('asc');
                                setActiveHeaderFilter(null);
                              }}
                              style={{ width: '100%', justifyContent: 'flex-start', padding: '6px 12px', fontSize: '13px', height: '32px' }}
                            >
                              🔼 Ordenar A → Z
                            </button>
                            <button 
                              type="button"
                              className="btn btn-secondary"
                              onClick={() => {
                                setVoterSortKey('name');
                                setVoterSortDirection('desc');
                                setActiveHeaderFilter(null);
                              }}
                              style={{ width: '100%', justifyContent: 'flex-start', padding: '6px 12px', fontSize: '13px', height: '32px' }}
                            >
                              🔽 Ordenar Z → A
                            </button>
                          </div>
                        )}
                      </th>

                      {/* Columna Asistencia */}
                      <th style={{ padding: '12px 16px', position: 'relative' }}>
                        <div 
                          className="excel-filter-trigger"
                          style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', cursor: 'pointer', userSelect: 'none' }}
                          onClick={() => setActiveHeaderFilter(activeHeaderFilter === 'presence' ? null : 'presence')}
                        >
                          <span>Asistencia</span>
                          <span style={{ fontSize: '11px', color: voterPresenceFilter !== 'all' || voterSortKey === 'presence' ? 'var(--primary)' : 'var(--text-muted)' }}>
                            {voterSortKey === 'presence' ? (voterSortDirection === 'asc' ? '▲' : '▼') : '⛛'}
                          </span>
                        </div>
                        {activeHeaderFilter === 'presence' && (
                          <div className="excel-filter-dropdown" style={{ position: 'absolute', top: '100%', left: '16px', background: 'var(--bg-card)', border: '1px solid var(--border)', borderRadius: 'var(--radius-md)', padding: '14px', zIndex: 999, width: '200px', boxShadow: '0 4px 16px rgba(0,0,0,0.15)', display: 'flex', flexDirection: 'column', gap: '10px' }} onClick={(e) => e.stopPropagation()}>
                            <button 
                              type="button"
                              className="btn btn-secondary"
                              onClick={() => {
                                setVoterSortKey('presence');
                                setVoterSortDirection('asc');
                                setActiveHeaderFilter(null);
                              }}
                              style={{ width: '100%', justifyContent: 'flex-start', padding: '6px 12px', fontSize: '13px', height: '32px' }}
                            >
                              📈 Presentes primero
                            </button>
                            <button 
                              type="button"
                              className="btn btn-secondary"
                              onClick={() => {
                                setVoterSortKey('presence');
                                setVoterSortDirection('desc');
                                setActiveHeaderFilter(null);
                              }}
                              style={{ width: '100%', justifyContent: 'flex-start', padding: '6px 12px', fontSize: '13px', height: '32px' }}
                            >
                              📉 Ausentes primero
                            </button>
                            <hr style={{ border: 0, borderTop: '1px solid var(--border)', margin: '4px 0' }} />
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                              <label style={{ fontSize: '11px', color: 'var(--text-secondary)', fontWeight: 600, display: 'block' }}>Filtrar valor:</label>
                              {[
                                { value: 'all', label: 'Todos' },
                                { value: 'present', label: 'Presentes' },
                                { value: 'absent', label: 'Ausentes' }
                              ].map(opt => (
                                <label key={opt.value} style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px', cursor: 'pointer', margin: 0, fontWeight: 'normal', color: 'var(--text-primary)' }}>
                                  <input 
                                    type="radio" 
                                    name="presenceFilterOpt" 
                                    checked={voterPresenceFilter === opt.value}
                                    onChange={() => {
                                      setVoterPresenceFilter(opt.value);
                                      setSelectedVoterIds([]);
                                    }}
                                    style={{ cursor: 'pointer', width: '14px', height: '14px', margin: 0 }}
                                  />
                                  <span>{opt.label}</span>
                                </label>
                              ))}
                            </div>
                          </div>
                        )}
                      </th>

                      {/* Columna Progreso */}
                      <th style={{ padding: '12px 16px', position: 'relative' }}>
                        <div 
                          className="excel-filter-trigger"
                          style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', cursor: 'pointer', userSelect: 'none' }}
                          onClick={() => setActiveHeaderFilter(activeHeaderFilter === 'progress' ? null : 'progress')}
                        >
                          <span>Progreso de Votación</span>
                          <span style={{ fontSize: '11px', color: voterProgressFilter !== 'all' || voterSortKey === 'progress' ? 'var(--primary)' : 'var(--text-muted)' }}>
                            {voterSortKey === 'progress' ? (voterSortDirection === 'asc' ? '▲' : '▼') : '⛛'}
                          </span>
                        </div>
                        {activeHeaderFilter === 'progress' && (
                          <div className="excel-filter-dropdown" style={{ position: 'absolute', top: '100%', left: '16px', background: 'var(--bg-card)', border: '1px solid var(--border)', borderRadius: 'var(--radius-md)', padding: '14px', zIndex: 999, width: '220px', boxShadow: '0 4px 16px rgba(0,0,0,0.15)', display: 'flex', flexDirection: 'column', gap: '10px' }} onClick={(e) => e.stopPropagation()}>
                            <button 
                              type="button"
                              className="btn btn-secondary"
                              onClick={() => {
                                setVoterSortKey('progress');
                                setVoterSortDirection('asc');
                                setActiveHeaderFilter(null);
                              }}
                              style={{ width: '100%', justifyContent: 'flex-start', padding: '6px 12px', fontSize: '13px', height: '32px' }}
                            >
                              📈 Completos primero
                            </button>
                            <button 
                              type="button"
                              className="btn btn-secondary"
                              onClick={() => {
                                setVoterSortKey('progress');
                                setVoterSortDirection('desc');
                                setActiveHeaderFilter(null);
                              }}
                              style={{ width: '100%', justifyContent: 'flex-start', padding: '6px 12px', fontSize: '13px', height: '32px' }}
                            >
                              📉 Pendientes primero
                            </button>
                            <hr style={{ border: 0, borderTop: '1px solid var(--border)', margin: '4px 0' }} />
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                              <label style={{ fontSize: '11px', color: 'var(--text-secondary)', fontWeight: 600, display: 'block' }}>Filtrar valor:</label>
                              {[
                                { value: 'all', label: 'Todos' },
                                { value: 'pending', label: 'Sin Votar' },
                                { value: 'partial', label: 'Parcial' },
                                { value: 'completed', label: 'Ya Votó' }
                              ].map(opt => (
                                <label key={opt.value} style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px', cursor: 'pointer', margin: 0, fontWeight: 'normal', color: 'var(--text-primary)' }}>
                                  <input 
                                    type="radio" 
                                    name="progressFilterOpt" 
                                    checked={voterProgressFilter === opt.value}
                                    onChange={() => {
                                      setVoterProgressFilter(opt.value);
                                      setSelectedVoterIds([]);
                                    }}
                                    style={{ cursor: 'pointer', width: '14px', height: '14px', margin: 0 }}
                                  />
                                  <span>{opt.label}</span>
                                </label>
                              ))}
                            </div>
                          </div>
                        )}
                      </th>
                      
                      <th style={{ padding: '12px 16px' }}>Acciones</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sortedVoters.map((voter, index) => {
                      const progress = getVoterProgress(voter);
                      const isChecked = selectedVoterIds.some(id => String(id) === String(voter.id));
                      return (
                        <tr key={`${voter.id || 'voter'}-${index}`}>
                          <td data-label="Seleccionar" style={{ padding: '10px 16px' }}>
                            <input 
                              type="checkbox" 
                              checked={isChecked} 
                              onChange={(e) => {
                                const voterIdStr = String(voter.id);
                                if (e.target.checked) {
                                  setSelectedVoterIds([...selectedVoterIds.map(id => String(id)), voterIdStr]);
                                } else {
                                  setSelectedVoterIds(selectedVoterIds.filter(id => String(id) !== voterIdStr));
                                }
                              }}
                              style={{ cursor: 'pointer', width: '16px', height: '16px' }}
                            />
                          </td>
                          <td data-label="Nombre y Apellido" style={{ padding: '10px 16px', fontWeight: 600 }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                              <div style={{
                                width: '28px',
                                height: '28px',
                                borderRadius: '50%',
                                backgroundColor: 'var(--primary-light)',
                                color: 'var(--primary)',
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'center',
                                fontWeight: 'bold',
                                fontSize: '11px',
                                flexShrink: 0
                              }}>
                                {voter.name?.charAt(0)}{voter.lastName?.charAt(0)}
                              </div>
                              <span>{voter.name} {voter.lastName}</span>
                            </div>
                          </td>
                          <td data-label="Asistencia" style={{ padding: '10px 16px' }}>
                            <Tooltip text={voter.isPresent !== false ? "Marcar como ausente (no podrá votar)" : "Marcar como presente (habilitar para votar)"} position="top">
                              <label style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', cursor: 'pointer', userSelect: 'none', fontSize: '13px', margin: 0 }}>
                                <input 
                                  type="checkbox" 
                                  checked={voter.isPresent !== false}
                                  onChange={(e) => handleToggleVoterPresence(voter, e.target.checked)}
                                  style={{ cursor: 'pointer', width: '15px', height: '15px', margin: 0 }}
                                />
                                <span style={{ color: voter.isPresent !== false ? 'var(--success)' : 'var(--text-muted)', fontWeight: voter.isPresent !== false ? '600' : 'normal' }}>
                                  {voter.isPresent !== false ? 'Presente' : 'Ausente'}
                                </span>
                              </label>
                            </Tooltip>
                          </td>
                          <td data-label="Estado Votación" style={{ padding: '10px 16px' }}>
                            {progress.isComplete ? (
                              <span className="badge badge-success">Ya Votó</span>
                            ) : progress.hasVoted ? (
                              <Tooltip text={progress.pendingCandidates.length > 0 ? `Pendientes por votar:\n${progress.pendingCandidates.map(c => `• ${c.firstName} ${c.lastName}`).join('\n')}` : 'No hay candidatos pendientes'} position="left">
                                <span className="badge badge-info" style={{ backgroundColor: 'rgba(59, 130, 246, 0.1)', color: 'var(--primary)', cursor: 'help' }}>Parcial ({progress.votedCount}/{progress.totalCount})</span>
                              </Tooltip>
                            ) : (
                              <Tooltip text={progress.pendingCandidates.length > 0 ? `Pendientes por votar:\n${progress.pendingCandidates.map(c => `• ${c.firstName} ${c.lastName}`).join('\n')}` : 'No hay candidatos pendientes'} position="left">
                                <span className="badge badge-danger" style={{ backgroundColor: 'rgba(239, 68, 68, 0.1)', color: 'var(--danger)', cursor: 'help' }}>Sin Votación</span>
                              </Tooltip>
                            )}
                          </td>
                          <td data-label="Acciones" style={{ padding: '10px 16px' }}>
                            <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap' }}>
                              {progress.hasVoted && (
                                <Tooltip text="Borrar los votos de este miembro para permitirle votar de nuevo.">
                                  <button 
                                    type="button"
                                    className="btn btn-secondary" 
                                    onClick={() => handleResetVoterVotes(voter)}
                                    style={{ padding: '4px 10px', fontSize: '11px', height: '28px', display: 'flex', alignItems: 'center', border: '1px solid var(--border)' }}
                                  >
                                    Restablecer
                                  </button>
                                </Tooltip>
                              )}
                              <Tooltip text="Retirar a este miembro de la lista de votantes autorizados.">
                                <button 
                                  type="button" 
                                  className="btn btn-danger"
                                  onClick={() => handleDeleteVoter(voter.id)}
                                  style={{ padding: '4px 8px', fontSize: '12px', height: '28px', display: 'flex', alignItems: 'center', justifyContent: 'center' }}
                                >
                                  Eliminar
                                </button>
                              </Tooltip>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>
      )}

      {/* TAB SHARE / QR */}
      {activeTab === 'share' && (
        <div>
          <h2 style={{ fontFamily: 'var(--font-heading)', fontWeight: 700, marginBottom: '20px' }}>Compartir Sistema de Votación</h2>
          
          <div className="card" style={{ border: '1px solid var(--border)', display: 'flex', flexDirection: 'column', alignItems: 'center', padding: '32px' }}>
            <p style={{ textAlign: 'center', color: 'var(--text-secondary)', maxWidth: '500px', marginBottom: '24px' }}>
              Las personas pueden ingresar al panel de votación escaneando el código QR generado abajo desde sus teléfonos celulares.
            </p>

            <div className="qr-container">
              <div className="qr-image-wrapper">
                <img 
                  src={qrCodeUrl} 
                  className="qr-image" 
                  alt="Código QR de Votación" 
                  onError={handleQrError}
                />
              </div>
              
              <div className="qr-url-text">{votingUrl}</div>
              
              <div style={{ display: 'flex', gap: '12px' }}>
                <Tooltip text="Copiar el enlace de votación al portapapeles para compartirlo.">
                  <button className="btn btn-secondary" onClick={handleCopyVotingLink}>
                    Copiar Enlace
                  </button>
                </Tooltip>
                <Tooltip text="Guardar la imagen del código QR de votación en tu dispositivo.">
                  <a href={qrCodeUrl} download="qr-votacion-icc.png" target="_blank" rel="noreferrer" className="btn btn-primary">
                    Descargar Código QR
                  </a>
                </Tooltip>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* TAB CONFIGURATION REMOVIDA DE LA UI */}

      {cropImageFile && (
        <ImageCropper 
          imageFile={cropImageFile}
          onCrop={(croppedFile) => {
            setCandPhotoFile(croppedFile);
            setCropImageFile(null);
          }}
          onCancel={() => {
            setCropImageFile(null);
            if (fileInputRef.current) fileInputRef.current.value = '';
          }}
        />
      )}
    </div>
  );
}

// ----------------------------------------------------
// FUNCIONES AUXILIARES PARA EXPORTACIÓN DE EXCEL
// ----------------------------------------------------

// Estila el fondo y bordes de celdas combinadas de tarjetas ANTES de realizar el merge
function styleAndMergeCard(ws, range, title, value) {
  const [start, end] = range.split(':');
  const startCol = start.charCodeAt(0) - 65 + 1; // 'B' -> 2
  const startRow = parseInt(start.substring(1));
  const endCol = (end || start).charCodeAt(0) - 65 + 1;
  const endRow = parseInt((end || start).substring(1));

  // 1. Estilar todas las celdas en el rango individualmente
  for (let r = startRow; r <= endRow; r++) {
    for (let c = startCol; c <= endCol; c++) {
      const cell = ws.getCell(r, c);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF8FAFC' } // Slate 50
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFCBD5E1' } }, // Slate 300
        left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
      };
    }
  }

  // 2. Realizar el merge de celdas si es un rango
  if (range.includes(':')) {
    ws.mergeCells(range);
  }

  // 3. Escribir el valor alineado al centro en la celda maestra (esquina superior izquierda)
  const masterCell = ws.getCell(startRow, startCol);
  masterCell.value = {
    richText: [
      { text: title + '\n', font: { name: 'Segoe UI', size: 9, color: { argb: 'FF64748B' } } },
      { text: value, font: { name: 'Segoe UI', size: 16, bold: true, color: { argb: 'FF0F172A' } } }
    ]
  };
  masterCell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
}

// Dibuja un gráfico de barras agrupadas en un canvas temporal y devuelve la imagen Base64
function generateChartImage(candidates, getCandidateStats) {
  const canvas = document.createElement('canvas');
  canvas.width = 800;
  canvas.height = 430;
  const ctx = canvas.getContext('2d');

  // Fondo blanco limpio
  ctx.fillStyle = '#FFFFFF';
  ctx.fillRect(0, 0, canvas.width, canvas.height);

  // Agregar un borde sutil al gráfico para dar estructura de tarjeta
  ctx.strokeStyle = '#F1F5F9';
  ctx.lineWidth = 4;
  ctx.strokeRect(2, 2, canvas.width - 4, canvas.height - 4);

  const paddingLeft = 80;
  const paddingRight = 40;
  const paddingTop = 70;
  const paddingBottom = 80;
  const graphWidth = canvas.width - paddingLeft - paddingRight;
  const graphHeight = canvas.height - paddingTop - paddingBottom;

  // Título del Gráfico
  ctx.fillStyle = '#0F172A'; // Slate 900
  ctx.font = 'bold 18px "Segoe UI", Arial, sans-serif';
  ctx.textAlign = 'left';
  ctx.fillText('Resultados de Votación: Aprobaciones vs Objeciones', paddingLeft, 38);

  // Leyenda
  const legendX = canvas.width - 260;
  const legendY = 25;
  
  // Leyenda Aprueba
  ctx.fillStyle = '#10B981'; // Emerald 500
  ctx.fillRect(legendX, legendY, 15, 15);
  ctx.fillStyle = '#475569'; // Slate 600
  ctx.font = '13px "Segoe UI", Arial, sans-serif';
  ctx.fillText('Aprueba', legendX + 22, legendY + 12);

  // Leyenda No Aprueba
  ctx.fillStyle = '#EF4444'; // Red 500
  ctx.fillRect(legendX + 110, legendY, 15, 15);
  ctx.fillStyle = '#475569';
  ctx.fillText('No Aprueba', legendX + 132, legendY + 12);

  // Datos
  const data = candidates.map(c => {
    const stats = getCandidateStats(c.id);
    return {
      name: `${c.firstName || ''} ${c.lastName || ''}`.trim(),
      approves: stats.approves,
      disapproves: stats.disapproves,
      total: stats.total
    };
  });

  // Encontrar el valor máximo para escalar el eje Y
  let maxVal = 5;
  data.forEach(d => {
    const candMax = Math.max(d.approves, d.disapproves);
    if (candMax > maxVal) {
      maxVal = candMax;
    }
  });
  maxVal = Math.ceil(maxVal / 5) * 5;

  // Dibujar líneas de cuadrícula y etiquetas del eje Y
  const yTicks = 5;
  ctx.strokeStyle = '#F1F5F9'; // Slate 100
  ctx.lineWidth = 1.5;
  ctx.fillStyle = '#64748B'; // Slate 500
  ctx.font = '12px "Segoe UI", Arial, sans-serif';
  ctx.textAlign = 'right';

  for (let i = 0; i <= yTicks; i++) {
    const val = (maxVal / yTicks) * i;
    const y = paddingTop + graphHeight - (i * (graphHeight / yTicks));
    
    // Línea de cuadrícula
    ctx.beginPath();
    ctx.moveTo(paddingLeft, y);
    ctx.lineTo(canvas.width - paddingRight, y);
    ctx.stroke();

    // Etiqueta
    ctx.fillText(Math.round(val).toString(), paddingLeft - 12, y + 4);
  }

  // Ejes X y Y
  ctx.strokeStyle = '#CBD5E1'; // Slate 300
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(paddingLeft, paddingTop);
  ctx.lineTo(paddingLeft, paddingTop + graphHeight);
  ctx.lineTo(canvas.width - paddingRight, paddingTop + graphHeight);
  ctx.stroke();

  // Función auxiliar para dibujar barras con esquinas superiores redondeadas
  function drawRoundedBar(cCtx, x, y, width, height, radius) {
    if (height <= 0) return;
    cCtx.beginPath();
    cCtx.moveTo(x + radius, y);
    cCtx.lineTo(x + width - radius, y);
    cCtx.quadraticCurveTo(x + width, y, x + width, y + radius);
    cCtx.lineTo(x + width, y + height);
    cCtx.lineTo(x, y + height);
    cCtx.lineTo(x, y + radius);
    cCtx.quadraticCurveTo(x, y, x + radius, y);
    cCtx.closePath();
    cCtx.fill();
  }

  // Dibujar Barras
  const numGroups = data.length;
  const groupWidth = graphWidth / numGroups;
  const barWidth = groupWidth * 0.33; // 33% del ancho de grupo
  const groupGap = groupWidth * 0.14; // Separación a los lados

  data.forEach((d, idx) => {
    const groupX = paddingLeft + (idx * groupWidth) + groupGap;
    const cornerRad = Math.min(5, barWidth / 2);

    // Barra de "Aprueba" (Verde)
    const appHeight = d.total > 0 ? (d.approves / maxVal) * graphHeight : 0;
    const appX = groupX;
    const appY = paddingTop + graphHeight - appHeight;
    
    if (appHeight > 0) {
      ctx.fillStyle = '#10B981';
      drawRoundedBar(ctx, appX, appY, barWidth, appHeight, cornerRad);
      
      // Número encima de la barra
      ctx.fillStyle = '#065F46';
      ctx.font = 'bold 12px "Segoe UI", Arial, sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText(d.approves.toString(), appX + barWidth / 2, appY - 6);
    }

    // Barra de "No Aprueba" (Rojo)
    const disHeight = d.total > 0 ? (d.disapproves / maxVal) * graphHeight : 0;
    const disX = groupX + barWidth + 6; // Espacio de 6px entre barras
    const disY = paddingTop + graphHeight - disHeight;

    if (disHeight > 0) {
      ctx.fillStyle = '#EF4444';
      drawRoundedBar(ctx, disX, disY, barWidth, disHeight, cornerRad);

      // Número encima de la barra
      ctx.fillStyle = '#991B1B';
      ctx.font = 'bold 12px "Segoe UI", Arial, sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText(d.disapproves.toString(), disX + barWidth / 2, disY - 6);
    }

    // Eje X: Nombre del candidato
    ctx.fillStyle = '#334155'; // Slate 700
    ctx.font = 'bold 12px "Segoe UI", Arial, sans-serif';
    ctx.textAlign = 'center';
    
    const labelX = groupX + barWidth + 3;
    const labelY = paddingTop + graphHeight + 22;

    const nameParts = d.name.split(' ');
    if (nameParts.length > 2) {
      const line1 = nameParts.slice(0, 2).join(' ');
      const line2 = nameParts.slice(2).join(' ');
      ctx.fillText(line1, labelX, labelY);
      ctx.fillText(line2, labelX, labelY + 16);
    } else {
      ctx.fillText(d.name, labelX, labelY);
    }
  });

  return canvas.toDataURL('image/png');
}
