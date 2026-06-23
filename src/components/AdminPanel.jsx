import { useState, useRef, useEffect } from 'react';
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

  // Estados de Gestión de Votantes
  const [voterFirstName, setVoterFirstName] = useState('');
  const [voterLastName, setVoterLastName] = useState('');
  const [bulkVotersText, setBulkVotersText] = useState('');
  const [voterSearch, setVoterSearch] = useState('');
  const [voterStatusFilter, setVoterStatusFilter] = useState('all'); // 'all' | 'pending' | 'partial' | 'completed'
  const [selectedVoterIds, setSelectedVoterIds] = useState([]);

  // Filtro de resultados
  const [resultsFilter, setResultsFilter] = useState('all');

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
    membersSheet.appendRow(["ID", "Nombre", "Apellido"]);
  }
  
  var candidatesSheet = sheet.getSheetByName("Candidatos");
  if (!candidatesSheet) {
    candidatesSheet = sheet.insertSheet("Candidatos");
    candidatesSheet.appendRow(["ID", "Nombre", "Apellido", "Testimonio", "Foto"]);
  } else {
    // Asegurar columna Foto
    var headers = candidatesSheet.getRange(1, 1, 1, candidatesSheet.getLastColumn()).getValues()[0];
    if (headers.indexOf("Foto") === -1) {
      candidatesSheet.getRange(1, headers.length + 1).setValue("Foto");
    }
  }
  
  if (!sheet.getSheetByName("Votos")) {
    var newSheet = sheet.insertSheet("Votos");
    newSheet.appendRow(["ID Candidato", "Nombre Candidato", "Apellido Candidato", "Estado", "Motivo", "ID Votante", "Timestamp"]);
  }

  // Asegurar columnas Has Voted y Fecha Voto en hoja Miembros
  var membersHeaders = membersSheet.getRange(1, 1, 1, membersSheet.getLastColumn()).getValues()[0];
  if (membersHeaders.indexOf("Has Voted") === -1) {
    membersSheet.getRange(1, membersHeaders.length + 1).setValue("Has Voted");
    membersSheet.getRange(1, membersHeaders.length + 2).setValue("Fecha Voto");
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
    membersSheet.appendRow(["ID", "Nombre", "Apellido"]);
  }
  
  var candidatesSheet = sheet.getSheetByName("Candidatos");
  if (!candidatesSheet) {
    candidatesSheet = sheet.insertSheet("Candidatos");
    candidatesSheet.appendRow(["ID", "Nombre", "Apellido", "Testimonio", "Foto"]);
  } else {
    // Asegurar columna Foto
    var headers = candidatesSheet.getRange(1, 1, 1, candidatesSheet.getLastColumn()).getValues()[0];
    if (headers.indexOf("Foto") === -1) {
      candidatesSheet.getRange(1, headers.length + 1).setValue("Foto");
    }
  }
  
  if (!sheet.getSheetByName("Votos")) {
    var newSheet = sheet.insertSheet("Votos");
    newSheet.appendRow(["ID Candidato", "Nombre Candidato", "Apellido Candidato", "Estado", "Motivo", "ID Votante", "Timestamp"]);
  }

  // Asegurar columnas Has Voted y Fecha Voto en hoja Miembros
  var membersSheet2 = sheet.getSheetByName("Miembros") || sheet.getSheetByName("Votantes");
  if (membersSheet2) {
    var mHeaders = membersSheet2.getRange(1, 1, 1, membersSheet2.getLastColumn()).getValues()[0];
    if (mHeaders.indexOf("Has Voted") === -1) {
      membersSheet2.getRange(1, mHeaders.length + 1).setValue("Has Voted");
      membersSheet2.getRange(1, mHeaders.length + 2).setValue("Fecha Voto");
    }
  }

  var data = JSON.parse(e.postData.contents);
  var action = data.action;

  if (action === "updateVoters") {
    membersSheet.clear();
    membersSheet.appendRow(["ID", "Nombre", "Apellido", "Has Voted", "Fecha Voto"]);
    data.voters.forEach(function(v) {
      membersSheet.appendRow([
        v.id || v.ID || "", 
        v.name || v.Nombre || "", 
        v.lastName || v.Apellido || "",
        v.hasVoted === true ? true : false,
        v.votedAt || ""
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
      lastName: lName
    };

    const updatedVoters = [...voters, newVoter];

    try {
      if (config.sheetUrlVoters) {
        if (!isConnected) {
          throw new Error("No hay conexión con Google Sheets. No se puede agregar el votante.");
        }
        await updateVotersInSheets(config.sheetUrlVoters, updatedVoters);
        showToast("Votante registrado en Google Sheets", "success");
        refreshData();
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
          lastName
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
        showToast("Votante eliminado", "success");
        refreshData();
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

  // Limpiar lista de votantes
  const handleClearAllVoters = async () => {
    if (!window.confirm("¿Seguro que deseas eliminar a TODOS los votantes?")) return;
    try {
      if (config.sheetUrlVoters) {
        if (!isConnected) {
          throw new Error("No hay conexión con Google Sheets. No se puede vaciar la lista de votantes.");
        }
        await updateVotersInSheets(config.sheetUrlVoters, []);
        showToast("Lista de votantes vaciada", "success");
        refreshData();
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
  // Ahora usa el campo hasVoted del votante en lugar de contar votos por nombre
  const getVoterProgress = (voter) => {
    const participated = voter?.hasVoted === true;
    return {
      votedCount: participated ? candidates.length : 0,
      totalCount: candidates.length,
      hasVoted: participated,
      isComplete: participated,
      pendingCandidates: participated ? [] : [...candidates]
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
        showToast(`Participación de ${fullName} restablecida en Google Sheets`, "success");
        refreshData();
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
      const filteredIds = filteredVoters.map(v => v.id);
      setSelectedVoterIds(selectedVoterIds.filter(id => !filteredIds.includes(id)));
    } else {
      const filteredIds = filteredVoters.map(v => v.id);
      const newSelected = Array.from(new Set([...selectedVoterIds, ...filteredIds]));
      setSelectedVoterIds(newSelected);
    }
  };

  // Restablecer votos de múltiples votantes en lote (VOTO SECRETO)
  const handleBulkResetVotes = async () => {
    const selectedVoters = voters.filter(v => selectedVoterIds.includes(v.id));
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
        showToast(`Participación de ${selectedVoters.length} miembros restablecida en Google Sheets`, "success");
        refreshData();
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
    const idsToDelete = new Set(selectedVoterIds);
    if (idsToDelete.size === 0) return;

    if (!window.confirm(`¿Estás seguro de que deseas eliminar a los ${idsToDelete.size} votantes seleccionados del censo?`)) {
      return;
    }

    setIsUploading(true);
    const updatedVoters = voters.filter(v => !idsToDelete.has(v.id));
    try {
      if (config.sheetUrlVoters) {
        if (!isConnected) {
          throw new Error("No hay conexión con Google Sheets. No se pueden eliminar los votantes.");
        }
        await updateVotersInSheets(config.sheetUrlVoters, updatedVoters);
        showToast(`${idsToDelete.size} votantes eliminados de Google Sheets`, "success");
        refreshData();
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

  // Reset de base de datos local
  const handleResetLocalData = () => {
    if (window.confirm("¿Deseas restaurar la base de datos local a los valores iniciales de prueba? Esto borrará tus cambios locales.")) {
      localStorage.removeItem('icc_local_candidates');
      localStorage.removeItem('icc_local_voters');
      localStorage.removeItem('icc_local_votes');
      localStorage.removeItem('icc_candidates_cleared_by_user');
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

  // Filtrar Votantes en lista considerando búsqueda y estado de participación (VOTO SECRETO)
  const filteredVoters = voters.filter(v => {
    const fullName = `${v.name || ''} ${v.lastName || ''}`.trim().toLowerCase();
    const matchesSearch = fullName.includes(voterSearch.toLowerCase());
    if (!matchesSearch) return false;

    // Con voto secreto: hasVoted = true/false (binario, no parcial)
    if (voterStatusFilter === 'completed') return v.hasVoted === true;
    if (voterStatusFilter === 'pending') return !v.hasVoted;
    if (voterStatusFilter === 'partial') return false; // Ya no existe estado parcial
    return true; // 'all'
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
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
            <h2 style={{ fontFamily: 'var(--font-heading)', fontWeight: 700, margin: 0 }}>Resultados de la Votación</h2>
            {votes.length > 0 && (
              <Tooltip text="Eliminar permanentemente todos los votos de la base de datos." position="left">
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
            )}
          </div>

          {/* Tarjetas de Estadísticas Globales */}
          <div className="grid-3" style={{ marginBottom: '32px' }}>
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
                {voters.filter(v => v.hasVoted).length}
                <span style={{ fontSize: '14px', fontWeight: 400, color: 'var(--text-secondary)', marginLeft: '6px' }}>
                  / {voters.length}
                </span>
              </div>
            </div>
          </div>

          <h3 style={{ fontFamily: 'var(--font-heading)', fontWeight: 600, marginBottom: '16px' }}>Votación por Candidato</h3>
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
            <div className="grid-2" style={{ marginBottom: '32px' }}>
              {candidates.map(cand => {
                const stats = getCandidateStats(cand.id);
                const fullName = `${cand.firstName || ''} ${cand.lastName || ''}`.trim();
                return (
                  <div key={cand.id} className={`card result-card ${stats.approvalRate >= 75 ? 'high-approval' : ''}`}>
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
          )}

          {/* Tabla de Votos Detallada */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
            <h3 style={{ fontFamily: 'var(--font-heading)', fontWeight: 600, margin: 0 }}>Desglose de Votos Recibidos</h3>
            <div>
              <label style={{ fontSize: '13px', marginRight: '8px', fontWeight: 600, color: 'var(--text-secondary)' }}>Filtrar candidato:</label>
              <select 
                className="form-control" 
                style={{ display: 'inline-block', width: 'auto', padding: '6px 12px', fontSize: '13px' }}
                value={resultsFilter}
                onChange={(e) => setResultsFilter(e.target.value)}
              >
                <option value="all">Todos los candidatos</option>
                {candidates.map(c => (
                  <option key={c.id} value={c.id}>{`${c.firstName || ''} ${c.lastName || ''}`.trim()}</option>
                ))}
              </select>
            </div>
          </div>

          <div className="votos-table-container">
            {displayVotes.length === 0 ? (
              <div style={{ padding: '24px', textAlign: 'center', color: 'var(--text-muted)' }}>
                No se han registrado votos todavía para este criterio.
              </div>
            ) : (
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
                  {displayVotes.map((vote, index) => (
                    <tr key={index}>
                      <td style={{ fontWeight: 600 }}>{vote.candidateName}</td>
                      <td>
                        <span className={`badge ${vote.status === 'Aprueba' ? 'badge-success' : 'badge-danger'}`}>
                          {vote.status}
                        </span>
                      </td>
                      <td style={{ color: vote.status === 'No Aprueba' ? 'var(--danger)' : 'var(--text-secondary)', fontStyle: vote.status === 'No Aprueba' ? 'italic' : 'normal' }}>
                        {vote.reason ? vote.reason : '—'}
                      </td>
                      <td style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
                        {vote.date ? new Date(vote.date).toLocaleString('es-ES') : '—'}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>

          {/* Database actions panel at the bottom */}
          <div style={{ marginTop: '40px', paddingTop: '24px', borderTop: '1px solid var(--border)', display: 'flex', gap: '12px', justifyContent: 'flex-end', flexWrap: 'wrap' }}>
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
                return (
                  <div key={cand.id} className="candidate-row-card">
                    <div className="candidate-info">
                      <CandidatePhoto
                        photo={cand.photo}
                        firstName={cand.firstName}
                        lastName={cand.lastName}
                        className="candidate-avatar"
                      />
                      <div>
                        <div className="candidate-name">{fullName}</div>
                        <div className="candidate-testimony-snippet">{cand.testimony}</div>
                      </div>
                    </div>
                    <div style={{ display: 'flex', gap: '8px' }}>
                      <Tooltip text="Modificar la información, foto o testimonio de este candidato.">
                        <button 
                          type="button"
                          className="btn btn-secondary" 
                          onClick={() => handleStartEditCandidate(cand)}
                          style={{ padding: '6px 12px', fontSize: '12px' }}
                        >
                          Editar
                        </button>
                      </Tooltip>
                      <Tooltip text="Dar de baja de forma permanente a este candidato de la votación.">
                        <button 
                          type="button"
                          className="btn btn-danger" 
                          onClick={() => handleDeleteCandidate(cand.id)}
                          style={{ padding: '6px 12px', fontSize: '12px' }}
                        >
                          Eliminar
                        </button>
                      </Tooltip>
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

            {/* Buscador y Filtros Rápidos */}
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '12px', alignItems: 'center', marginBottom: '20px' }}>
              <div style={{ flex: '1', minWidth: '200px' }}>
                <input 
                  type="text" 
                  className="form-control" 
                  placeholder="Buscar votante por nombre..."
                  value={voterSearch}
                  onChange={(e) => setVoterSearch(e.target.value)}
                  style={{ margin: 0 }}
                />
              </div>
              <div className="tabs-header" style={{ margin: 0, padding: '4px', alignSelf: 'stretch', display: 'flex', alignItems: 'center' }}>
                <Tooltip text="Mostrar a todos los miembros autorizados sin filtros." position="top">
                  <button 
                    type="button" 
                    className={`tab-btn ${voterStatusFilter === 'all' ? 'active' : ''}`}
                    onClick={() => {
                      setVoterStatusFilter('all');
                      setSelectedVoterIds([]); // Limpiar selección al cambiar de filtro
                    }}
                    style={{ padding: '6px 12px', fontSize: '12px' }}
                  >
                    Todos
                  </button>
                </Tooltip>
                <Tooltip text="Mostrar solo a los miembros que no han emitido ningún voto." position="top">
                  <button 
                    type="button" 
                    className={`tab-btn ${voterStatusFilter === 'pending' ? 'active' : ''}`}
                    onClick={() => {
                      setVoterStatusFilter('pending');
                      setSelectedVoterIds([]);
                    }}
                    style={{ padding: '6px 12px', fontSize: '12px' }}
                  >
                    Sin Votar
                  </button>
                </Tooltip>
                <Tooltip text="Mostrar a los miembros que han votado a algunos candidatos pero no a todos." position="top">
                  <button 
                    type="button" 
                    className={`tab-btn ${voterStatusFilter === 'partial' ? 'active' : ''}`}
                    onClick={() => {
                      setVoterStatusFilter('partial');
                      setSelectedVoterIds([]);
                    }}
                    style={{ padding: '6px 12px', fontSize: '12px' }}
                  >
                    Parcial
                  </button>
                </Tooltip>
                <Tooltip text="Mostrar solo a los miembros que ya han completado todos sus votos." position="top">
                  <button 
                    type="button" 
                    className={`tab-btn ${voterStatusFilter === 'completed' ? 'active' : ''}`}
                    onClick={() => {
                      setVoterStatusFilter('completed');
                      setSelectedVoterIds([]);
                    }}
                    style={{ padding: '6px 12px', fontSize: '12px' }}
                  >
                    Ya Votó
                  </button>
                </Tooltip>
              </div>
            </div>

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

            {/* Fila de Selección Maestra */}
            {filteredVoters.length > 0 && (
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', padding: '10px 16px', backgroundColor: 'var(--bg-input)', border: '1px solid var(--border)', borderBottom: 'none', borderRadius: 'var(--radius-md) var(--radius-md) 0 0', fontSize: '13px', color: 'var(--text-secondary)' }}>
                <input 
                  type="checkbox" 
                  checked={filteredVoters.length > 0 && filteredVoters.every(v => selectedVoterIds.includes(v.id))}
                  onChange={() => handleToggleSelectAll(filteredVoters.length > 0 && filteredVoters.every(v => selectedVoterIds.includes(v.id)))}
                  style={{ cursor: 'pointer', width: '16px', height: '16px' }}
                />
                <span style={{ fontWeight: '600' }}>Seleccionar todos los de esta lista ({filteredVoters.length})</span>
              </div>
            )}

            {filteredVoters.length === 0 ? (
              <p style={{ color: 'var(--text-muted)', textAlign: 'center', padding: '24px 0', border: '1px solid var(--border)', borderRadius: 'var(--radius-md)', margin: 0 }}>No se encontraron votantes.</p>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', maxHeight: '400px', overflowY: 'auto', padding: '4px', border: '1px solid var(--border)', borderTop: 'none', borderRadius: '0 0 var(--radius-md) var(--radius-md)' }}>
                {filteredVoters.map((voter, index) => {
                  const progress = getVoterProgress(voter);
                  const isChecked = selectedVoterIds.includes(voter.id);
                  return (
                    <div key={voter.id || index} className="candidate-row-card" style={{ padding: '10px 16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', margin: 0, border: 'none', borderBottom: index < filteredVoters.length - 1 ? '1px solid var(--border)' : 'none', borderRadius: 0 }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                        <input 
                          type="checkbox" 
                          checked={isChecked} 
                          onChange={(e) => {
                            if (e.target.checked) {
                              setSelectedVoterIds([...selectedVoterIds, voter.id]);
                            } else {
                              setSelectedVoterIds(selectedVoterIds.filter(id => id !== voter.id));
                            }
                          }}
                          style={{ cursor: 'pointer', width: '16px', height: '16px' }}
                        />
                        <div style={{
                          width: '32px',
                          height: '32px',
                          borderRadius: '50%',
                          backgroundColor: 'var(--primary-light)',
                          color: 'var(--primary)',
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'center',
                          fontWeight: 'bold',
                          fontSize: '12px'
                        }}>
                          {voter.name?.charAt(0)}{voter.lastName?.charAt(0)}
                        </div>
                        <div className="candidate-name" style={{ fontSize: '15px' }}>{voter.name} {voter.lastName}</div>
                      </div>
                      
                      <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                        {/* Estado del Voto */}
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

                        {/* Botón Restaurar */}
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

                        {/* Botón Eliminar */}
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
                    </div>
                  );
                })}
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
