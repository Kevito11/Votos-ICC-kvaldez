/**
 * Utilidades de API para conectar con Google Sheets (vía Google Apps Script Web App)
 * y con Supabase Storage (para subir fotos).
 */

/**
 * Obtener todos los datos (candidatos, votantes y votos) desde Google Sheets
 * @param {string} sheetUrl - URL de la Web App de Google Apps Script
 */
export async function fetchVotersFromSheets(sheetUrl) {
  if (!sheetUrl) throw new Error("URL de Google Sheets para Votantes no configurada");
  
  const url = `${sheetUrl}${sheetUrl.includes('?') ? '&' : '?'}action=getData&_t=${Date.now()}`;
  
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Error al leer votantes de Google Sheets: ${response.statusText}`);
  }
  
  const data = await response.json();
  return data;
}

export async function fetchCandidatesAndVotesFromSheets(sheetUrl) {
  if (!sheetUrl) throw new Error("URL de Google Sheets para Candidatos no configurada");
  
  const url = `${sheetUrl}${sheetUrl.includes('?') ? '&' : '?'}action=getData&_t=${Date.now()}`;
  
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Error al leer candidatos de Google Sheets: ${response.statusText}`);
  }
  
  const data = await response.json();
  return data;
}

/**
 * Enviar una petición POST general a Google Apps Script usando content-type text/plain
 * para evitar bloqueos de CORS preflight (OPTIONS).
 * @param {string} sheetUrl 
 * @param {object} payload 
 */
async function postToSheets(sheetUrl, payload) {
  if (!sheetUrl) throw new Error("URL de Google Sheets no configurada");
  
  const response = await fetch(sheetUrl, {
    method: "POST",
    headers: {
      "Content-Type": "text/plain", // Truco de CORS para Google Apps Script
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    throw new Error(`Error en servidor de Google Sheets: ${response.statusText}`);
  }

  const result = await response.json();
  if (result.status === "error") {
    throw new Error(result.msg || "Error devuelto por Google Sheets");
  }
  return result;
}

/**
 * Registrar un voto en Google Sheets
 */
export async function addVoteToSheets(sheetUrl, { candidateId, candidateFirstName, candidateLastName, voterFirstName, voterLastName, status, reason }) {
  return postToSheets(sheetUrl, {
    action: "addVote",
    candidateId,
    candidateFirstName,
    candidateLastName,
    voterFirstName,
    voterLastName,
    status,
    reason,
  });
}

/**
 * Agregar un candidato a Google Sheets
 */
export async function addCandidateToSheets(sheetUrl, { id, firstName, lastName, testimony, photoBase64, photoName }) {
  return postToSheets(sheetUrl, {
    action: "addCandidate",
    id,
    firstName,
    lastName,
    testimony,
    photoBase64,
    photoName,
  });
}

/**
 * Eliminar un candidato de Google Sheets
 */
export async function deleteCandidateFromSheets(sheetUrl, id) {
  return postToSheets(sheetUrl, {
    action: "deleteCandidate",
    id,
  });
}

/**
 * Vaciar la lista de candidatos en Google Sheets
 */
export async function clearCandidatesInSheets(sheetUrl) {
  return postToSheets(sheetUrl, {
    action: "clearCandidates",
  });
}

/**
 * Actualizar un candidato existente en Google Sheets
 */
export async function updateCandidateInSheets(sheetUrl, { id, firstName, lastName, testimony, photoBase64, photoName }) {
  return postToSheets(sheetUrl, {
    action: "updateCandidate",
    id,
    firstName,
    lastName,
    testimony,
    photoBase64,
    photoName,
  });
}


/**
 * Actualizar la lista de votantes en Google Sheets
 */
export async function updateVotersInSheets(sheetUrl, votersList) {
  return postToSheets(sheetUrl, {
    action: "updateVoters",
    voters: votersList,
  });
}

/**
 * Sube una imagen a Supabase Storage y retorna la URL pública del objeto.
 * @param {File} file - El archivo de imagen seleccionado por el usuario
 * @param {string} supabaseUrl - URL de Supabase (ej. https://xxxxxx.supabase.co)
 * @param {string} supabaseKey - Anon/Public key de Supabase
 * @param {string} bucketName - Nombre del bucket de almacenamiento (por defecto 'candidatos')
 */
export async function uploadPhotoToSupabase(file, supabaseUrl, supabaseKey, bucketName = "candidatos") {
  if (!supabaseUrl || !supabaseKey) {
    throw new Error("Configuración de Supabase incompleta");
  }

  const cleanUrl = supabaseUrl.replace(/\/$/, "");
  const fileExt = file.name.split(".").pop();
  // Crear un nombre de archivo único para evitar colisiones
  const fileName = `${Date.now()}-${Math.random().toString(36).substring(2, 8)}.${fileExt}`;
  
  const uploadUrl = `${cleanUrl}/storage/v1/object/${bucketName}/${fileName}`;

  const response = await fetch(uploadUrl, {
    method: "POST",
    headers: {
      "apikey": supabaseKey,
      "Authorization": `Bearer ${supabaseKey}`,
      "Content-Type": file.type,
    },
    body: file,
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    throw new Error(errorData.message || `Error al subir imagen: ${response.statusText}`);
  }

  // URL pública del archivo subido en Supabase
  const publicUrl = `${cleanUrl}/storage/v1/object/public/${bucketName}/${fileName}`;
  return publicUrl;
}

/**
 * Limpia y normaliza la URL de una foto (especialmente de Google Drive) para que se renderice correctamente en <img>
 */
export function getCleanPhotoUrl(photoUrl) {
  if (!photoUrl) return "";
  if (photoUrl.startsWith("data:")) return photoUrl;
  
  try {
    let fileId = "";
    if (photoUrl.includes("id=")) {
      // Formato: https://docs.google.com/uc?export=view&id=FILE_ID
      fileId = photoUrl.split("id=")[1].split("&")[0];
    } else if (photoUrl.includes("/d/")) {
      // Formato: https://drive.google.com/file/d/FILE_ID/view
      fileId = photoUrl.split("/d/")[1].split("/")[0];
    } else if (photoUrl.includes("googleusercontent.com/d/")) {
      // Formato: https://lh3.googleusercontent.com/d/FILE_ID
      fileId = photoUrl.split("googleusercontent.com/d/")[1];
    }
    
    if (fileId) {
      return `https://lh3.googleusercontent.com/d/${fileId}`;
    }
  } catch (e) {
    console.error("Error al limpiar URL de foto:", e);
  }
  
  return photoUrl;
}

/**
 * Restaurar los votos de un votante específico en Google Sheets (permitiéndole volver a votar)
 * Reutiliza la acción masiva 'resetMultipleVotersVotes' para asegurar compatibilidad sin requerir redespliegue de Apps Script.
 */
export async function resetVoterVotesInSheets(sheetUrl, { voterFirstName, voterLastName }) {
  return postToSheets(sheetUrl, {
    action: "resetMultipleVotersVotes",
    voters: [{ firstName: voterFirstName, lastName: voterLastName }]
  });
}

/**
 * Vaciar la lista de todos los votos en Google Sheets
 */
export async function clearVotesInSheets(sheetUrl) {
  return postToSheets(sheetUrl, {
    action: "clearVotes",
  });
}

/**
 * Restaurar los votos de múltiples votantes en Google Sheets
 */
export async function resetMultipleVotersVotesInSheets(sheetUrl, votersList) {
  return postToSheets(sheetUrl, {
    action: "resetMultipleVotersVotes",
    voters: votersList // Lista de objetos { firstName, lastName }
  });
}
