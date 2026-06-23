import { useState, useEffect, useRef } from 'react';
import { addVoteToSheets, getCleanPhotoUrl } from '../utils/api';
import Tooltip from './Tooltip';


export default function VoterPanel({
  config,
  candidates,
  voters,
  votes,
  isConnected,
  refreshData,
  showToast,
  setVotes,
  isLoading
}) {
  const [selectedVoter, setSelectedVoter] = useState(null);
  const [voterSearchText, setVoterSearchText] = useState('');
  const [suggestions, setSuggestions] = useState([]);
  const [isFocused, setIsFocused] = useState(false);
  const [highlightedIndex, setHighlightedIndex] = useState(-1);
  const suggestionsRef = useRef(null);

  // Estados del Wizard de Votación
  const [voteStatus, setVoteStatus] = useState(null); // 'approve' | 'disapprove'
  const [disapproveReason, setDisapproveReason] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [hasFinished, setHasFinished] = useState(false);

  // Filtrar sugerencias de votantes basados en la búsqueda
  useEffect(() => {
    const timer = setTimeout(() => {
      if (voterSearchText.trim() === '') {
        setSuggestions([]);
        return;
      }

      const filtered = voters.filter(v => {
        const fullName = `${v.name || ''} ${v.lastName || ''}`.trim().toLowerCase();
        return fullName.includes(voterSearchText.toLowerCase());
      });
      setSuggestions(filtered);
      setHighlightedIndex(0);
    }, 0);

    return () => clearTimeout(timer);
  }, [voterSearchText, voters]);

  // Verificar si un votante ya ha votado por un candidato específico de forma estricta
  const hasVotedFor = (voter, candidateId) => {
    const voterFirstNameNormalized = (voter.name || '').trim().toLowerCase();
    const voterLastNameNormalized = (voter.lastName || '').trim().toLowerCase();
    
    return votes.some(v => {
      // Comparar IDs de forma robusta (tanto string como numérica)
      const cId = String(v.candidateId || v.ID_Candidato || v["ID Candidato"] || '').trim();
      const targetId = String(candidateId || '').trim();
      
      const cIdInt = parseInt(cId, 10);
      const targetIdInt = parseInt(targetId, 10);
      
      const idsMatch = cId === targetId || (!isNaN(cIdInt) && !isNaN(targetIdInt) && cIdInt === targetIdInt);
      if (!idsMatch) return false;

      // Comparar por campos de nombre/apellido divididos si están disponibles
      const vf = (v.voterFirstName || '').trim().toLowerCase();
      const vl = (v.voterLastName || '').trim().toLowerCase();
      
      if (vf && vl) {
        return vf === voterFirstNameNormalized && vl === voterLastNameNormalized;
      }
      
      // Fallback a nombre completo consolidado
      const voterFullName = `${voterFirstNameNormalized} ${voterLastNameNormalized}`.trim();
      const vName = (v.voterName || v["Nombre del Miembro"] || v.Nombre_Miembro || v.Nombre_Votante || v["Nombre Votante"] || '').trim().toLowerCase();
      return vName === voterFullName;
    });
  };

  // Obtener la lista de candidatos pendientes para el votante seleccionado
  const getPendingCandidates = () => {
    if (!selectedVoter) return [];
    return candidates.filter(cand => !hasVotedFor(selectedVoter, cand.id));
  };

  const pendingCandidates = getPendingCandidates();
  const currentCandidate = pendingCandidates[0]; // Siempre tomamos el primer candidato pendiente de la lista

  // Manejar la selección de un votante
  const handleSelectVoter = async (voter) => {
    setSelectedVoter(voter);
    const fullName = `${voter.name || ''} ${voter.lastName || ''}`.trim();
    setVoterSearchText(fullName);
    setSuggestions([]);
    setVoteStatus(null);
    setDisapproveReason('');
    setHasFinished(false);
    showToast(`Sesión iniciada como: ${fullName}`, "info");

    // Sincronizar inmediatamente al iniciar sesión de voto
    if (isConnected) {
      try {
        await refreshData();
      } catch (e) {
        console.error("Error al sincronizar al iniciar sesión:", e);
      }
    }
  };

  // Enviar Voto
  const handleConfirmVote = async () => {
    if (!selectedVoter) return;
    if (!currentCandidate) return;
    if (!voteStatus) {
      showToast("Por favor, selecciona 'Aprobar' o 'No Aprobar'", "error");
      return;
    }

    if (voteStatus === 'disapprove') {
      const reason = disapproveReason.trim();
      if (!reason) {
        showToast("Debes justificar por qué no apruebas la membresía", "error");
        return;
      }

      // Validar respuestas evasivas o demasiado simples ("Porque no", "No", etc.)
      const clean = reason.toLowerCase().replace(/[.,/#!$%^&*;:{}=\-_`~()?]/g, "").replace(/\s+/g, " ").trim();
      const blacklist = [
        "no", "porque no", "por que no", "porque si", "porque sí", "por que si", "por que sí",
        "nada", "ninguna", "ninguno", "ningun", "ningún", "no se", "no sé", "nose", "ok", "okay",
        "asdf", "asdfg", "asdfgh", "qwerty", "no quiero", "no lo apruebo", "no apruebo", "no lo se", "no lo sé",
        "prueba", "test", "sin justificacion", "sin justificación", "sin comentarios"
      ];

      const isBlacklisted = blacklist.includes(clean);
      const isTooShort = clean.length < 10;
      const hasFewWords = clean.split(' ').filter(w => w.length > 1).length < 2;

      if (isBlacklisted || isTooShort || hasFewWords) {
        showToast("Por favor, proporciona una breve explicación válida de tu objeción (evita respuestas genéricas como 'No' o 'Porque no').", "error");
        return;
      }
    }

    // Evitar doble votación (seguridad extra en el cliente)
    if (hasVotedFor(selectedVoter, currentCandidate.id)) {
      showToast("Ya has votado por este candidato.", "error");
      setVoteStatus(null);
      setDisapproveReason('');
      if (pendingCandidates.length <= 1) {
        setHasFinished(true);
      }
      return;
    }

    setIsSubmitting(true);
    const voteData = {
      candidateId: currentCandidate.id,
      candidateFirstName: currentCandidate.firstName,
      candidateLastName: currentCandidate.lastName,
      voterFirstName: selectedVoter.name,
      voterLastName: selectedVoter.lastName,
      status: voteStatus === 'approve' ? 'Aprueba' : 'No Aprueba',
      reason: voteStatus === 'disapprove' ? disapproveReason.trim() : ''
    };

    try {
      if (config.sheetUrlCandidates) {
        if (!isConnected) {
          throw new Error("No hay conexión con Google Sheets. Por favor, verifica tu conexión a internet.");
        }
        // Enviar a Google Sheets
        await addVoteToSheets(config.sheetUrlCandidates, voteData);
        showToast("Voto registrado en Google Sheets", "success");
        // Sincronizar datos globales esperando a que termine
        await refreshData();
      } else {
        // Guardar localmente
        const voteWithTime = {
          candidateId: voteData.candidateId,
          candidateName: `${voteData.candidateFirstName} ${voteData.candidateLastName}`.trim(),
          voterName: `${voteData.voterFirstName} ${voteData.voterLastName}`.trim(),
          status: voteData.status,
          reason: voteData.reason,
          timestamp: new Date().toISOString()
        };
        const localVotes = localStorage.getItem('icc_local_votes');
        const currentVotes = localVotes ? JSON.parse(localVotes) : [];
        const updatedVotes = [...currentVotes, voteWithTime];
        
        setVotes(updatedVotes);
        localStorage.setItem('icc_local_votes', JSON.stringify(updatedVotes));
        showToast("Voto registrado localmente", "success");
      }

      // Reiniciar estado del voto
      setVoteStatus(null);
      setDisapproveReason('');
 
      // Avanzar en el wizard (si este era el último candidato pendiente, finaliza el flujo)
      if (pendingCandidates.length <= 1) {
        setHasFinished(true);
      }

    } catch (error) {
      console.error(error);
      showToast(`Error al enviar voto: ${error.message}`, "error");
    } finally {
      setIsSubmitting(false);
    }
  };

  // Reiniciar sesión/pantalla de votación
  const handleResetSession = () => {
    setSelectedVoter(null);
    setVoterSearchText('');
    setSuggestions([]);
    setVoteStatus(null);
    setDisapproveReason('');
    setHasFinished(false);
  };

  // Soporte de navegación por teclado en el buscador
  const handleKeyDown = (e) => {
    if (suggestions.length === 0) return;

    if (e.key === 'ArrowDown') {
      e.preventDefault();
      setHighlightedIndex(prev => (prev + 1) % suggestions.length);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      setHighlightedIndex(prev => (prev - 1 + suggestions.length) % suggestions.length);
    } else if (e.key === 'Enter') {
      e.preventDefault();
      if (highlightedIndex >= 0 && highlightedIndex < suggestions.length) {
        handleSelectVoter(suggestions[highlightedIndex]);
      }
    } else if (e.key === 'Escape') {
      setSuggestions([]);
    }
  };

  return (
    <div className="voter-container">
      
      {/* PASO 1: SELECCIONAR VOTANTE */}
      {!selectedVoter && (
        <div className="card voter-card" style={{ padding: '32px 24px', textAlign: 'center' }}>
          <div style={{ fontSize: '48px', marginBottom: '16px' }}>📖</div>
          <h2 style={{ fontFamily: 'var(--font-heading)', fontWeight: 800, marginBottom: '12px' }}>
            Identificación de Votante
          </h2>
          <p style={{ color: 'var(--text-secondary)', fontSize: '15px', marginBottom: '24px' }}>
            Por favor, busca y selecciona tu nombre en la lista oficial de miembros autorizados para votar.
          </p>

          <div className="form-group search-voter-wrapper">
            <label className="form-label" style={{ textAlign: 'center' }}>Escribe tu nombre</label>
            <input 
              type="text" 
              className="form-control" 
              placeholder="Buscar mi nombre..."
              value={voterSearchText}
              onChange={(e) => setVoterSearchText(e.target.value)}
              onFocus={() => setIsFocused(true)}
              onBlur={() => setTimeout(() => setIsFocused(false), 200)}
              onKeyDown={handleKeyDown}
              style={{ textAlign: 'center', fontSize: '18px', padding: '14px' }}
              autoFocus
            />

            {isFocused && suggestions.length > 0 && (
              <div className="voters-suggestions-list" ref={suggestionsRef}>
                {suggestions.map((v, index) => (
                  <div 
                    key={v.id || index} 
                    className={`voter-suggestion-item ${index === highlightedIndex ? 'highlighted' : ''}`}
                    onMouseDown={() => handleSelectVoter(v)}
                  >
                    {v.name} {v.lastName}
                  </div>
                ))}
              </div>
            )}

            {isFocused && voterSearchText.trim() !== '' && suggestions.length === 0 && (
              <div className="voters-suggestions-list">
                <div className="no-suggestions">No se encontró ningún miembro con ese nombre</div>
              </div>
            )}
          </div>

          <div style={{ marginTop: '32px', borderTop: '1px solid var(--border)', paddingTop: '16px', fontSize: '13px', color: 'var(--text-muted)' }}>
            ¿No apareces en la lista? Comunícate con el panel administrativo para registrarte.
          </div>
        </div>
      )}

      {/* PASO 2: VOTACIÓN ACTIVA */}
      {selectedVoter && !hasFinished && pendingCandidates.length > 0 && (
        <div className="card voter-card" style={{ padding: '24px' }}>
          
          {/* Cabecera Wizard */}
          <div className="wizard-header">
            <div>
              <span className="step-indicator">Candidato {candidates.length - pendingCandidates.length + 1} de {candidates.length}</span>
              <h3 style={{ fontFamily: 'var(--font-heading)', fontWeight: 700, margin: '2px 0 0' }}>Votación de Membresía</h3>
            </div>
          </div>

          {/* Perfil del Candidato */}
          <div className="candidate-vote-profile">
            <div className="candidate-vote-avatar-wrapper">
              {currentCandidate.photo ? (
                <img 
                  src={getCleanPhotoUrl(currentCandidate.photo)} 
                  className="candidate-vote-avatar" 
                  alt={`${currentCandidate.firstName || ''} ${currentCandidate.lastName || ''}`} 
                />
              ) : (
                <div 
                  className="candidate-vote-avatar placeholder-avatar" 
                  style={{ 
                    display: 'flex', 
                    alignItems: 'center', 
                    justifyContent: 'center', 
                    backgroundColor: 'var(--primary-light)', 
                    color: 'var(--primary)', 
                    fontWeight: 'bold', 
                    fontSize: '40px',
                    borderRadius: '50%',
                    boxShadow: 'var(--shadow-md)',
                    border: '4px solid var(--primary-light)',
                    width: '140px',
                    height: '140px',
                    margin: '0 auto'
                  }}
                >
                  {currentCandidate.firstName?.charAt(0)}{currentCandidate.lastName?.charAt(0)}
                </div>
              )}
            </div>
            <div className="candidate-vote-name">{currentCandidate.firstName} {currentCandidate.lastName}</div>
            <div className="candidate-vote-tag">Postulante a Miembro</div>
          </div>

          {/* Testimonio */}
          <div className="testimony-container">
            <div className="testimony-title">
              <span>✍</span> Testimonio de Conversión
            </div>
            <div className="testimony-content">
              "{currentCandidate.testimony}"
            </div>
          </div>

          {/* Botones de Aprobación */}
          <div className="vote-actions-wrapper">
            <Tooltip text="Aprobar la solicitud de membresía de este candidato.">
              <button 
                type="button"
                className={`btn btn-vote btn-vote-approve ${voteStatus === 'approve' ? 'selected' : ''}`}
                onClick={() => {
                  setVoteStatus('approve');
                  setDisapproveReason('');
                }}
              >
                <span>✓</span> Apruebo
              </button>
            </Tooltip>
            <Tooltip text="Objetar la solicitud de membresía indicando una justificación válida.">
              <button 
                type="button"
                className={`btn btn-vote btn-vote-disapprove ${voteStatus === 'disapprove' ? 'selected' : ''}`}
                onClick={() => setVoteStatus('disapprove')}
              >
                <span>✗</span> No Apruebo
              </button>
            </Tooltip>
          </div>

          {/* Justificación obligatoria para desaprobación */}
          <div className={`reason-slide-container ${voteStatus === 'disapprove' ? 'visible' : ''}`}>
            <div className="reason-title">Justificación de la Objeción (Requerido)</div>
            <textarea 
              className="form-control" 
              rows="3" 
              placeholder="Por favor explica el por qué no apruebas la membresía de este candidato. Esta información será revisada con confidencialidad por el pastorado."
              value={disapproveReason}
              onChange={(e) => setDisapproveReason(e.target.value)}
              style={{ fontSize: '14px' }}
              required={voteStatus === 'disapprove'}
            ></textarea>
          </div>

          {/* Confirmar Acción */}
          <div className="confirm-action-wrapper">
            <Tooltip text="Registrar tu decisión de forma definitiva y confidencial en el sistema.">
              <button 
                className="btn btn-primary" 
                style={{ width: '100%', height: '48px', fontSize: '16px' }}
                onClick={handleConfirmVote}
                disabled={isSubmitting || !voteStatus}
              >
                {isSubmitting ? "Registrando voto..." : "Confirmar y Enviar Voto"}
              </button>
            </Tooltip>
          </div>
        </div>
      )}

      {/* CARGANDO DATOS (Para evitar saltos a la pantalla de agradecimiento) */}
      {selectedVoter && isLoading && pendingCandidates.length === 0 && (
        <div className="card voter-card" style={{ padding: '48px 24px', textAlign: 'center' }}>
          <div className="spinner" style={{ display: 'inline-block', width: '40px', height: '40px', border: '4px solid var(--border)', borderTopColor: 'var(--primary)', borderRadius: '50%', animation: 'spin 1s linear infinite', marginBottom: '12px' }}></div>
          <div style={{ marginTop: '12px', color: 'var(--text-secondary)' }}>Actualizando información de votación...</div>
          <style>{`
            @keyframes spin {
              0% { transform: rotate(0deg); }
              100% { transform: rotate(360deg); }
            }
          `}</style>
        </div>
      )}

      {/* CASO 3: YA VOTÓ POR TODOS (DENTRO DEL WIZARD) O PANTALLA AGRADECIMIENTO */}
      {selectedVoter && !isLoading && (hasFinished || pendingCandidates.length === 0) && (
        <div className="card voter-card thank-you-view">
          <div className="thank-you-icon">✓</div>
          <h2 className="thank-you-title">¡Votos Completados!</h2>
          <p className="thank-you-text">
            Hermano(a) <strong>{`${selectedVoter.name || ''} ${selectedVoter.lastName || ''}`.trim()}</strong>, agradecemos tu participación en el proceso de votación. Tus respuestas han sido registradas confidencialmente en la base de datos de la iglesia.
          </p>
          
          <Tooltip text="Cerrar tu sesión de votación y regresar a la pantalla de inicio.">
            <button 
              className="btn btn-primary" 
              onClick={handleResetSession}
              style={{ padding: '12px 24px', fontWeight: 700, borderRadius: 'var(--radius-md)' }}
            >
              Finalizar Sesión
            </button>
          </Tooltip>

          <div className="thank-you-verse">
            "Y considerémonos unos a otros para estimularnos al amor y a las buenas obras."<br />
            <strong>Hebreos 10:24</strong>
          </div>
        </div>
      )}

    </div>
  );
}
