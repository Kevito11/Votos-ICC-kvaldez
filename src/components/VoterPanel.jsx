import { useState, useEffect, useRef } from 'react';
import { addVoteToSheets, markVoterParticipation } from '../utils/api';
import CandidatePhoto from './CandidatePhoto';
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
  const [votedCandidateIds, setVotedCandidateIds] = useState([]);

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

  // Verificar si un votante ya ha participado en la votación (VOTO SECRETO)
  const hasVotedFor = (voter, candidateId) => {
    if (!voter) return false;
    const pastVotedIds = votes
      .filter(v => String(v.voterId).trim() === String(voter.id).trim())
      .map(v => String(v.candidateId).trim());
    const sessionVotedIds = votedCandidateIds.map(id => String(id).trim());
    return pastVotedIds.includes(String(candidateId).trim()) || sessionVotedIds.includes(String(candidateId).trim());
  };

  // Obtener etiqueta de estado del votante para mostrar en las sugerencias (sin revelar su decisión)
  const getVoterStatusLabel = (voter) => {
    if (!voter) return { text: "Pendiente", style: {} };
    
    const pastVotedIds = votes
      .filter(v => String(v.voterId).trim() === String(voter.id).trim())
      .map(v => String(v.candidateId).trim());
    
    // Si completó todos los candidatos
    if (voter.hasVoted && pastVotedIds.length >= candidates.length) {
      return { 
        text: "Ya Votó", 
        style: { backgroundColor: 'rgba(16, 185, 129, 0.1)', color: '#10b981', border: '1px solid rgba(16, 185, 129, 0.2)' } 
      };
    }
    
    // Si tiene votos parciales
    if (pastVotedIds.length > 0 && pastVotedIds.length < candidates.length) {
      return { 
        text: `Parcial (${pastVotedIds.length}/${candidates.length})`, 
        style: { backgroundColor: 'rgba(59, 130, 246, 0.1)', color: '#3b82f6', border: '1px solid rgba(59, 130, 246, 0.2)' } 
      };
    }
    
    // Caso de marca de participación sin votos individuales
    if (voter.hasVoted && pastVotedIds.length === 0) {
      return { 
        text: "Ya Votó", 
        style: { backgroundColor: 'rgba(16, 185, 129, 0.1)', color: '#10b981', border: '1px solid rgba(16, 185, 129, 0.2)' } 
      };
    }
    
    // Si no ha votado por ninguno
    return { 
      text: "Pendiente", 
      style: { backgroundColor: 'rgba(107, 114, 128, 0.1)', color: '#6b7280', border: '1px solid rgba(107, 114, 128, 0.2)' } 
    };
  };

  // Obtener la lista de candidatos pendientes para el votante seleccionado
  const getPendingCandidates = () => {
    if (!selectedVoter) return [];
    
    // Obtener los votos ya registrados de este votante en la base de datos
    const pastVotedIds = votes
      .filter(v => String(v.voterId).trim() === String(selectedVoter.id).trim())
      .map(v => String(v.candidateId).trim());
      
    // Si ya participó y tiene todos sus votos registrados en Sheets/Local, o no le quedan candidatos por votar
    if (selectedVoter.hasVoted && pastVotedIds.length >= candidates.length) {
      return [];
    }
    
    // Si ya participó (hasVoted es true) pero no tiene ningún voto registrado en la hoja de votos,
    // por seguridad lo bloqueamos para evitar que vote de nuevo desde cero si ya participó.
    if (selectedVoter.hasVoted && pastVotedIds.length === 0) {
      return [];
    }
    
    // Combinar los votos registrados previamente en el servidor/local con los de la sesión en memoria
    const sessionVotedIds = votedCandidateIds.map(id => String(id).trim());
    const allVotedIds = [...new Set([...pastVotedIds, ...sessionVotedIds])];
    
    return candidates.filter(cand => !allVotedIds.includes(String(cand.id).trim()));
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
    setVotedCandidateIds([]); // Limpiar la lista de votos de la sesión actual
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
      voterId: selectedVoter.id, // ID opaco — no expone el nombre del votante
      status: voteStatus === 'approve' ? 'Aprueba' : 'No Aprueba',
      reason: voteStatus === 'disapprove' ? disapproveReason.trim() : ''
    };

    // ¿Es el primer voto de esta sesión? (para marcar participación después del primer voto)
    const pastVotedIds = votes
      .filter(v => String(v.voterId).trim() === String(selectedVoter.id).trim())
      .map(v => String(v.candidateId).trim());
    const isFirstVote = pastVotedIds.length === 0 && votedCandidateIds.length === 0;

    try {
      if (config.sheetUrlCandidates) {
        if (!isConnected) {
          throw new Error("No hay conexión con Google Sheets. Por favor, verifica tu conexión a internet.");
        }
        // Enviar voto secreto a Google Sheets
        await addVoteToSheets(config.sheetUrlCandidates, voteData);

        // Marcar al votante como participante en el primer voto para prevenir doble votación
        if (isFirstVote) {
          try {
            await markVoterParticipation(config.sheetUrlCandidates, { voterId: selectedVoter.id });
            // Actualizar el estado local del votante inmediatamente
            setVotes(prev => prev); // trigger re-read
          } catch (markErr) {
            console.warn("No se pudo marcar participación:", markErr);
          }
        }

        showToast("Voto registrado", "success");
        
        // Registrar en la lista de candidatos votados de la sesión
        setVotedCandidateIds(prev => [...prev, currentCandidate.id]);

        // Sincronizar datos globales esperando a que termine
        await refreshData();
      } else {
        // Guardar localmente
        const voteWithTime = {
          candidateId: voteData.candidateId,
          candidateName: `${voteData.candidateFirstName} ${voteData.candidateLastName}`.trim(),
          voterId: voteData.voterId,
          status: voteData.status,
          reason: voteData.reason,
          timestamp: new Date().toISOString()
        };
        const localVotes = localStorage.getItem('icc_local_votes');
        const currentVotes = localVotes ? JSON.parse(localVotes) : [];
        const updatedVotes = [...currentVotes, voteWithTime];
        
        setVotes(updatedVotes);
        localStorage.setItem('icc_local_votes', JSON.stringify(updatedVotes));

        // Marcar participación localmente en el primer voto
        if (isFirstVote) {
          const localVoters = localStorage.getItem('icc_local_voters');
          const currentVoters = localVoters ? JSON.parse(localVoters) : [];
          const updatedVoters = currentVoters.map(v => 
            v.id === selectedVoter.id ? { ...v, hasVoted: true, votedAt: new Date().toISOString() } : v
          );
          localStorage.setItem('icc_local_voters', JSON.stringify(updatedVoters));
          // OJO: No alteramos el estado local selectedVoter.hasVoted inmediatamente para que no rompa el flujo de candidatos pendientes de esta sesión.
        }

        showToast("Voto registrado localmente", "success");
        
        // Registrar en la lista de candidatos votados de la sesión
        setVotedCandidateIds(prev => [...prev, currentCandidate.id]);
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
    setVotedCandidateIds([]); // Limpiar la lista de votos de la sesión al finalizar
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
                {suggestions.map((v, index) => {
                  const statusInfo = getVoterStatusLabel(v);
                  return (
                    <div 
                      key={v.id || index} 
                      className={`voter-suggestion-item ${index === highlightedIndex ? 'highlighted' : ''}`}
                      onMouseDown={() => handleSelectVoter(v)}
                      style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 14px', cursor: 'pointer' }}
                    >
                      <span style={{ fontWeight: '500' }}>{v.name} {v.lastName}</span>
                      <span style={{ 
                        fontSize: '11px', 
                        fontWeight: '700', 
                        padding: '3px 8px', 
                        borderRadius: '12px',
                        textTransform: 'uppercase',
                        letterSpacing: '0.5px',
                        ...statusInfo.style 
                      }}>
                        {statusInfo.text}
                      </span>
                    </div>
                  );
                })}
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
      {selectedVoter && !hasFinished && pendingCandidates.length > 0 && !isSubmitting && !isLoading && (
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
              <CandidatePhoto
                photo={currentCandidate.photo}
                firstName={currentCandidate.firstName}
                lastName={currentCandidate.lastName}
                className="candidate-vote-avatar"
                style={{ width: '140px', height: '140px', boxShadow: 'var(--shadow-md)', border: '4px solid var(--primary-light)' }}
              />
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

      {/* TRANSICIÓN DE VOTACIÓN / PROCESANDO VOTO */}
      {selectedVoter && !hasFinished && pendingCandidates.length > 0 && (isSubmitting || isLoading) && (
        <div className="card voter-card" style={{ padding: '48px 24px', textAlign: 'center', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', minHeight: '300px' }}>
          <div className="spinner" style={{ display: 'inline-block', width: '50px', height: '50px', border: '4px solid var(--border)', borderTopColor: 'var(--primary)', borderRadius: '50%', animation: 'spin 1s linear infinite', marginBottom: '20px' }}></div>
          <h3 style={{ fontFamily: 'var(--font-heading)', fontWeight: 700, marginBottom: '8px' }}>Procesando tu decisión</h3>
          <p style={{ color: 'var(--text-secondary)', fontSize: '14px', maxWidth: '320px', margin: '0 auto' }}>
            Registrando tu voto de forma confidencial y preparando el siguiente candidato...
          </p>
          <style>{`
            @keyframes spin {
              0% { transform: rotate(0deg); }
              100% { transform: rotate(360deg); }
            }
          `}</style>
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
