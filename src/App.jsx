import { useState, useEffect, useCallback, useRef } from 'react';
import AdminPanel from './components/AdminPanel';
import VoterPanel from './components/VoterPanel';
import Tooltip from './components/Tooltip';
import { fetchVotersFromSheets, fetchCandidatesAndVotesFromSheets, addCandidateToSheets } from './utils/api';
import { MOCK_CANDIDATES, MOCK_VOTERS, MOCK_VOTES } from './utils/mockData';
import './App.css';


function App() {
  // Configuración de conexiones (persistencia local)
  const [config, setConfig] = useState(() => {
    const saved = localStorage.getItem('icc_voting_config');    const defaultConfig = {
      sheetUrl: 'https://script.google.com/macros/s/AKfycbzh8dvSWSPw7UqhJCU0-xsUs_aFZwAN2ytzVUW_19wwHKLUdc6BEFefnRckVmfoDI-aOA/exec',
      sheetUrlVoters: 'https://script.google.com/macros/s/AKfycbzh8dvSWSPw7UqhJCU0-xsUs_aFZwAN2ytzVUW_19wwHKLUdc6BEFefnRckVmfoDI-aOA/exec',
      sheetUrlCandidates: 'https://script.google.com/macros/s/AKfycbzh8dvSWSPw7UqhJCU0-xsUs_aFZwAN2ytzVUW_19wwHKLUdc6BEFefnRckVmfoDI-aOA/exec',
      supabaseUrl: '',
      supabaseKey: '',
      supabaseBucket: 'candidatos'
    };
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        let updated = false;

        // Migración a estructura unificada
        if (parsed.sheetUrl) {
          parsed.sheetUrlVoters = parsed.sheetUrl;
          parsed.sheetUrlCandidates = parsed.sheetUrl;
          updated = true;
        } else if (parsed.sheetUrlVoters) {
          parsed.sheetUrl = parsed.sheetUrlVoters;
          parsed.sheetUrlCandidates = parsed.sheetUrlVoters;
          updated = true;
        }

        // Si el usuario tiene el URL por defecto antiguo, forzar actualización al nuevo URL proporcionado
        const oldDefaults = [
          'https://script.google.com/macros/s/AKfycbxNyYPjtozsB3DoupMxCEQFluVctCP2moyc90Xn8RQPkmnCKZtiWGj8sCRee1CYBfDcew/exec',
          'https://script.google.com/macros/s/AKfycbyIxYEnN8i38ukHeBVlZ1iNyLxChAvHJUmoLO_rtG72hJSySUKm-5WiJLZJ3Zfkv7PgWQ/exec',
          'https://script.google.com/macros/s/AKfycbxec44t3dSFfF0cuOvuWwetAWmkCJTJrSCHoZVJ0PKbPBg_JXXdO4ZCrRUj-UFzXxfug/exec',
          'https://script.google.com/macros/s/AKfycbzYec44t3dSFfF0cuOvuWwetAWmkCJTJrSCHoZVJ0PKbPBg_JXXdO4ZCrRUj-UFzXxfug/exec'
        ];
        if (oldDefaults.includes(parsed.sheetUrl) || oldDefaults.includes(parsed.sheetUrlVoters)) {
          parsed.sheetUrl = defaultConfig.sheetUrl;
          parsed.sheetUrlVoters = defaultConfig.sheetUrlVoters;
          parsed.sheetUrlCandidates = defaultConfig.sheetUrlCandidates;
          updated = true;
        }

        if (!parsed.sheetUrl) {
          parsed.sheetUrl = defaultConfig.sheetUrl;
          parsed.sheetUrlVoters = defaultConfig.sheetUrlVoters;
          parsed.sheetUrlCandidates = defaultConfig.sheetUrlCandidates;
          updated = true;
        }

        if (updated) {
          localStorage.setItem('icc_voting_config', JSON.stringify(parsed));
        }
        return parsed;
      } catch (e) {
        console.error("Error al parsear config de localStorage:", e);
      }
    }
    return defaultConfig;
  });

  // Datos de la aplicación
  const [candidates, setCandidates] = useState([]);
  const [voters, setVoters] = useState([]);
  const [votes, setVotes] = useState([]);

  // Estados de control
  const [isLoading, setIsLoading] = useState(false);
  const [isConnected, setIsConnected] = useState(false);
  const [toasts, setToasts] = useState([]);
  const hasLoadedFromSheetsRef = useRef(false);
  
  // Enrutamiento simple basado en URL o estado local
  const [view, setView] = useState(() => {
    const params = new URLSearchParams(window.location.search);
    return params.get('view') || 'landing';
  });

  // Escuchar cambios de historial en el navegador
  useEffect(() => {
    const handlePopState = () => {
      const params = new URLSearchParams(window.location.search);
      setView(params.get('view') || 'landing');
    };
    window.addEventListener('popstate', handlePopState);
    return () => window.removeEventListener('popstate', handlePopState);
  }, []);

  // Función para navegar actualizando la URL
  const navigate = (newView) => {
    setView(newView);
    const url = newView === 'landing' ? window.location.pathname : `${window.location.pathname}?view=${newView}`;
    window.history.pushState({ view: newView }, '', url);
  };

  // Guardar configuración y actualizar conexión
  const saveConfig = (newConfig) => {
    setConfig(newConfig);
    localStorage.setItem('icc_voting_config', JSON.stringify(newConfig));
    showToast("Configuración guardada correctamente", "success");
  };

  // Toast System Helper
  const showToast = (message, type = 'info') => {
    const id = Date.now();
    setToasts((prev) => [...prev, { id, message, type }]);
    setTimeout(() => {
      setToasts((prev) => prev.filter(t => t.id !== id));
    }, 4000);
  };

  // Cargar datos
  const loadData = useCallback(async (activeConfig = config, showNotification = false) => {
    setIsLoading(true);
    try {
      const isSheetsMode = !!(activeConfig.sheetUrlVoters && activeConfig.sheetUrlCandidates);

      if (!isSheetsMode) {
        // Cargar candidatos locales solo si estamos en Modo Local
        let localCand = localStorage.getItem('icc_local_candidates');
        let parsedCand = null;
        if (localCand) {
          try {
            parsedCand = JSON.parse(localCand);
            if (parsedCand.length <= 2 || parsedCand.length === 7 || (parsedCand.length > 0 && parsedCand[0].name)) {
              localStorage.setItem('icc_local_candidates', JSON.stringify(MOCK_CANDIDATES));
              parsedCand = MOCK_CANDIDATES;
            }
          } catch (e) {
            console.error(e);
          }
        }
        if (!parsedCand) {
          localStorage.setItem('icc_local_candidates', JSON.stringify(MOCK_CANDIDATES));
          parsedCand = MOCK_CANDIDATES;
        }
        setCandidates(parsedCand);
      }

      if (isSheetsMode) {
        try {
          let votersData = null;
          let votesData = null;

          // Si son la misma URL, hacer solo una petición
          if (activeConfig.sheetUrlVoters === activeConfig.sheetUrlCandidates) {
            const combinedData = await fetchVotersFromSheets(activeConfig.sheetUrlVoters);
            votersData = combinedData;
            votesData = combinedData;
          } else {
            // Fallback por si son distintas
            const [v1, v2] = await Promise.all([
              fetchVotersFromSheets(activeConfig.sheetUrlVoters),
              fetchCandidatesAndVotesFromSheets(activeConfig.sheetUrlCandidates)
            ]);
            votersData = v1;
            votesData = v2;
          }

          // Mapear votantes a objeto { id, name, lastName }
          const mappedVoters = (votersData.voters || []).map((v, index) => {
            if (typeof v === 'object' && v !== null) {
              return {
                id: v.id || v.ID || v.Id || `voter-${index + 1}`,
                name: v.name || v.Nombre || v.nombre || '',
                lastName: v.lastName || v.Apellido || v.apellido || ''
              };
            }
            // Fallback por si la hoja aún tiene formato antiguo
            const parts = String(v).trim().split(' ');
            return {
              id: `voter-${index + 1}`,
              name: parts[0] || '',
              lastName: parts.slice(1).join(' ') || ''
            };
          });

          // Mapear votos desde la única hoja
          const mappedVotes = (votesData.votes || []).map(v => {
            const candidateFirstName = v.candidateFirstName || v["Nombre Candidato"] || v.Nombre_Candidato || '';
            const candidateLastName = v.candidateLastName || v["Apellido Candidato"] || v.Apellido_Candidato || '';
            const voterFirstName = v.voterFirstName || v["Nombre Miembro"] || v.Nombre_Miembro || v.NombreMiembro || '';
            const voterLastName = v.voterLastName || v["Apellido Miembro"] || v.Apellido_Miembro || v.ApellidoMiembro || '';
            
            // Fallbacks por si la hoja aún tiene formato antiguo (un solo campo)
            const oldCandName = v.candidateName || v.Nombre_Candidato || v["Nombre Candidato"] || v.Solicitante || v.solicitante || '';
            const oldVoterName = v.voterName || v.Nombre_Votante || v["Nombre Votante"] || v["Nombre del Miembro"] || '';

            const finalCandFirst = candidateFirstName || oldCandName.split(' ')[0] || '';
            const finalCandLast = candidateLastName || oldCandName.split(' ').slice(1).join(' ') || '';
            const finalVoterFirst = voterFirstName || oldVoterName.split(' ')[0] || '';
            const finalVoterLast = voterLastName || oldVoterName.split(' ').slice(1).join(' ') || '';

            return {
              candidateId: v.candidateId || v.ID_Candidato || v["ID Candidato"] || v.candidateid || '',
              candidateFirstName: finalCandFirst,
              candidateLastName: finalCandLast,
              candidateName: `${finalCandFirst} ${finalCandLast}`.trim(),
              voterFirstName: finalVoterFirst,
              voterLastName: finalVoterLast,
              voterName: `${finalVoterFirst} ${finalVoterLast}`.trim(),
              status: v.status || v.Estado || v.estado || v.status || '',
              reason: v.reason || v.Motivo || v.motivo || v.reason || '',
              timestamp: v.timestamp || v.Fecha || v.fecha || v.timestamp || new Date().toISOString()
            };
          });

          // Mapear candidatos desde la hoja "Candidatos" de votersData
          let sheetCandidates = votersData.candidates || [];

          // Auto-sembrado de candidatos de prueba si la hoja en Google Sheets está vacía y no fue vaciada explícitamente por el usuario
          if (sheetCandidates.length === 0 && localStorage.getItem('icc_candidates_cleared_by_user') !== 'true' && activeConfig.sheetUrlVoters) {
            try {
              console.log("Detectados candidatos vacíos en Google Sheets. Registrando candidatos de prueba...");
              for (const cand of MOCK_CANDIDATES) {
                const alreadyExists = sheetCandidates.some(c => 
                  String(c.id || c.ID || c.Id || '').trim() === String(cand.id).trim() ||
                  ((c.firstName || c.Nombre || '') === cand.firstName && (c.lastName || c.Apellido || '') === cand.lastName)
                );
                if (!alreadyExists) {
                  await addCandidateToSheets(activeConfig.sheetUrlVoters, {
                    id: cand.id,
                    firstName: cand.firstName,
                    lastName: cand.lastName,
                    testimony: cand.testimony
                  });
                  if (!localStorage.getItem(`icc_photo_${cand.id}`)) {
                    try {
                      localStorage.setItem(`icc_photo_${cand.id}`, cand.photo);
                    } catch (e) {
                      console.error(`Error al guardar foto en localStorage para el candidato ${cand.id}:`, e);
                    }
                  }
                }
              }
              // Volver a consultar datos actualizados
              const refetchedData = activeConfig.sheetUrlVoters === activeConfig.sheetUrlCandidates
                ? await fetchVotersFromSheets(activeConfig.sheetUrlVoters)
                : await fetchCandidatesAndVotesFromSheets(activeConfig.sheetUrlCandidates);
              sheetCandidates = refetchedData.candidates || [];
            } catch (seedError) {
              console.error("Error al sembrar candidatos automáticamente:", seedError);
            }
          }

          const mappedCandidates = sheetCandidates.map(c => {
            const id = String(c.id || c.ID || c.Id || '').trim();
            const localPhoto = localStorage.getItem(`icc_photo_${id}`);
            const sheetPhoto = c.photo || c.Foto || c.foto || '';
            return {
              id: id,
              firstName: c.firstName || c.Nombre || c.nombre || '',
              lastName: c.lastName || c.Apellido || c.apellido || '',
              testimony: c.testimony || c.Testimonio || c.testimonio || '',
              photo: sheetPhoto || localPhoto || '',
              createdAt: c.createdAt || new Date().toISOString()
            };
          });
          setCandidates(mappedCandidates);

          setVoters(mappedVoters);
          setVotes(mappedVotes);
          setIsConnected(true);
          hasLoadedFromSheetsRef.current = true;
          if (showNotification) {
            showToast("Datos sincronizados con Google Sheets", "success");
          }
        } catch (error) {
          console.error(error);
          setIsConnected(false);
          
          if (showNotification) {
            showToast("Error de conexión con Google Sheets. No se pudieron obtener los datos de la nube.", "error");
          }

          // Si nunca hemos logrado cargar desde Sheets, vaciamos el estado para no mostrar datos residuales o locales falsos
          if (!hasLoadedFromSheetsRef.current) {
            setCandidates([]);
            setVoters([]);
            setVotes([]);
          }
        }
      } else {
        setIsConnected(false);
        let localVoters = localStorage.getItem('icc_local_voters');
        let localVotes = localStorage.getItem('icc_local_votes');

        let parsedVoters = localVoters ? JSON.parse(localVoters) : MOCK_VOTERS;
        if (parsedVoters.length > 0 && typeof parsedVoters[0] === 'string') {
          parsedVoters = parsedVoters.map((v, index) => {
            const parts = String(v).trim().split(' ');
            return {
              id: `voter-${index + 1}`,
              name: parts[0] || '',
              lastName: parts.slice(1).join(' ') || ''
            };
          });
          localStorage.setItem('icc_local_voters', JSON.stringify(parsedVoters));
        }

        setVoters(parsedVoters);
        setVotes(localVotes ? JSON.parse(localVotes) : MOCK_VOTES);
        
        if (!localVoters) {
          localStorage.setItem('icc_local_voters', JSON.stringify(MOCK_VOTERS));
          localStorage.setItem('icc_local_votes', JSON.stringify(MOCK_VOTES));
        }
      }
    } finally {
      setIsLoading(false);
    }
  }, [config]);

  // Cargar al montar el componente o cuando cambie la URL de sheets
  useEffect(() => {
    // Migración única para borrar registros locales anteriores de votantes y votos
    if (localStorage.getItem('icc_voters_cleared_2026') !== 'true') {
      localStorage.removeItem('icc_local_voters');
      localStorage.removeItem('icc_local_votes');
      localStorage.setItem('icc_voters_cleared_2026', 'true');
    }
    // Migración única para borrar candidatos locales
    if (localStorage.getItem('icc_candidates_cleared_2026') !== 'true') {
      localStorage.removeItem('icc_local_candidates');
      const keysToDelete = [];
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key && key.startsWith('icc_photo_')) {
          keysToDelete.push(key);
        }
      }
      keysToDelete.forEach(k => localStorage.removeItem(k));
      localStorage.setItem('icc_candidates_cleared_2026', 'true');
    }
    const timer = setTimeout(() => {
      loadData(config, true);
    }, 0);
    return () => clearTimeout(timer);
  }, [config.sheetUrlVoters, config.sheetUrlCandidates, loadData, config]);


  return (
    <div className="app-container">
      {/* Toast notifications container */}
      <div className="toast-container">
        {toasts.map((toast) => (
          <div key={toast.id} className={`toast toast-${toast.type}`}>
            {toast.type === 'success' && <span>✓</span>}
            {toast.type === 'error' && <span>⚠</span>}
            {toast.type === 'info' && <span>ℹ</span>}
            <div>{toast.message}</div>
          </div>
        ))}
      </div>

      {/* Header General */}
      <header className="app-header">
        <div 
          className="logo-section" 
          onClick={view === 'voter' ? undefined : () => navigate('landing')} 
          style={{ cursor: view === 'voter' ? 'default' : 'pointer' }}
        >
          <div>
            <span className="church-logo">
              <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                <path d="M12 2L2 22h20L12 2z"/>
                <path d="M12 7v10M9 11h6"/>
              </svg>
              ICC
            </span>
            <div className="app-title-sub">Convertidos a Cristo</div>
          </div>
        </div>

        <div className="nav-buttons">
          {view !== 'landing' && view !== 'voter' && (
            <Tooltip text="Regresar a la pantalla de bienvenida principal." position="bottom">
              <button className="btn btn-secondary" onClick={() => navigate('landing')}>
                Inicio
              </button>
            </Tooltip>
          )}
          {view === 'landing' && (
            <>
              <Tooltip text="Ir directamente a la cabina de votación." position="bottom">
                <button className="btn btn-primary" onClick={() => navigate('voter')}>
                  Votar
                </button>
              </Tooltip>
              <Tooltip text="Ir al panel de administración del sistema." position="bottom">
                <button className="btn btn-secondary" onClick={() => navigate('admin')}>
                  Administración
                </button>
              </Tooltip>
            </>
          )}
          {isConnected && view !== 'landing' && view !== 'voter' && (
            <Tooltip text="Sincronizar y descargar los últimos datos desde Google Sheets." position="bottom">
              <button 
                className={`btn btn-secondary ${isLoading ? 'loading' : ''}`} 
                onClick={() => loadData(config, true)}
                disabled={isLoading}
              >
                {isLoading ? 'Sincronizando...' : 'Sincronizar'}
              </button>
            </Tooltip>
          )}
        </div>
      </header>

      {/* Renderizado de vistas */}
      {view === 'landing' && (
        <main style={{ flexGrow: 1, display: 'flex', flexDirection: 'column', justifyContent: 'center', alignItems: 'center', py: 40 }}>
          <div style={{ textAlign: 'center', maxWidth: '600px', margin: 'auto' }}>
            <h1 style={{ fontFamily: 'var(--font-heading)', fontWeight: 800, fontSize: '3.5rem', lineHeight: '1.1', marginBottom: '16px', background: 'linear-gradient(135deg, var(--primary), var(--accent))', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent' }}>
              Sistema de Votación de Membresía
            </h1>
            <p style={{ color: 'var(--text-secondary)', fontSize: '1.25rem', marginBottom: '40px' }}>
              Bienvenido al portal oficial de recepción de membresía de la Iglesia de Convertidos a Cristo. Selecciona una opción para comenzar.
            </p>

            <div className="grid-2" style={{ marginTop: '24px' }}>
              <div className="card" onClick={() => navigate('voter')} style={{ cursor: 'pointer', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '16px', padding: '40px 24px', textAlign: 'center' }}>
                <div style={{ width: '64px', height: '64px', borderRadius: '50%', backgroundColor: 'var(--primary-light)', display: 'flex', alignItems: 'center', justify: 'center', color: 'var(--primary)', fontSize: '24px', justifyContent: 'center' }}>
                  📥
                </div>
                <h3 style={{ fontFamily: 'var(--font-heading)', fontSize: '20px', fontWeight: 700 }}>Panel de Votación</h3>
                <p style={{ color: 'var(--text-muted)', fontSize: '14px' }}>
                  Accede para emitir tu voto de aprobación de nuevos miembros utilizando tu nombre.
                </p>
                <Tooltip text="Iniciar el asistente de votación para emitir tu veredicto." position="top">
                  <button className="btn btn-primary" style={{ width: '100%', marginTop: '12px' }}>Ingresar como Votante</button>
                </Tooltip>
              </div>

              <div className="card" onClick={() => navigate('admin')} style={{ cursor: 'pointer', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '16px', padding: '40px 24px', textAlign: 'center' }}>
                <div style={{ width: '64px', height: '64px', borderRadius: '50%', backgroundColor: 'var(--accent-light)', display: 'flex', alignItems: 'center', justify: 'center', color: 'var(--accent)', fontSize: '24px', justifyContent: 'center' }}>
                  ⚙️
                </div>
                <h3 style={{ fontFamily: 'var(--font-heading)', fontSize: '20px', fontWeight: 700 }}>Administración</h3>
                <p style={{ color: 'var(--text-muted)', fontSize: '14px' }}>
                  Registra candidatos, sube fotos, administra la lista de votantes autorizados y visualiza resultados.
                </p>
                <Tooltip text="Acceder a las herramientas de control y configuración." position="top">
                  <button className="btn btn-secondary" style={{ width: '100%', marginTop: '12px' }}>Ingresar como Administrador</button>
                </Tooltip>
              </div>
            </div>

            <div style={{ marginTop: '48px' }}>
              {!isConnected ? (
                <span className="conn-status status-disconnected">
                  ⚡ Modo Local Activado (Configurar Google Sheets en Admin)
                </span>
              ) : (
                <span className="conn-status status-connected">
                  ● Conectado a Google Sheets
                </span>
              )}
            </div>
          </div>
        </main>
      )}

      {view === 'admin' && (
        <AdminPanel
          config={config}
          saveConfig={saveConfig}
          candidates={candidates}
          voters={voters}
          votes={votes}
          isLoading={isLoading}
          isConnected={isConnected}
          refreshData={loadData}
          showToast={showToast}
          setCandidates={setCandidates}
          setVoters={setVoters}
          setVotes={setVotes}
        />
      )}

      {view === 'voter' && (
        <VoterPanel
          config={config}
          candidates={candidates}
          voters={voters}
          votes={votes}
          isConnected={isConnected}
          refreshData={loadData}
          showToast={showToast}
          setVotes={setVotes}
          isLoading={isLoading}
        />
      )}
    </div>
  );
}

export default App;
