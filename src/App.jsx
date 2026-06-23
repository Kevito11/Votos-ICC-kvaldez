import { useState, useEffect, useCallback, useRef } from 'react';
import AdminPanel from './components/AdminPanel';
import VoterPanel from './components/VoterPanel';
import Tooltip from './components/Tooltip';
import { fetchVotersFromSheets, fetchCandidatesAndVotesFromSheets, addCandidateToSheets } from './utils/api';
import { MOCK_CANDIDATES, MOCK_VOTERS, MOCK_VOTES } from './utils/mockData';
import './App.css';


function App() {
  // Configuración de conexiones (definida en código y sobreescribible vía variables de entorno .env / Netlify)
  const [config] = useState(() => {
    const defaultSheetUrl = import.meta.env.VITE_SHEET_URL || 'https://script.google.com/macros/s/AKfycbzh8dvSWSPw7UqhJCU0-xsUs_aFZwAN2ytzVUW_19wwHKLUdc6BEFefnRckVmfoDI-aOA/exec';
    return {
      sheetUrl: defaultSheetUrl,
      sheetUrlVoters: import.meta.env.VITE_SHEET_URL_VOTERS || defaultSheetUrl,
      sheetUrlCandidates: import.meta.env.VITE_SHEET_URL_CANDIDATES || defaultSheetUrl,
      supabaseUrl: import.meta.env.VITE_SUPABASE_URL || '',
      supabaseKey: import.meta.env.VITE_SUPABASE_KEY || '',
      supabaseBucket: import.meta.env.VITE_SUPABASE_BUCKET || 'candidatos'
    };
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
  
  // Seguridad y autenticación del Administrador
  const [isAdminAuthenticated, setIsAdminAuthenticated] = useState(() => {
    return sessionStorage.getItem('icc_admin_authenticated') === 'true';
  });
  const [adminPasswordInput, setAdminPasswordInput] = useState('');
  const [adminLoginError, setAdminLoginError] = useState(false);

  const handleAdminLoginSubmit = (e) => {
    e.preventDefault();
    const correctPassword = import.meta.env.VITE_ADMIN_PASSWORD || '1234';
    if (adminPasswordInput === correctPassword) {
      setIsAdminAuthenticated(true);
      sessionStorage.setItem('icc_admin_authenticated', 'true');
      setAdminPasswordInput('');
      setAdminLoginError(false);
      showToast("Acceso de administrador concedido", "success");
    } else {
      setAdminLoginError(true);
      showToast("Contraseña incorrecta", "error");
    }
  };

  const handleAdminLogout = () => {
    setIsAdminAuthenticated(false);
    sessionStorage.removeItem('icc_admin_authenticated');
    showToast("Sesión de administrador cerrada", "info");
  };

  // Guardar referencia actualizada de la función logout para evitar re-vincular los listeners
  const logoutRef = useRef(handleAdminLogout);
  useEffect(() => {
    logoutRef.current = handleAdminLogout;
  });

  // Cierre de sesión automático tras 5 minutos de inactividad
  useEffect(() => {
    if (!isAdminAuthenticated) return;

    let inactivityTimer;

    const resetInactivityTimer = () => {
      if (inactivityTimer) clearTimeout(inactivityTimer);
      inactivityTimer = setTimeout(() => {
        logoutRef.current();
        showToast("Sesión de administrador cerrada por inactividad de 5 minutos", "warning");
      }, 5 * 60 * 1000); // 5 minutos (300000 ms)
    };

    const userActivityEvents = ['mousemove', 'mousedown', 'keypress', 'scroll', 'touchstart'];
    userActivityEvents.forEach(event => window.addEventListener(event, resetInactivityTimer));

    resetInactivityTimer();

    return () => {
      if (inactivityTimer) clearTimeout(inactivityTimer);
      userActivityEvents.forEach(event => window.removeEventListener(event, resetInactivityTimer));
    };
  }, [isAdminAuthenticated]);
  
  // Enrutamiento simple basado en URL (rutas limpias)
  const [view, setView] = useState(() => {
    const path = window.location.pathname.replace(/^\/|\/$/g, '');
    return (path === '' || path === 'index.html') ? 'landing' : path;
  });

  // Escuchar cambios de historial en el navegador
  useEffect(() => {
    const handlePopState = () => {
      const path = window.location.pathname.replace(/^\/|\/$/g, '');
      setView((path === '' || path === 'index.html') ? 'landing' : path);
    };
    window.addEventListener('popstate', handlePopState);
    return () => window.removeEventListener('popstate', handlePopState);
  }, []);

  // Función para navegar actualizando la URL con rutas limpias
  const navigate = (newView) => {
    setView(newView);
    const url = newView === 'landing' ? '/' : `/${newView}`;
    window.history.pushState({ view: newView }, '', url);
  };

  // (saveConfig ha sido removido porque la configuración ahora se gestiona directamente en código/entorno)

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
              candidateId: String(v.candidateId || v.ID_Candidato || v["ID Candidato"] || v.candidateid || '').trim(),
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
        <div className="logo-section">
          <picture>
            <source srcSet="/logo-white.png" media="(prefers-color-scheme: dark)" />
            <img 
              src="/logo-color.png" 
              alt="Iglesia de Convertidos a Cristo" 
              className="church-logo-img" 
            />
          </picture>
        </div>

        <div className="nav-buttons">
          {isConnected && view === 'admin' && isAdminAuthenticated && (
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
          {view === 'admin' && isAdminAuthenticated && (
            <Tooltip text="Cerrar la sesión administrativa actual." position="bottom">
              <button 
                className="btn btn-danger" 
                onClick={handleAdminLogout}
                style={{ padding: '8px 14px', fontSize: '13px' }}
              >
                Cerrar Sesión
              </button>
            </Tooltip>
          )}
        </div>
      </header>

      {/* Renderizado de vistas */}
      {view === 'landing' && (
        <main style={{ flexGrow: 1, display: 'flex', flexDirection: 'column', justifyContent: 'center', alignItems: 'center', py: 40 }}>
          <div className="card" style={{ textAlign: 'center', maxWidth: '480px', margin: 'auto', padding: '40px 32px', borderRadius: 'var(--radius-lg)', boxShadow: 'var(--shadow-xl)', border: '1px solid var(--border)' }}>
            <div style={{ marginBottom: '24px' }}>
              <picture style={{ display: 'block', margin: '0 auto' }}>
                <source srcSet="/logo-white.png" media="(prefers-color-scheme: dark)" />
                <img 
                  src="/logo-color.png" 
                  alt="Iglesia de Convertidos a Cristo" 
                  className="landing-logo"
                  style={{ height: '54px', maxWidth: '100%', objectFit: 'contain' }}
                />
              </picture>
            </div>
            <div style={{ width: '80px', height: '80px', borderRadius: '50%', backgroundColor: 'var(--accent-light)', display: 'flex', alignItems: 'center', color: 'var(--accent)', fontSize: '36px', margin: '0 auto 24px', justifyContent: 'center' }}>
              🔒
            </div>
            <h2 style={{ fontFamily: 'var(--font-heading)', fontWeight: 800, fontSize: '24px', marginBottom: '16px', color: 'var(--text-primary)' }}>
              Acceso Restringido
            </h2>
            <p style={{ color: 'var(--text-secondary)', fontSize: '15px', lineHeight: '1.6', marginBottom: '16px' }}>
              Esta dirección no cuenta con un portal de acceso público para el sistema de votación.
            </p>
            <p style={{ color: 'var(--text-muted)', fontSize: '13.5px', lineHeight: '1.6' }}>
              Por favor, utilice los enlaces directos o los códigos QR oficiales proporcionados por la administración de la iglesia para ingresar al panel de votación o administración.
            </p>
          </div>
        </main>
      )}

      {view === 'admin' && !isAdminAuthenticated && (
        <main style={{ flexGrow: 1, display: 'flex', flexDirection: 'column', justifyContent: 'center', alignItems: 'center', py: 40 }}>
          <form onSubmit={handleAdminLoginSubmit} className="card" style={{ textAlign: 'center', maxWidth: '380px', width: '100%', margin: 'auto', padding: '32px 24px', borderRadius: 'var(--radius-lg)', boxShadow: 'var(--shadow-xl)', border: '1px solid var(--border)' }}>
            <div style={{ marginBottom: '20px' }}>
              <picture style={{ display: 'block', margin: '0 auto' }}>
                <source srcSet="/logo-white.png" media="(prefers-color-scheme: dark)" />
                <img 
                  src="/logo-color.png" 
                  alt="Iglesia de Convertidos a Cristo" 
                  className="login-logo"
                  style={{ height: '48px', maxWidth: '100%', objectFit: 'contain' }}
                />
              </picture>
            </div>
            <div style={{ width: '64px', height: '64px', borderRadius: '50%', backgroundColor: 'var(--primary-light)', display: 'flex', alignItems: 'center', color: 'var(--primary)', fontSize: '28px', margin: '0 auto 16px', justifyContent: 'center' }}>
              🔑
            </div>
            <h2 style={{ fontFamily: 'var(--font-heading)', fontWeight: 800, fontSize: '20px', marginBottom: '8px', color: 'var(--text-primary)' }}>
              Acceso de Administrador
            </h2>
            <p style={{ color: 'var(--text-secondary)', fontSize: '13.5px', marginBottom: '24px' }}>
              Ingrese la contraseña de seguridad para acceder al panel de control y configuraciones.
            </p>

            <div className="form-group" style={{ marginBottom: '16px' }}>
              <input 
                type="password" 
                className="form-control" 
                placeholder="Contraseña"
                value={adminPasswordInput}
                onChange={(e) => setAdminPasswordInput(e.target.value)}
                style={{ textAlign: 'center', fontSize: '16px', padding: '12px' }}
                autoFocus
                required
              />
              {adminLoginError && (
                <div style={{ color: 'var(--danger)', fontSize: '12.5px', marginTop: '8px', fontWeight: '600' }}>
                  Contraseña incorrecta. Intente de nuevo.
                </div>
              )}
            </div>

            <button type="submit" className="btn btn-primary" style={{ width: '100%', padding: '12px' }}>
              Ingresar al Panel
            </button>
          </form>
        </main>
      )}

      {view === 'admin' && isAdminAuthenticated && (
        <AdminPanel
          config={config}
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
