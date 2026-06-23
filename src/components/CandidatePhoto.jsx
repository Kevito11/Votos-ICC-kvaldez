import { useState } from 'react';
import { getDriveFallbackUrls } from '../utils/api';

/**
 * Componente de imagen de candidato con reintentos automáticos de URL.
 * Si la URL principal falla, prueba alternativas de Google Drive.
 * Si todas fallan, muestra las iniciales del candidato.
 */
export default function CandidatePhoto({ photo, firstName = '', lastName = '', className = '', style = {} }) {
  const fallbacks = photo ? getDriveFallbackUrls(photo) : [];
  const [urlIndex, setUrlIndex] = useState(0);
  const [allFailed, setAllFailed] = useState(!photo);

  const handleError = () => {
    const nextIndex = urlIndex + 1;
    if (nextIndex < fallbacks.length) {
      setUrlIndex(nextIndex);
    } else {
      setAllFailed(true);
    }
  };

  if (allFailed || !photo) {
    return (
      <div
        className={`placeholder-avatar ${className}`}
        style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          backgroundColor: 'var(--primary-light)',
          color: 'var(--primary)',
          fontWeight: 'bold',
          borderRadius: '50%',
          userSelect: 'none',
          ...style,
        }}
      >
        {firstName?.charAt(0)?.toUpperCase()}{lastName?.charAt(0)?.toUpperCase()}
      </div>
    );
  }

  return (
    <img
      src={fallbacks[urlIndex]}
      className={className}
      alt={`${firstName} ${lastName}`}
      onError={handleError}
      style={{ objectFit: 'cover', ...style }}
    />
  );
}
