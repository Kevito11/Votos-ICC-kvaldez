import { useState, useRef, useEffect } from 'react';

/**
 * Componente Tooltip que envuelve a un elemento y muestra una descripción flotante
 * al pasar el cursor (hover) o al mantener presionado en pantallas táctiles (long-press).
 *
 * @param {string} text - Texto descriptivo a mostrar.
 * @param {React.ReactNode} children - Elemento hijo al que se aplica el tooltip.
 * @param {string} position - Ubicación ('top', 'bottom', 'left', 'right'). Por defecto: 'top'.
 */
export default function Tooltip({ text, children, position = 'top', style }) {
  const [isVisible, setIsVisible] = useState(false);
  const [coords, setCoords] = useState({ top: 0, left: 0 });
  const [hShift, setHShift] = useState(0);
  const [activePosition, setActivePosition] = useState(position);
  const containerRef = useRef(null);
  const hoverTimeout = useRef(null);
  const touchTimeout = useRef(null);
  const touchStartPos = useRef({ x: 0, y: 0 });
  const isLongPress = useRef(false);

  // Si no hay texto para describir, renderizar los hijos tal cual
  if (!text) return children;

  const showTooltip = () => {
    if (containerRef.current) {
      const rect = containerRef.current.getBoundingClientRect();
      let top = 0;
      let left = 0;
      let shift = 0;
      
      const centerX = rect.left + rect.width / 2;
      
      const isMobile = window.innerWidth <= 768;
      // En móvil, preferimos siempre 'top' para evitar la obstrucción del dedo,
      // a menos que el elemento esté tan arriba en la pantalla que el tooltip se saldría del viewport (espacio < 80px).
      let effectivePosition = position;
      if (isMobile) {
        effectivePosition = rect.top < 80 ? 'bottom' : 'top';
      }
      
      if (effectivePosition === 'top' || effectivePosition === 'bottom') {
        const bubbleWidth = 220; // Ancho máximo aproximado del tooltip
        const minLeft = 12;
        const maxRight = window.innerWidth - 12;
        const isRightHalf = centerX > window.innerWidth / 2;
        
        // Desplazamiento base de 45px hacia el lado opuesto al borde más cercano
        const baseShift = isRightHalf ? -45 : 45;
        let newCenterX = centerX + baseShift;
        
        // Mantener dentro de los límites de la pantalla
        if (newCenterX - bubbleWidth / 2 < minLeft) {
          newCenterX = minLeft + bubbleWidth / 2;
        } else if (newCenterX + bubbleWidth / 2 > maxRight) {
          newCenterX = maxRight - bubbleWidth / 2;
        }
        
        shift = newCenterX - centerX;
      }
      
      if (effectivePosition === 'top') {
        top = rect.top;
        left = centerX;
      } else if (effectivePosition === 'bottom') {
        top = rect.bottom;
        left = centerX;
      } else if (effectivePosition === 'left') {
        top = rect.top + rect.height / 2;
        left = rect.left;
      } else if (effectivePosition === 'right') {
        top = rect.top + rect.height / 2;
        left = rect.right;
      }
      
      setCoords({ top, left });
      setHShift(shift);
      setActivePosition(effectivePosition);
    }
    setIsVisible(true);
  };

  const hideTooltip = () => {
    setIsVisible(false);
    if (hoverTimeout.current) clearTimeout(hoverTimeout.current);
    if (touchTimeout.current) clearTimeout(touchTimeout.current);
  };

  // --- Manejo en Escritorio (Mouse) ---
  const handleMouseEnter = () => {
    if (hoverTimeout.current) clearTimeout(hoverTimeout.current);
    hoverTimeout.current = setTimeout(() => {
      showTooltip();
    }, 500); // 500ms de retraso para evitar tooltips accidentales en pasadas rápidas
  };

  const handleMouseLeave = () => {
    hideTooltip();
  };

  // --- Manejo en Dispositivos Táctiles (Touch) ---
  const handleTouchStart = (e) => {
    const touch = e.touches[0];
    touchStartPos.current = { x: touch.clientX, y: touch.clientY };
    isLongPress.current = false;
    if (touchTimeout.current) clearTimeout(touchTimeout.current);
    touchTimeout.current = setTimeout(() => {
      isLongPress.current = true;
      showTooltip();
      // Vibración de respuesta táctil (si está soportada)
      if (navigator.vibrate) {
        navigator.vibrate(40);
      }
    }, 500); // 500ms es el estándar para pulsación larga
  };

  const handleTouchEnd = (e) => {
    if (isLongPress.current) {
      // Prevenir el click / activación de la función del botón tras el long-press
      e.preventDefault();
      // Prevenir que el evento bubble a otros contenedores
      e.stopPropagation();
    }
    hideTooltip();
  };

  const handleTouchMove = (e) => {
    const touch = e.touches[0];
    const dx = touch.clientX - touchStartPos.current.x;
    const dy = touch.clientY - touchStartPos.current.y;
    const distance = Math.sqrt(dx * dx + dy * dy);
    // Permitir un margen de movimiento (10px) para evitar cancelaciones por vibración natural del dedo
    if (distance > 10) {
      hideTooltip();
    }
  };

  // Ocultar el tooltip al hacer scroll en cualquier parte para evitar desprendimiento
  useEffect(() => {
    if (!isVisible) return;
    
    const handleScroll = () => {
      hideTooltip();
    };
    
    window.addEventListener('scroll', handleScroll, { passive: true });
    window.addEventListener('wheel', handleScroll, { passive: true });

    return () => {
      window.removeEventListener('scroll', handleScroll);
      window.removeEventListener('wheel', handleScroll);
    };
  }, [isVisible]);

  // Limpiar temporizadores cuando el componente se desmonte
  useEffect(() => {
    return () => {
      if (hoverTimeout.current) clearTimeout(hoverTimeout.current);
      if (touchTimeout.current) clearTimeout(touchTimeout.current);
    };
  }, []);

  return (
    <div 
      ref={containerRef}
      className="tooltip-container"
      style={style}
      onMouseEnter={handleMouseEnter}
      onMouseLeave={handleMouseLeave}
      onTouchStart={handleTouchStart}
      onTouchEnd={handleTouchEnd}
      onTouchMove={handleTouchMove}
      onTouchCancel={hideTooltip}
    >
      {children}
      {isVisible && (
        <div 
          className={`tooltip-bubble tooltip-${activePosition}`}
          style={{
            position: 'fixed',
            top: `${coords.top}px`,
            left: `${coords.left}px`,
            bottom: 'auto',
            right: 'auto',
            '--h-shift': `${hShift}px`,
            transform: 
              activePosition === 'top' ? 'translate(calc(-50% + var(--h-shift, 0px)), -100%) translate(0, -8px)' :
              activePosition === 'bottom' ? 'translate(calc(-50% + var(--h-shift, 0px)), 8px)' :
              activePosition === 'left' ? 'translate(-100%, -50%) translate(-8px, 0)' :
              'translate(8px, -50%)'
          }}
        >
          {text}
          <div className="tooltip-arrow"></div>
        </div>
      )}
    </div>
  );
}
