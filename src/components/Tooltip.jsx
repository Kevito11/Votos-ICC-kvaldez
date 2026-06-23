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
  const hoverTimeout = useRef(null);
  const touchTimeout = useRef(null);
  const isLongPress = useRef(false);

  // Si no hay texto para describir, renderizar los hijos tal cual
  if (!text) return children;

  const showTooltip = () => {
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
  const handleTouchStart = () => {
    isLongPress.current = false;
    if (touchTimeout.current) clearTimeout(touchTimeout.current);
    touchTimeout.current = setTimeout(() => {
      isLongPress.current = true;
      showTooltip();
      // Vibración de respuesta táctil (si está soportada)
      if (navigator.vibrate) {
        navigator.vibrate(40);
      }
    }, 800); // 800ms manteniendo presionado para activar la descripción
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

  const handleTouchMove = () => {
    // Si el usuario desliza el dedo, asumimos scroll o cancelación del gesto
    hideTooltip();
  };

  // Limpiar temporizadores cuando el componente se desmonte
  useEffect(() => {
    return () => {
      if (hoverTimeout.current) clearTimeout(hoverTimeout.current);
      if (touchTimeout.current) clearTimeout(touchTimeout.current);
    };
  }, []);

  return (
    <div 
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
        <div className={`tooltip-bubble tooltip-${position}`}>
          {text}
          <div className="tooltip-arrow"></div>
        </div>
      )}
    </div>
  );
}
