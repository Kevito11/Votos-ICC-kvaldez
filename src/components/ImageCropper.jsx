import { useState, useRef, useEffect } from 'react';
import Tooltip from './Tooltip';


export default function ImageCropper({ imageFile, onCrop, onCancel }) {
  const [imageSrc, setImageSrc] = useState('');
  const [zoom, setZoom] = useState(1);
  const [offset, setOffset] = useState({ x: 0, y: 0 });
  const [isDragging, setIsDragging] = useState(false);
  const dragStart = useRef({ x: 0, y: 0 });
  const imgRef = useRef(null);
  const containerRef = useRef(null);
  const [imgDimensions, setImgDimensions] = useState({ width: 0, height: 0 });

  useEffect(() => {
    if (imageFile) {
      const reader = new FileReader();
      reader.readAsDataURL(imageFile);
      reader.onload = () => {
        setImageSrc(reader.result);
      };
      reader.onerror = (err) => {
        console.error("Error reading file:", err);
      };
    }
  }, [imageFile]);

  const handleImageLoad = (e) => {
    const img = e.target;
    const containerSize = 300;
    const aspect = img.naturalWidth / img.naturalHeight;
    let w, h;
    if (aspect > 1) {
      // Landscape: fit height first, scale width
      h = containerSize;
      w = containerSize * aspect;
    } else {
      // Portrait or square: fit width first, scale height
      w = containerSize;
      h = containerSize / aspect;
    }
    setImgDimensions({ width: w, height: h });
    setOffset({ x: 0, y: 0 });
    setZoom(1);
  };

  const handleMouseDown = (e) => {
    setIsDragging(true);
    dragStart.current = { x: e.clientX - offset.x, y: e.clientY - offset.y };
  };

  const handleMouseMove = (e) => {
    if (!isDragging) return;
    setOffset({
      x: e.clientX - dragStart.current.x,
      y: e.clientY - dragStart.current.y
    });
  };

  const handleMouseUp = () => {
    setIsDragging(false);
  };

  // Touch Support
  const handleTouchStart = (e) => {
    if (e.touches.length === 1) {
      setIsDragging(true);
      dragStart.current = {
        x: e.touches[0].clientX - offset.x,
        y: e.touches[0].clientY - offset.y
      };
    }
  };

  const handleTouchMove = (e) => {
    if (!isDragging || e.touches.length !== 1) return;
    setOffset({
      x: e.touches[0].clientX - dragStart.current.x,
      y: e.touches[0].clientY - dragStart.current.y
    });
  };

  const handleCrop = () => {
    const img = imgRef.current;
    const container = containerRef.current;
    if (!img || !container) return;

    const canvas = document.createElement('canvas');
    canvas.width = 200;
    canvas.height = 200;
    const ctx = canvas.getContext('2d');

    const containerWidth = container.offsetWidth;
    const containerHeight = container.offsetHeight;

    // Crop circle dimension (180x180 px box)
    const cropSize = 180;
    const cropX = (containerWidth - cropSize) / 2;
    const cropY = (containerHeight - cropSize) / 2;

    // Get displayed client rect of both elements
    const imgRect = img.getBoundingClientRect();
    const containerRect = container.getBoundingClientRect();

    // Find top-left coordinates of the displayed scaled image relative to the container
    const imgLeft = imgRect.left - containerRect.left;
    const imgTop = imgRect.top - containerRect.top;
    
    // Scale factor to map displayed coordinates to original image coordinates
    const scaleX = img.naturalWidth / imgRect.width;
    const scaleY = img.naturalHeight / imgRect.height;

    // Source coordinates on the original image
    const sourceX = (cropX - imgLeft) * scaleX;
    const sourceY = (cropY - imgTop) * scaleY;
    const sourceWidth = cropSize * scaleX;
    const sourceHeight = cropSize * scaleY;

    // Clear and draw on canvas
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, 200, 200);

    ctx.drawImage(
      img,
      sourceX,
      sourceY,
      sourceWidth,
      sourceHeight,
      0,
      0,
      200,
      200
    );

    canvas.toBlob((blob) => {
      if (blob) {
        // Create new File from blob to preserve candidate save structure
        const croppedFile = new File([blob], imageFile.name, {
          type: 'image/jpeg',
          lastModified: Date.now()
        });
        onCrop(croppedFile);
      }
    }, 'image/jpeg', 0.85);
  };

  return (
    <div className="crop-modal-overlay" style={{
      position: 'fixed',
      top: 0,
      left: 0,
      right: 0,
      bottom: 0,
      backgroundColor: 'rgba(15, 23, 42, 0.85)',
      backdropFilter: 'blur(8px)',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      zIndex: 2000,
      animation: 'fadeIn 0.2s ease-out'
    }}>
      <div className="card crop-modal-card" style={{
        backgroundColor: 'var(--bg-card)',
        border: '1px solid var(--border)',
        borderRadius: 'var(--radius-lg)',
        padding: '24px',
        width: '90%',
        maxWidth: '350px',
        boxShadow: 'var(--shadow-xl)',
        textAlign: 'center'
      }}>
        <h3 style={{ fontFamily: 'var(--font-heading)', fontWeight: 700, marginTop: 0, marginBottom: '8px' }}>
          Ajustar Fotografía
        </h3>
        <p style={{ color: 'var(--text-secondary)', fontSize: '13px', marginBottom: '20px' }}>
          Arrastra la imagen y usa el zoom para encuadrar el rostro del candidato en el círculo.
        </p>

        {/* Viewport de recorte interactivo */}
        <div 
          ref={containerRef}
          onMouseDown={handleMouseDown}
          onMouseMove={handleMouseMove}
          onMouseUp={handleMouseUp}
          onMouseLeave={handleMouseUp}
          onTouchStart={handleTouchStart}
          onTouchMove={handleTouchMove}
          onTouchEnd={handleMouseUp}
          style={{
            position: 'relative',
            width: '300px',
            height: '300px',
            margin: '0 auto 16px',
            overflow: 'hidden',
            borderRadius: 'var(--radius-md)',
            backgroundColor: '#0f172a',
            cursor: 'move',
            userSelect: 'none',
            touchAction: 'none'
          }}
        >
          {imageSrc && (
            <img 
              ref={imgRef}
              src={imageSrc}
              onLoad={handleImageLoad}
              alt="Para recortar"
              style={{
                position: 'absolute',
                top: '50%',
                left: '50%',
                transform: `translate(-50%, -50%) translate(${offset.x}px, ${offset.y}px) scale(${zoom})`,
                width: imgDimensions.width ? `${imgDimensions.width}px` : 'auto',
                height: imgDimensions.height ? `${imgDimensions.height}px` : 'auto',
                maxWidth: 'none',
                maxHeight: 'none',
                pointerEvents: 'none',
                userSelect: 'none'
              }}
            />
          )}

          {/* Máscara de recorte circular SVG */}
          <svg style={{
            position: 'absolute',
            top: 0,
            left: 0,
            width: '100%',
            height: '100%',
            pointerEvents: 'none'
          }}>
            <defs>
              <mask id="crop-mask">
                <rect width="100%" height="100%" fill="white" />
                <circle cx="50%" cy="50%" r="90" fill="black" />
              </mask>
            </defs>
            <rect width="100%" height="100%" fill="rgba(15, 23, 42, 0.6)" mask="url(#crop-mask)" />
            <circle cx="50%" cy="50%" r="90" stroke="var(--primary)" strokeWidth="2.5" fill="none" />
          </svg>
        </div>

        {/* Control deslizante de zoom */}
        <div style={{ marginBottom: '24px', textAlign: 'left' }}>
          <label className="form-label" style={{ display: 'flex', justifyContent: 'space-between', fontSize: '13px' }}>
            <span>Zoom</span>
            <span style={{ fontWeight: 'normal', color: 'var(--text-muted)' }}>{zoom.toFixed(1)}x</span>
          </label>
          <input 
            type="range" 
            min="0.2" 
            max="3" 
            step="0.05" 
            value={zoom}
            onChange={(e) => setZoom(parseFloat(e.target.value))}
            style={{
              width: '100%',
              accentColor: 'var(--primary)',
              cursor: 'pointer'
            }}
          />
        </div>

        {/* Botones de acción */}
        <div style={{ display: 'flex', gap: '12px' }}>
          <Tooltip text="Descartar la imagen y cerrar el editor." position="top" style={{ flex: 1 }}>
            <button 
              type="button" 
              className="btn btn-secondary" 
              onClick={onCancel}
              style={{ width: '100%' }}
            >
              Cancelar
            </button>
          </Tooltip>
          <Tooltip text="Recortar la imagen en forma de círculo y aplicarla al candidato." position="top" style={{ flex: 1 }}>
            <button 
              type="button" 
              className="btn btn-primary" 
              onClick={handleCrop}
              style={{ width: '100%' }}
            >
              Aceptar
            </button>
          </Tooltip>
        </div>
      </div>
      <style>{`
        @keyframes fadeIn {
          from { opacity: 0; }
          to { opacity: 1; }
        }
      `}</style>
    </div>
  );
}
