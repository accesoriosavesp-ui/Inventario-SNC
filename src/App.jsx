import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { initializeApp } from "firebase/app";
import { getFirestore, doc, getDoc, setDoc } from "firebase/firestore";

// Configuración oficial de tu Firebase Console
const firebaseConfig = {
  apiKey: "AIzaSyB8NwywIMy-w_7sXutsqk1XEmS50-in2Jk",
  authDomain: "inventario-1acf3.firebaseapp.com",
  projectId: "inventario-1acf3",
  storageBucket: "inventario-1acf3.firebasestorage.app",
  messagingSenderId: "475761738952",
  appId: "1:475761738952:web:7fce73e1fa272ae29651fd",
  measurementId: "G-SF97N6QV2K"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);
const DOC_ID = "mi_inventario_snc"; 

function App() {
  const [inventario, setInventario] = useState([]);
  const [busqueda, setBusqueda] = useState('');
  const [fasePantalla, setFasePantalla] = useState('bienvenida');

  // CARGA DESDE LA NUBE: Sustituye por completo al antiguo localStorage
  useEffect(() => {
    const cargarDatosDesdeFirebase = async () => {
      try {
        const docRef = doc(db, "datos", DOC_ID);
        const docSnap = await getDoc(docRef);
        if (docSnap.exists()) {
          setInventario(docSnap.data().items || []);
        }
      } catch (error) {
        console.error("Error al sincronizar con Firebase al arrancar:", error);
      }
    };
    cargarDatosDesdeFirebase();
  }, []);

  // GUARDAR EN LA NUBE: Reemplaza a guardarEnMemoria local
  const guardarEnFirebase = async (nuevoInventario) => {
    setInventario(nuevoInventario);
    try {
      await setDoc(doc(db, "datos", DOC_ID), { items: nuevoInventario });
    } catch (error) {
      console.error("Error al escribir datos en la nube:", error);
    }
  };

  const iniciarSesionAnimado = () => {
    setFasePantalla('saliendo'); 
    setTimeout(() => {
      setFasePantalla('dashboard'); 
    }, 700); 
  };

  const handleExcelImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const bstr = evt.target.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0]; 
        const ws = wb.Sheets[wsname];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
        
        let idxHeader = -1;
        let idxCodigo = -1, idxDescripcion = -1, idxExistencia = -1, idxPrecio = -1;
        let idxDepartamento = -1, idxRef = -1, idxSistema = -1;

        for (let i = 0; i < rows.length; i++) {
          const row = rows[i];
          if (!Array.isArray(row)) continue;
          const codigo = row.findIndex(c => String(c).toUpperCase().includes('CÓDIGO') || String(c).toUpperCase().includes('CODIGO'));
          const desc = row.findIndex(c => String(c).toUpperCase().includes('DESCRIPCIÓN') || String(c).toUpperCase().includes('DESCRIPCION'));
          
          if (codigo !== -1 && desc !== -1) {
            idxHeader = i;
            idxCodigo = codigo;
            idxDescripcion = desc;
            idxExistencia = row.findIndex(c => String(c).toUpperCase().includes('EXISTENCIA'));
            idxPrecio = row.findIndex(c => String(c).toUpperCase().includes('PRECIO UNIDAD'));
            idxDepartamento = row.findIndex(c => String(c).toUpperCase().includes('DEPARTAMENTO'));
            idxRef = row.findIndex(c => String(c).toUpperCase().includes('NUMERO DE REF'));
            idxSistema = row.findIndex(c => String(c).toUpperCase().includes('SISTEMA'));
            break;
          }
        }

        if (idxHeader === -1) {
          alert("No se pudo vincular la estructura del Excel.");
          return;
        }

        const inventarioProcesado = [];
        for (let i = idxHeader + 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length === 0) continue;

          const id = String(row[idxCodigo] || "").trim();
          const nombre = String(row[idxDescripcion] || "").trim();
          
          if (id && nombre) {
            inventarioProcesado.push({
              id,
              nombre,
              departamento: String(row[idxDepartamento] || "General").trim(),
              referencia: String(row[idxRef] || "N/A").trim(),
              cantidad: Number(row[idxExistencia] || 0),
              precio: Number(row[idxPrecio] || 0),
              sistema: String(row[idxSistema] || "PENDIENTE").trim(),
              ultimaActualizacion: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
            });
          }
        }

        await guardarEnFirebase(inventarioProcesado);
        alert(`¡Sincronización exitosa con Firebase! Se cargaron ${inventarioProcesado.length} artículos.`);
      } catch (error) {
        alert("Hubo un error al procesar el Excel.");
      }
    };
    reader.readAsBinaryString(file);
  };

  const modificarStock = (id, cambio) => {
    const nuevo = inventario.map(item => {
      if (item.id === id) {
        const nuevaCantidad = item.cantidad + cambio;
        return { 
          ...item, 
          cantidad: nuevaCantidad < 0 ? 0 : nuevaCantidad, 
          ultimaActualizacion: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) 
        };
      }
      return item;
    });
    guardarEnFirebase(nuevo);
  };

  const productosFiltrados = inventario.filter(item => 
    item.nombre.toLowerCase().includes(busqueda.toLowerCase()) || 
    item.id.toLowerCase().includes(busqueda.toLowerCase()) ||
    item.referencia.toLowerCase().includes(busqueda.toLowerCase())
  );

  return (
    <div style={{ fontFamily: "'Montserrat', sans-serif", backgroundColor: '#2D0208', minHeight: '100vh', color: '#ffffff' }}>
      
      {/* RESETEO GLOBAL PARA ELIMINAR LOS BORDES BLANCOS DE LA PÁGINA */}
      <style>{`
        html, body {
          margin: 0 !important;
          padding: 0 !important;
          background-color: #2D0208 !important;
          width: 100%;
          height: 100%;
        }

        @import url('https://fonts.googleapis.com/css2?family=Great+Vibes&family=Montserrat:wght@300;400;700;900&display=swap');

        .contenedor-bienvenida {
          transition: transform 0.7s cubic-bezier(0.76, 0, 0.24, 1), opacity 0.6s ease;
        }
        .contenedor-bienvenida.subir-efecto {
          transform: translateY(-100vh);
          opacity: 0;
        }
        
        .texto-oro-cursivo {
          font-family: 'Great Vibes', cursive;
          background: linear-gradient(to bottom, #FFF3CC 0%, #E5A93C 50%, #9E6B0F 100%);
          -webkit-background-clip: text;
          -webkit-text-fill-color: transparent;
          font-size: 95px;
          font-weight: 400;
          line-height: 0.9;
          margin: 0;
          padding: 0 20px; 
          text-shadow: 0px 4px 10px rgba(0,0,0,0.4);
        }
        
        .texto-oro-bloque {
          font-family: 'Montserrat', sans-serif;
          font-weight: 900;
          font-size: 105px;
          letter-spacing: -2px;
          line-height: 0.9;
          background: linear-gradient(to bottom, #FFFFFF 0%, #F5C667 30%, #D4931A 70%, #634102 100%);
          -webkit-background-clip: text;
          -webkit-text-fill-color: transparent;
          margin: 5px 0;
          filter: drop-shadow(0px 4px 6px rgba(0,0,0,0.7));
        }

        .boton-premium {
          background: transparent;
          border: 1px solid rgba(255, 255, 255, 0.25);
          color: #ffffff;
          padding: 12px 28px;
          border-radius: 4px;
          font-size: 11px;
          font-weight: 700;
          cursor: pointer;
          transition: all 0.35s ease;
          letter-spacing: 2px;
          text-transform: uppercase;
        }
        .boton-premium:hover {
          border-color: #F5C667 !important;
          color: #2D0208 !important;
          background: #F5C667 !important;
          box-shadow: 0 0 25px rgba(245, 198, 103, 0.4);
        }
        .input-oscuro {
          background: rgba(255, 255, 255, 0.04);
          border: 1px solid rgba(255, 255, 255, 0.1);
          color: #ffffff;
          padding: 14px 18px;
          border-radius: 4px;
          outline: none;
          transition: all 0.3s;
        }
        .input-oscuro:focus {
          border-color: #F5C667;
          background: rgba(255, 255, 255, 0.08);
        }
        .tabla-premium th {
          border-bottom: 1px solid rgba(255, 255, 255, 0.12) !important;
          color: #F5C667 !important;
          font-size: 11px !important;
          letter-spacing: 2px !important;
          text-transform: uppercase !important;
          padding: 18px 20px !important;
        }
        .btn-control {
          width: 28px;
          height: 28px;
          border-radius: 4px;
          border: 1px solid rgba(255,255,255,0.2);
          background: transparent;
          color: #ffffff;
          cursor: pointer;
          transition: all 0.2s;
        }
        .btn-control:hover {
          border-color: #F5C667;
          color: #F5C667;
        }
      `}</style>

      {/* INTERFAZ 1: ENTRADA */}
      {(fasePantalla === 'bienvenida' || fasePantalla === 'saliendo') && (
        <div 
          className={`contenedor-bienvenida ${fasePantalla === 'saliendo' ? 'subir-efecto' : ''}`}
          style={{ 
            height: '100vh', 
            background: 'radial-gradient(circle at center, #510711 0%, #160003 100%)', 
            display: 'flex', 
            justify: 'center', 
            alignItems: 'center', 
            position: 'relative',
            zIndex: 10,
            padding: '0 60px',
            boxSizing: 'border-box'
          }}
        >
          <div style={{ position: 'absolute', top: 0, left: 0, right: 0, padding: '40px 60px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span style={{ fontSize: '11px', fontWeight: '700', letterSpacing: '5px', color: 'rgba(255,255,255,0.3)' }}>SNC SYSTEM</span>
            <button onClick={iniciarSesionAnimado} className="boton-premium">
              Iniciar sesión
            </button>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '80px', width: '100%', maxWidth: '1200px', alignItems: 'center' }}>
            
            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', textAlign: 'center' }}>
              <h1 className="texto-oro-cursivo">Inventario</h1>
              <h2 className="texto-oro-bloque">SNC</h2>
              <p style={{ fontSize: '11px', letterSpacing: '4px', color: '#F5C667', fontWeight: '700', margin: '15px 0 0 0', textTransform: 'uppercase', opacity: 0.85 }}>
                Suministro Nacional de Cauchos
              </p>
            </div>

            <div style={{ display: 'flex', justifyContent: 'center' }}>
              <div style={{ 
                backgroundColor: '#FCEEA7', 
                width: '100%', 
                maxWidth: '420px', 
                height: '460px', 
                borderRadius: '210px 210px 0 0', 
                display: 'flex', 
                flexDirection: 'column', 
                justifyContent: 'center', 
                alignItems: 'center',
                padding: '40px',
                boxShadow: '0 15px 35px rgba(0,0,0,0.6)',
                textAlign: 'center',
                color: '#2D0208',
                boxSizing: 'border-box'
              }}>
                <span style={{ fontSize: '24px', fontWeight: '900', letterSpacing: '6px', marginBottom: '35px', textTransform: 'uppercase' }}>
                  ¡Bienvenidos!
                </span>
                <p style={{ fontSize: '15px', fontWeight: '700', lineHeight: '1.7', margin: 0, fontStyle: 'italic', padding: '0 10px', textTransform: 'uppercase' }}>
                  «Pon en las manos del Señor todas tus obras, y tus proyectos se cumplirán»
                </p>
              </div>
            </div>

          </div>

          <span style={{ position: 'absolute', bottom: '40px', fontSize: '10px', color: 'rgba(255,255,255,0.2)', letterSpacing: '2px', fontWeight: '700' }}>
            SNC, C.A. © 2026
          </span>
        </div>
      )}

      {/* INTERFAZ 2: PANEL DE CONTROL */}
      {fasePantalla === 'dashboard' && (
        <div style={{ padding: '60px 40px', backgroundColor: '#2D0208', minHeight: '100vh', boxSizing: 'border-box' }}>
          <div style={{ maxWidth: '1350px', margin: '0 auto' }}>
            
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', marginBottom: '45px', borderBottom: '1px solid rgba(255,255,255,0.08)', paddingBottom: '25px' }}>
              <div>
                <h2 style={{ fontSize: '26px', fontWeight: '300', margin: 0, letterSpacing: '-0.5px', color: 'rgba(255, 255, 255, 0.9)' }}>
                  Sistema <span style={{ fontWeight: '700', color: '#F5C667' }}>SNC 2026</span>
                </h2>
                <p style={{ color: 'rgba(255,255,255,0.4)', margin: '6px 0 0 0', fontSize: '13px' }}>Terminal de operaciones de Suministro Nacional de Cauchos.</p>
              </div>
              <button 
                onClick={() => setFasePantalla('bienvenida')}
                style={{ background: 'transparent', border: 'none', color: 'rgba(255,255,255,0.4)', cursor: 'pointer', fontSize: '11px', textTransform: 'uppercase', letterSpacing: '1px', textDecoration: 'underline' }}
              >
                Cerrar Sesión
              </button>
            </div>

            <div style={{ gridTemplateColumns: 'repeat(auto-fit, minmax(300px, 1fr))', display: 'grid', gap: '30px', marginBottom: '45px' }}>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                <span style={{ fontSize: '11px', color: '#F5C667', letterSpacing: '1.5px', textTransform: 'uppercase', fontWeight: '700' }}>Sincronizar Catálogo (Excel)</span>
                <input type="file" accept=".xlsx, .xls" onChange={handleExcelImport} className="input-oscuro" style={{ padding: '9px' }} />
              </div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                <span style={{ fontSize: '11px', color: '#F5C667', letterSpacing: '1.5px', textTransform: 'uppercase', fontWeight: '700' }}>Buscador Avanzado</span>
                <input type="text" placeholder="Buscar por descripción, código o referencia..." value={busqueda} onChange={(e) => setBusqueda(e.target.value)} className="input-oscuro" />
              </div>
            </div>

            <div style={{ background: 'rgba(0,0,0,0.2)', borderRadius: '6px', border: '1px solid rgba(255,255,255,0.06)', overflow: 'hidden' }}>
              <div style={{ overflowX: 'auto' }}>
                <table className="tabla-premium" style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left' }}>
                  <thead>
                    <tr>
                      <th>Código</th>
                      <th>N° Referencia</th>
                      <th>Descripción</th>
                      <th>Departamento</th>
                      <th style={{ textAlign: 'center' }}>Existencia</th>
                      <th>Precio Unidad</th>
                      <th>Sistema</th>
                      <th>Última Modificación</th>
                    </tr>
                  </thead>
                  <tbody>
                    {productosFiltrados.length === 0 ? (
                      <tr>
                        <td colSpan="8" style={{ padding: '55px', textAlign: 'center', color: 'rgba(255,255,255,0.35)', fontSize: '13px' }}>
                          {inventario.length === 0 ? "No hay registros vinculados. Importa tu archivo Excel para iniciar." : "No hay coincidencias para el término buscado."}
                        </td>
                      </tr>
                    ) : (
                      productosFiltrados.map((item, index) => (
                        <tr key={item.id + index} style={{ backgroundColor: index % 2 === 0 ? 'transparent' : 'rgba(255,255,255,0.02)' }}>
                          <td style={{ fontWeight: '700', color: '#ffffff' }}>{item.id}</td>
                          <td style={{ color: 'rgba(255,255,255,0.45)', fontFamily: 'monospace' }}>{item.referencia}</td>
                          <td style={{ color: 'rgba(255,255,255,0.85)', maxWidth: '280px', fontSize: '13px' }}>{item.nombre}</td>
                          <td>
                            <span style={{ border: '1px solid rgba(255,255,255,0.12)', color: 'rgba(255,255,255,0.5)', padding: '2px 8px', borderRadius: '4px', fontSize: '11px' }}>
                              {item.departamento}
                            </span>
                          </td>
                          <td>
                            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '10px' }}>
                              <button onClick={() => modificarStock(item.id, -1)} className="btn-control">-</button>
                              <span style={{ fontWeight: '600', minWidth: '22px', textAlign: 'center' }}>{item.cantidad}</span>
                              <button onClick={() => modificarStock(item.id, 1)} className="btn-control">+</button>
                            </div>
                          </td>
                          <td style={{ fontWeight: '600', color: '#34d399' }}>${item.precio.toFixed(2)}</td>
                          <td>
                            <span style={{ 
                              border: item.sistema.toUpperCase() === 'INGRESADO' ? '1px solid rgba(52, 211, 153, 0.3)' : '1px solid rgba(251, 191, 36, 0.3)', 
                              color: item.sistema.toUpperCase() === 'INGRESADO' ? '#34d399' : '#fbbf24', 
                              padding: '2px 7px', borderRadius: '4px', fontSize: '11px'
                            }}>
                              {item.sistema}
                            </span>
                          </td>
                          <td style={{ fontSize: '12px', color: 'rgba(255,255,255,0.3)' }}>{item.ultimaActualizacion}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>

          </div>
        </div>
      )}
    </div>
  );
}

export default App;