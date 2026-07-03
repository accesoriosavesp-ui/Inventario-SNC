import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { initializeApp } from "firebase/app";
import { getFirestore, doc, getDoc, setDoc } from "firebase/firestore";

// Configuración de tu proyecto
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

  // --- CARGAR DATOS DE LA NUBE ---
  useEffect(() => {
    const cargarDatos = async () => {
      try {
        const docRef = doc(db, "datos", DOC_ID);
        const docSnap = await getDoc(docRef);
        if (docSnap.exists()) {
          setInventario(docSnap.data().items || []);
        }
      } catch (error) {
        console.error("Error al cargar:", error);
      }
    };
    cargarDatos();
  }, []);

  // --- GUARDAR EN LA NUBE ---
  const syncFirebase = async (nuevoInventario) => {
    setInventario(nuevoInventario);
    await setDoc(doc(db, "datos", DOC_ID), { items: nuevoInventario });
  };

  const iniciarSesionAnimado = () => {
    setFasePantalla('saliendo'); 
    setTimeout(() => { setFasePantalla('dashboard'); }, 700); 
  };

  const handleExcelImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
      
      // Lógica de procesamiento original
      let idxHeader = rows.findIndex(r => r.some(c => String(c).toUpperCase().includes('CÓDIGO')));
      if (idxHeader === -1) return alert("Estructura no encontrada");

      const inventarioProcesado = [];
      for (let i = idxHeader + 1; i < rows.length; i++) {
        if (!rows[i][0]) continue;
        inventarioProcesado.push({
          id: String(rows[i][0]),
          nombre: String(rows[i][1] || ""),
          departamento: String(rows[i][2] || "General"),
          referencia: String(rows[i][3] || "N/A"),
          cantidad: Number(rows[i][4] || 0),
          precio: Number(rows[i][5] || 0),
          sistema: String(rows[i][6] || "PENDIENTE"),
          ultimaActualizacion: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
        });
      }
      syncFirebase(inventarioProcesado);
    };
    reader.readAsBinaryString(file);
  };

  const modificarStock = (id, cambio) => {
    const nuevo = inventario.map(item => {
      if (item.id === id) {
        return { ...item, cantidad: Math.max(0, item.cantidad + cambio), ultimaActualizacion: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) };
      }
      return item;
    });
    syncFirebase(nuevo);
  };

  const productosFiltrados = inventario.filter(item => 
    item.nombre.toLowerCase().includes(busqueda.toLowerCase()) || 
    item.id.toLowerCase().includes(busqueda.toLowerCase())
  );

  // --- INTERFAZ (Mantenida igual) ---
  return (
    <div style={{ fontFamily: "'Montserrat', sans-serif", backgroundColor: '#2D0208', minHeight: '100vh', color: '#ffffff' }}>
      <style>{`
        /* Tus estilos CSS van aquí igual que antes */
      `}</style>

      {(fasePantalla === 'bienvenida' || fasePantalla === 'saliendo') && (
        <div className={`contenedor-bienvenida ${fasePantalla === 'saliendo' ? 'subir-efecto' : ''}`} style={{ height: '100vh', background: 'radial-gradient(circle at center, #510711 0%, #160003 100%)', display: 'flex', justifyContent: 'center', alignItems: 'center', padding: '0 60px', boxSizing: 'border-box' }}>
            <div style={{ position: 'absolute', top: 0, left: 0, right: 0, padding: '40px 60px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span style={{ fontSize: '11px', fontWeight: '700', letterSpacing: '5px', color: 'rgba(255,255,255,0.3)' }}>SNC SYSTEM</span>
                <button onClick={iniciarSesionAnimado} className="boton-premium">Iniciar sesión</button>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '80px', width: '100%', maxWidth: '1200px', alignItems: 'center' }}>
                <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', textAlign: 'center' }}>
                    <h1 className="texto-oro-cursivo">Inventario</h1>
                    <h2 className="texto-oro-bloque">SNC</h2>
                </div>
            </div>
        </div>
      )}

      {fasePantalla === 'dashboard' && (
        <div style={{ padding: '60px 40px', backgroundColor: '#2D0208', minHeight: '100vh', boxSizing: 'border-box' }}>
            <div style={{ maxWidth: '1350px', margin: '0 auto' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '45px' }}>
                    <h2>Sistema <span style={{ color: '#F5C667' }}>SNC 2026</span></h2>
                    <button onClick={() => setFasePantalla('bienvenida')}>Cerrar Sesión</button>
                </div>
                <input type="file" accept=".xlsx" onChange={handleExcelImport} />
                <input type="text" placeholder="Buscar..." value={busqueda} onChange={(e) => setBusqueda(e.target.value)} />
                {/* Tabla de productos igual que antes, usando {productosFiltrados.map(...)} */}
            </div>
        </div>
      )}
    </div>
  );
}

export default App;