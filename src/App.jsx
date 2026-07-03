import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { initializeApp } from "firebase/app";
import { getFirestore, doc, getDoc, setDoc } from "firebase/firestore";

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

  useEffect(() => {
    const cargarDatos = async () => {
      try {
        const docRef = doc(db, "datos", DOC_ID);
        const docSnap = await getDoc(docRef);
        if (docSnap.exists()) {
          setInventario(docSnap.data().items || []);
        }
      } catch (error) { console.error("Error:", error); }
    };
    cargarDatos();
  }, []);

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

  return (
    <div style={{ fontFamily: "'Montserrat', sans-serif", backgroundColor: '#2D0208', minHeight: '100vh', color: '#ffffff' }}>
      {(fasePantalla === 'bienvenida' || fasePantalla === 'saliendo') && (
        <div className={`contenedor-bienvenida ${fasePantalla === 'saliendo' ? 'subir-efecto' : ''}`} style={{ height: '100vh', background: 'radial-gradient(circle at center, #510711 0%, #160003 100%)', display: 'flex', justifyContent: 'center', alignItems: 'center' }}>
            <button onClick={iniciarSesionAnimado} style={{ padding: '15px 30px', cursor: 'pointer', background: '#F5C667', border: 'none', fontWeight: 'bold' }}>Iniciar sesión</button>
        </div>
      )}

      {fasePantalla === 'dashboard' && (
        <div style={{ padding: '60px 40px' }}>
          <h2>Sistema SNC 2026</h2>
          <input type="file" accept=".xlsx" onChange={handleExcelImport} />
          <input type="text" placeholder="Buscar..." value={busqueda} onChange={(e) => setBusqueda(e.target.value)} />
          <table style={{ width: '100%', marginTop: '20px', borderCollapse: 'collapse' }}>
            <thead>
              <tr style={{ color: '#F5C667', borderBottom: '1px solid #555' }}>
                <th>Código</th><th>Nombre</th><th>Cantidad</th><th>Acciones</th>
              </tr>
            </thead>
            <tbody>
              {productosFiltrados.map((item, index) => (
                <tr key={index} style={{ borderBottom: '1px solid #333' }}>
                  <td>{item.id}</td><td>{item.nombre}</td><td>{item.cantidad}</td>
                  <td>
                    <button onClick={() => modificarStock(item.id, -1)}>-</button>
                    <button onClick={() => modificarStock(item.id, 1)}>+</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

export default App;