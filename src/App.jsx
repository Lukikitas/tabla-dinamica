import { useState } from 'react';
import MainLayout from './components/Layout/MainLayout.jsx';
import Dashboard from './pages/Dashboard.jsx';
import ImportarProductoHora from './pages/ImportarProductoHora.jsx';
import Configuracion from './pages/Configuracion.jsx';

const pages = {
  dashboard: Dashboard,
  importar: ImportarProductoHora,
  configuracion: Configuracion,
};

function App() {
  const [currentPage, setCurrentPage] = useState('importar');
  const Page = pages[currentPage];

  return (
    <MainLayout currentPage={currentPage} onNavigate={setCurrentPage}>
      <Page />
    </MainLayout>
  );
}

export default App;
