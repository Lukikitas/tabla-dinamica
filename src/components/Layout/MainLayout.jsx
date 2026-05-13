const navItems = [
  { id: 'dashboard', label: 'Dashboard' },
  { id: 'importar', label: 'Importar Producto Hora' },
  { id: 'configuracion', label: 'Configuracion' },
];

function MainLayout({ children, currentPage, onNavigate }) {
  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <span className="brand-mark">PH</span>
          <div>
            <strong>Producto Hora</strong>
            <small>Migracion Excel VBA</small>
          </div>
        </div>

        <nav className="nav-list" aria-label="Navegacion principal">
          {navItems.map((item) => (
            <button
              key={item.id}
              className={item.id === currentPage ? 'nav-item active' : 'nav-item'}
              type="button"
              onClick={() => onNavigate(item.id)}
            >
              {item.label}
            </button>
          ))}
        </nav>
      </aside>

      <main className="main-content">{children}</main>
    </div>
  );
}

export default MainLayout;
