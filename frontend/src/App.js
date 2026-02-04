import React from 'react';
import { BrowserRouter as Router, Routes, Route, Link, NavLink, Navigate } from 'react-router-dom';
import { AuthProvider, useAuth } from './contexts/AuthContext';
import ProtectedRoute from './components/ProtectedRoute';
import Home from './pages/Home';
import Login from './pages/Login';
import Register from './pages/Register';
import './App.css';

function Navigation() {
  const { isAuthenticated, user, logout } = useAuth();

  return (
    <header className="App-header">
      <div className="header-content">
        <Link to="/" className="brand" aria-label="Go to Home">
          <div className="brand-mark">
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z" />
              <polyline points="14 2 14 8 20 8" />
              <path d="M12 13l2 2 4-4" stroke="url(#gradient-check)" strokeWidth="2.5" />
              <defs>
                <linearGradient id="gradient-check" x1="12" y1="13" x2="18" y2="11" gradientUnits="userSpaceOnUse">
                  <stop stopColor="#0ea5e9" />
                  <stop offset="1" stopColor="#6366f1" />
                </linearGradient>
              </defs>
            </svg>
          </div>
          <div className="brand-text">
            <div className="brand-title">Resume<span className="text-highlight">AI</span></div>
            <div className="brand-subtitle">Smart Shortlisting</div>
          </div>
        </Link>
        <nav>
          {isAuthenticated && (
            <NavLink to="/" className={({ isActive }) => `nav-link${isActive ? ' active' : ''}`}>Home</NavLink>
          )}
          {!isAuthenticated ? (
            <>
              <NavLink to="/login" className={({ isActive }) => `nav-link${isActive ? ' active' : ''}`}>Login</NavLink>
              <NavLink to="/register" className={({ isActive }) => `nav-link${isActive ? ' active' : ''}`}>Register</NavLink>
            </>
          ) : (
            <div className="user-profile-group">
              <div className="user-avatar">
                {user?.name?.charAt(0).toUpperCase()}
              </div>
              <div className="user-text">
                <span className="user-name">{user?.name}</span>
                <span className="user-role">{user?.role}</span>
              </div>
              <div className="divider-vertical"></div>
              <button onClick={logout} className="logout-btn-mini" title="Logout">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"></path>
                  <polyline points="16 17 21 12 16 7"></polyline>
                  <line x1="21" y1="12" x2="9" y2="12"></line>
                </svg>
              </button>
            </div>
          )}
        </nav>
      </div>
    </header>
  );
}

function App() {
  return (
    <AuthProvider>
      <Router>
        <div className="App">
          <Navigation />
          
          <main className="App-main">
            <Routes>
              <Route path="/login" element={<Login />} />
              <Route path="/register" element={<Register />} />
              
              <Route 
                path="/" 
                element={
                  <ProtectedRoute>
                    <Home />
                  </ProtectedRoute>
                } 
              />
              
              <Route 
                path="/admin/dashboard" 
                element={
                  <ProtectedRoute allowedRoles={['admin']}>
                    <Home />
                  </ProtectedRoute>
                } 
              />
              
              <Route 
                path="/user/dashboard" 
                element={
                  <ProtectedRoute allowedRoles={['user']}>
                    <Home />
                  </ProtectedRoute>
                } 
              />

              <Route path="/unauthorized" element={
                <div style={{ textAlign: 'center', padding: '50px' }}>
                  <h2>Unauthorized Access</h2>
                  <p>You don't have permission to access this page.</p>
                  <Link to="/">Go Home</Link>
                </div>
              } />

              <Route path="*" element={<Navigate to="/" replace />} />
            </Routes>
          </main>
          
          <footer className="App-footer">
            <p>&copy; 2026 AI-Powered Resume Shortlisting System</p>
          </footer>
        </div>
      </Router>
    </AuthProvider>
  );
}

export default App;
