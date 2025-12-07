// src/layouts/AppLayout.jsx
import React from 'react';
import PropTypes from 'prop-types';

export default function AppLayout({ children }) {
  return (
    <div className="app-layout">
      {/* Global Header */}
      <header className="navbar">
        <div className="brand">Excel Access Portal</div>
        <div className="row">
          {/* Inject buttons, theme toggle, or user menu here */}
        </div>
      </header>

      {/* Page Content */}
      <main className="main-content" role="main">
        {children}
      </main>

      {/* Global Footer */}
      <footer className="footer">
        <p>&copy; {new Date().getFullYear()} Excel Access Portal — All rights reserved.</p>
      </footer>
    </div>
  );
}

AppLayout.propTypes = {
  children: PropTypes.node.isRequired,
};
