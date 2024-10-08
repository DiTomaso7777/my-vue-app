import React from 'react';
import ReactDOM from 'react-dom/client';

import './index.css'
import { useAuthProvider } from './auth/hooks/auth-provider.tsx'
import App from './app.tsx';

const { AuthProvider } = useAuthProvider();

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <AuthProvider> 
      <App />
    </AuthProvider>
  </React.StrictMode>,
)
