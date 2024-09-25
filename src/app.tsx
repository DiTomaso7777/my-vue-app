import './app.css';
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from '@azure/msal-react';
import { loginRequest } from './auth/auth-config';
import { ContactsComponent } from './ContactsComponent';
import { SendMailComponent } from './SendMailComponent'; // Import SendMailComponent
import { SendMailOnBehalfComponent } from './SendMailOnBehalfComponent'; // Import SendMailOnBehalfComponent
import { BrowserRouter as Router, Route, Routes, Link } from 'react-router-dom'; // Import React Router components

function App() {
  const { instance } = useMsal();
  const activeAccount = instance.getActiveAccount();

  const handleLoginRedirect = () => {
    instance
      .loginRedirect({
        ...loginRequest,
        prompt: 'create',
      })
      .catch((error) => console.log(error));
  };

  const handleLogoutRedirect = () => {
    instance.logoutRedirect({
      postLogoutRedirectUri: '/',
    });
    window.location.reload();
  };

  const title = activeAccount ? activeAccount.name : 'My App';

  return (
    <Router>
      <div className="card">
        <AuthenticatedTemplate>
          {activeAccount ? (
            <>
              <button onClick={handleLogoutRedirect}>Logout</button>
              <p>You are signed in as {title}</p>
              {/* Navigation Links */}
              <nav>
                <Link to="/contacts">View Contacts</Link> |{' '}
                <Link to="/send-mail">Send New Mail</Link> |{' '}
                <Link to="/send-mail-on-behalf">Send Mail on Behalf</Link>
              </nav>

              {/* Define Routes */}
              <Routes>
                <Route path="/contacts" element={<ContactsComponent />} />
                <Route path="/send-mail" element={<SendMailComponent />} />
                <Route path="/send-mail-on-behalf" element={<SendMailOnBehalfComponent />} />
              </Routes>
            </>
          ) : null}
        </AuthenticatedTemplate>
        <UnauthenticatedTemplate>
          <>
            <button onClick={handleLoginRedirect}>Login</button>
            <p>Please sign in!</p>
          </>
        </UnauthenticatedTemplate>
      </div>
    </Router>
  );
}

export default App;
