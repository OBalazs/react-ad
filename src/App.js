import logo from './logo.svg';
import './App.css';
import { config } from './Config';
import { PublicClientApplication } from '@azure/msal-browser';
import React from 'react';

const msalInstance = new PublicClientApplication({
  auth: {
    clientId: config.appId,
    authority: config.authority,
    redirectUri: config.redirectUri,
  },
  cache: {
    cacheLocation: 'localStorage',
    storeAuthStateInCookie: false,
  },
});

class App extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      isAuthenticated: false,
      userName: '',
      msalInitialized: false, // Track initialization status
    };
    this.login = this.login.bind(this);
    this.logout = this.logout.bind(this);
  }

  async componentDidMount() {
    try {
      await msalInstance.initialize();
      this.setState({ msalInitialized: true });
      console.log("MSAL initialized successfully");
    } catch (error) {
      console.error("Failed to initialize MSAL:", error);
    }
  }

  async login() {
    if (!this.state.msalInitialized) {
      console.error("MSAL is not initialized yet.");
      return;
    }

    try {
      console.log("Login button clicked");
      const loginResponse = await msalInstance.loginPopup({
        scopes: ["User.Read", ...config.scopes], // Ensure User.Read is included
        prompt: "select_account",
      });
      console.log("Login successful:", loginResponse);

      // Fetch user profile information
      const accessToken = loginResponse.accessToken;
      const userProfile = await fetch("https://graph.microsoft.com/v1.0/me", {
        method: "GET",
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      }).then((response) => response.json());

      this.setState({
        isAuthenticated: true,
        userName: `${userProfile.displayName} (${userProfile.mail || userProfile.userPrincipalName})`, // Full name with email
      });
    } catch (error) {
      console.error("Login failed:", error);
      this.setState({
        isAuthenticated: false,
        userName: "",
      });
    }
  }

  logout() {
    if (!this.state.msalInitialized) {
      console.error("MSAL is not initialized yet.");
      return;
    }

    msalInstance.logoutPopup();
    this.setState({
      isAuthenticated: false,
      userName: '',
    });
  }

  render() {
    return (
      <div className="App">
        <nav className="App-nav">
              {!this.state.isAuthenticated ? (
                <button onClick={this.login}>Login</button>
              ) : (
                <button onClick={this.logout}>Logout</button>
              )}
          <div className="App-user-info">
            {this.state.isAuthenticated && `Welcome, ${this.state.userName}`}
          </div>
        </nav>
        <header className="App-header">
          <img src={logo} className="App-logo" alt="logo" />
          <p>
            {this.state.isAuthenticated ? 'You are logged in!' : 'Please log in'}
          </p>
        </header>
      </div>
    );
  }
}

export default App;