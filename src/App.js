import React, { Component } from 'react';
import './App.css';
import { Text, Button } from 'office-ui-fabric-react';
import { GraphFileBrowser } from '@microsoft/file-browser';
import { UserAgentApplication } from 'msal';

const scopes = ['files.readwrite.all', 'user.read'];

class App extends Component {
  constructor(props) {
    super(props);
    this.state = {
      token: null
    };
    this.msal = new UserAgentApplication(
        'f054b515-d62a-44dd-b760-ea4ec9c24c65',
        '',
        () => {}
    );
  }

  componentDidMount() {
    this._getMsalToken();
  }

  render() {
    return (
      <div className="App">
        <Text variant="xxLarge">File Browser Demo</Text>
        {
          this.state.token ?
            <GraphFileBrowser
              getAuthenticationToken={ () => { return this.state.token; } }
              onSuccess={alert}
              onRenderCancelButton={()=>null}
              onRenderSuccessButton={()=>null}
            />
            :
            <Text block>logging in...</Text>
        }
      </div>
    );
  }

  _getMsalToken() {
    debugger;
    let user = this.msal.getUser();
    if (user) {
        console.log("got user");
        this.msal.acquireTokenSilent(scopes).then(token => {
          console.log("silent token");
          this.setState({ token: token });
        }, err => {
            console.log("acquireTokenRedirect");
            this.msal.acquireTokenRedirect(scopes);
        });
    } else {
      console.log("loginpopup");
      //this.msal.loginRedirect(scopes);
      this.msal.loginPopup(scopes).then(token => {
        this.setState({ token: token });
      });
    }
  }
}

export default App;
