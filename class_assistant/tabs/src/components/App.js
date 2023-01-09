// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react';
import './App.css';
import * as microsoftTeams from '@microsoft/teams-js';
import { HashRouter as Router, Route,Link} from 'react-router-dom';
import Privacy from './about/Privacy';
import TermsOfUse from './about/TermsOfUse';
import Tab from './Tab';
import Tab_s from './Tab_s';
import TabConfig from './TabConfig';
import announce from './announcement';
import announce_s from './announcement_s';
import Material from './material';
import Material_s from './material_s';
import homework from './homework';
import homework_s from './homework_s';
import { Menu,Button} from "@fluentui/react-northstar"
import Profile from "./Profile";
import { TeamsFx } from "@microsoft/teamsfx";

class App extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      userInfo: {},
      items: [],
      table: [],
      newItemDescription: "",
      showLoginPage: true,
      isAddingItem: false,
      initialized: false,
      username:"",
      role:"",
    }
  }
  
  async componentDidMount() {
    await this.initTeamsFx();
    //await this.checkIsConsentNeeded();
  }

  async loginBtnClick(role) {
    try {
      // Popup login page to get user's access token
      await this.teamsfx.login(this.scope);
      
    } catch (err) {
      if (err instanceof Error && err.message?.includes("CancelledByUser")) {
        const helpLink = "https://aka.ms/teamsfx-auth-code-flow";
        err.message +=
          "\nIf you see \"AADSTS50011: The reply URL specified in the request does not match the reply URLs configured for the application\" " +
          "in the popup window, you may be using unmatched version for TeamsFx SDK (version >= 0.5.0) and Teams Toolkit (version < 3.3.0) or " +
          `cli (version < 0.11.0). Please refer to the help link for how to fix the issue: ${helpLink}` ;
      }

      alert("Login failed: " + err);
      return;
    }
    this.setState({
      role:role,
      showLoginPage:false
    })
  }

 async initTeamsFx() {
  const teamsfx = new TeamsFx();
  // Get the user info from access token
  const userInfo = await teamsfx.getUserInfo();
  const username=userInfo.displayName;
  this.setState({
    username:username
  })
  this.setState({
    userInfo: userInfo
  });
  this.teamsfx = teamsfx;
  this.scope = ["User.Read", "User.ReadBasic.All"];
  this.channelOrChatId = await this.getChannelOrChatId();
  }

  async getChannelOrChatId() {
    return new Promise((resolve) => {
      microsoftTeams.getContext((context) => {
        if (context.channelId) {
          resolve(context.channelId);
        } else if (context.chatId) {
          resolve(context.chatId);
        } else {
          resolve(this.state.userInfo.objectId);
        }
      });
    });
  }

  async checkIsConsentNeeded() {
    try {
      await this.teamsfx.getCredential().getToken(this.scope);
    } catch (error) {
      this.setState({
        showLoginPage: true
      });
      return true;
    }
    this.setState({
      showLoginPage: false
    });
    return false;
  }

  render(){
    return(
      <div >
      {this.state.showLoginPage === true&&<div className="auth">
        <Profile userInfo={this.state.userInfo} />
        <h2>Welcome to Sign In App!</h2>
        <Button primary onClick={() => this.loginBtnClick("teacher")}>Teacher Start</Button>
        <Button primary onClick={() => this.loginBtnClick("student")}>Student Start</Button>
      </div>}
      {this.state.role=="teacher"&&this.state.showLoginPage===false&&<Router>
      <Menu className='menu'
        items={[
          <div className='signin'>
            <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4"></path></svg>
            <Link to='/tab'>签到</Link>
          </div>,
          <div className='announcement'>
            <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-3 7h3m-3 4h3m-6-4h.01M9 16h.01"></path></svg>            
            <Link to='/announcement'>公告</Link>
          </div>,
          <div className='homework'> 
            <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"></path></svg>
            <Link to='/homework'>作业</Link>
          </div>,
          <div className='material'>
            <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 21h10a2 2 0 002-2V9.414a1 1 0 00-.293-.707l-5.414-5.414A1 1 0 0012.586 3H7a2 2 0 00-2 2v14a2 2 0 002 2z"></path></svg>
            <Link to='/material'>资料</Link>
          </div>,
        ]}
        primary
      />
      {/* <div>welcome {this.state.username}!</div> */}
      <div className='content'>
        <Route exact path='/privacy' component={Privacy} />
        <Route exact path='/termsofuse' component={TermsOfUse} />
        <Route exact path='/tab' component={Tab} />
        <Route exact path='/config' component={TabConfig} />
        <Route exact path='/announcement' component={announce} />
        <Route exact path='/homework' component={homework} />
        <Route exact path='/material' component={Material} />
      </div>
      </Router>}

      {this.state.role=="student"&&this.state.showLoginPage===false&&<Router>
      <Menu className='menu'
        items={[
          <div className='signin'>
            <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4"></path></svg>
            <Link to='/tab_s'>签到</Link>
          </div>,
          <div className='announcement'>
            <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-3 7h3m-3 4h3m-6-4h.01M9 16h.01"></path></svg>            
            <Link to='/announcement_s'>公告</Link>
          </div>,
          <div className='homework'> 
            <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"></path></svg>
            <Link to='/homework_s'>作业</Link>
          </div>,
          <div className='material'>
            <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 21h10a2 2 0 002-2V9.414a1 1 0 00-.293-.707l-5.414-5.414A1 1 0 0012.586 3H7a2 2 0 00-2 2v14a2 2 0 002 2z"></path></svg>
            <Link to='/material_s'>资料</Link>
          </div>,
        ]}
        primary
      />
      
      <div className='content'>
        <Route exact path='/privacy' component={Privacy} />
        <Route exact path='/termsofuse' component={TermsOfUse} />

        <Route exact path='/tab_s' component={Tab_s} />
        <Route exact path='/config' component={TabConfig} />

        <Route exact path='/announcement_s' component={announce_s} />

        <Route exact path='/homework_s' component={homework_s} />

        <Route exact path='/material_s' component={Material_s} />

      </div>
      </Router>}
      
      </div>
    );
  }
  
}

function main() {
  // Check for the Microsoft Teams SDK object.
  if (microsoftTeams) {
    return (
      
      <div >
      <div className="auth">
        <Profile userInfo={this.state.userInfo} />
        <h2>Welcome to Sign In App!</h2>
        <Button primary onClick={() => this.loginBtnClick("teacher")}>Teacher Start</Button>
        <Button primary onClick={() => this.loginBtnClick("student")}>Student Start</Button>
      </div>
      {role=="teacher"&&<Router>
      <Menu className='menu'
        items={[
          <div className='signin'>
            <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4"></path></svg>
            <Link to='/tab'>签到</Link>
          </div>,
          <div className='announcement'>
            <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-3 7h3m-3 4h3m-6-4h.01M9 16h.01"></path></svg>            
            <Link to='/announcement'>公告</Link>
          </div>,
          <div className='homework'> 
            <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"></path></svg>
            <Link to='/homework'>作业</Link>
          </div>,
          <div className='material'>
            <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 21h10a2 2 0 002-2V9.414a1 1 0 00-.293-.707l-5.414-5.414A1 1 0 0012.586 3H7a2 2 0 00-2 2v14a2 2 0 002 2z"></path></svg>
            <Link to='/material'>资料</Link>
          </div>,
        ]}
        primary
      />
      {role=="student"&&<Menu className='menu'
        items={[
          <div className='signin'>
            <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4"></path></svg>
            <Link to='/tab'>签到</Link>
          </div>,
          <div className='announcement'>
            <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-3 7h3m-3 4h3m-6-4h.01M9 16h.01"></path></svg>            
            <Link to='/announcement'>公告</Link>
          </div>,
          <div className='homework'> 
            <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"></path></svg>
            <Link to='/homework'>作业</Link>
          </div>,
          <div className='material'>
            <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 21h10a2 2 0 002-2V9.414a1 1 0 00-.293-.707l-5.414-5.414A1 1 0 0012.586 3H7a2 2 0 00-2 2v14a2 2 0 002 2z"></path></svg>
            <Link to='/material'>资料</Link>
          </div>,
        ]}
        primary
      />}
      <div className='content'>
        <Route exact path='/privacy' component={Privacy} />
        <Route exact path='/termsofuse' component={TermsOfUse} />
        <Route exact path='/tab' component={Tab} />
        <Route exact path='/config' component={TabConfig} />
        <Route exact path='/announcement' component={announce} />
        <Route exact path='/homework' component={homework} />
        <Route exact path='/material' component={Material} />
      </div>
      </Router>}
      
      </div>
    );
  } else {
    return (
      <h3>Microsoft Teams SDK not found.</h3>
    );
  }
}

export default App;
