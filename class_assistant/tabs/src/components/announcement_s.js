// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react';
import axios from 'axios';
import { TeamsFx } from "@microsoft/teamsfx";
import './announcement.css'
import { Checkbox, Button, Input, MenuButton,Text,Form ,FormInput, FormCheckbox, FormButton} from "@fluentui/react-northstar"


class announcement_s extends React.Component {

  constructor(props) {
    super(props);
    this.state = {
      userInfo: {},
      items: [],
      table: [],
      newItemDescription: "",
      newtitle: "",
      showLoginPage: false,
      isAddingItem: false,
      initialized: false,
      username:"",
      text:"",
      title:"",
    }
  }

  async componentDidMount() {
    await this.initTeamsFx();
    await this.initData();
  }

  async initData() {

      await this.getItems();
    
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
  }

  async callFunctionWithErrorHandling(command, method, options, params) {
    var message = [];
    var funcErrorMsg = "";
    try {
      // Get SSO token for the user
      const accessToken = await this.teamsfx.getCredential().getToken("");
      // Gets configuration for API
      const apiEndpoint = this.teamsfx.getConfig("apiEndpoint");
      const response = await axios.default.request({
        method: method,
        url: apiEndpoint + "/api/" + command,
        headers: {
          authorization: "Bearer " + accessToken.token
        },
        data: options,
        params
      });
      message = response.data;
    } catch (err) {
      if (err.response && err.response.status && err.response.status === 404) {
        funcErrorMsg =
          'There may be a problem with the deployment of Azure Function App, please deploy Azure Function (Run command palette "TeamsFx - Deploy Package") first before running this App';
      } else if (err.message === "Network Error") {
        funcErrorMsg =
          "Cannot call Azure Function due to network error, please check your network connection status and ";
        if (err.config.url.indexOf("localhost") >= 0) {
          funcErrorMsg +=
            'make sure to start Azure Function locally (Run "npm run start" command inside api folder from terminal) first before running this App';
        } else {
          funcErrorMsg +=
            'make sure to provision and deploy Azure Function (Run command palette "TeamsFx - Provision Resource" and "TeamsFx - Deploy Package") first before running this App';
        }
      } else {
        funcErrorMsg = err.toString();
        if (err.response?.data?.error) {
          funcErrorMsg += ": " + err.response.data.error;
        }
        alert(funcErrorMsg);
      }
    }
    return message;
  }

  async getItems() {   //刷新页面
    // Use client TeamsFx SDK to call "todo" Azure Function in "get" method to get all todo list which belong to user oid
    let result = await this.callFunctionWithErrorHandling("todo", "get", undefined, { flag:'announcement' });
    if ("Error" === result) {
      throw new Error("todo Function failed, please check Azure Functions log for details!");
    } else {
      this.setState({
        items: result,
        initialized: true,
      });
    }
  }

  async onAddItem(title,text) {      
    const newItems = JSON.parse(JSON.stringify(this.state.items));
    newItems.push({
      description: this.state.newItemDescription,
      text:this.state.newtitle,
    })
    
    this.setState({
      newItemDescription: "",
      newtitle:"",
      items: newItems
    });
    console.log(title)
    // Use client TeamsFx SDK to call "todo" Azure Function in "post" method to insert a new todo item under user oid
    await this.callFunctionWithErrorHandling("todo", "post", {title:title,text:text,flag:4}  //description给成time
    );
    this.refresh();
  }

  async onDelete(id,title){
    const newItems = this.state.items.filter(item => item.id !== id);
    this.setState({
      items: newItems
    })
    await this.callFunctionWithErrorHandling("todo", "delete", { title, flag:1 });
  }

  async refresh() {
    await this.getItems();
  }


  render() {

    const items = this.state.items?.map((item, index) =>
      <div key={index} className="item">
        <div className="title">
          <Text         //输出
            content={this.state.items[index].title}
          />
        </div>
        <div className="text">
          <Text         //输出
            content={this.state.items[index].text}
          />
        </div>
        
      </div>
    );

    return (
      <div>
        {this.state.showLoginPage === false && <div className="flex-container">
          <div className="todo-col">
            <div className="todo">
              <div className="header">
                <div className="title">
                  <h2>Announcement List</h2>
                  <span>{this.state.items.length} Announcement{this.state.items.length === 1 ? "" : "s"}</span>
                </div>

                
              </div>  
              {items}
              {this.state.isAddingItem && <div className="item add">
                <Form
                  onSubmit={() => {
                    const tit = document.getElementById("title")
                    const tex = document.getElementById("text")
                    
                    this.setState({newtitle:tit.value,newItemDescription:tex.value,title:tit.value,text:tex.value})
                    this.onAddItem(tit.value,tex.value)
                  }}
                >
                <FormInput
                  label="title"
                  name="title"
                  id="title"
                  required
                  showSuccessIndicator={false}
                />
                <FormInput
                  label="text"
                  name="text"
                  id="text"
                  required
                  showSuccessIndicator={false}
                />
                <FormButton content="Submit" />
                </Form>
              </div>}
            </div>
          </div>
        </div>}
      </div>
    );
    }
}
export default announcement_s;
