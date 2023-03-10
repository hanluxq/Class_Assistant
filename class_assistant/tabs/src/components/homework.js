// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react';
import axios from 'axios';
import { TeamsFx } from "@microsoft/teamsfx";
import { Checkbox, Button, Input, MenuButton,Text,Form ,FormInput, FormCheckbox, FormButton} from "@fluentui/react-northstar"

class Homework extends React.Component {

  constructor(props) {
    super(props);
    this.state = {
      userInfo: {},
      items: [],
      table: [],
      newItemDescription: "",
      nowtname:"",
      newAnswer:"",
      newtitle: "",
      showLoginPage: false,
      isAddingItem: false,
      isAddingAnswer: false,
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

  async getItems() {   //????????????
    // Use client TeamsFx SDK to call "todo" Azure Function in "get" method to get all todo list which belong to user oid
    let result = await this.callFunctionWithErrorHandling("todo", "get", undefined, { flag:'homework' });
    if ("Error" === result) {
      throw new Error("todo Function failed, please check Azure Functions log for details!");
    } else {
      this.setState({
        items: result,
        initialized: true,
      });
    }
  }

  async onAddItem(title,text) {       //????????????
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
    // Use client TeamsFx SDK to call "todo" Azure Function in "post" method to insert a new todo item under user oid
    await this.callFunctionWithErrorHandling("todo", "post", {title:title,text:text, flag:2}  
    );
    this.refresh();
  }

  async onDeleteItem(id,tname) {   //????????????
    const newItems = this.state.items.filter(item => item.id !== id);
    this.setState({
      items: newItems
    })
    console.log('delete table :'+id+','+tname)
    // Use client TeamsFx SDK to call "todo" Azure Function in "delete" method to delete a todo item
    await this.callFunctionWithErrorHandling("todo", "delete", { tname:tname, flag:2 });
  }

  async onSub(username) {   //???????????????  ???description(??????),userid
    console.log('tname:'+this.state.nowtname+',user:'+username+',answer:'+this.state.newAnswer)
    await this.callFunctionWithErrorHandling("todo", "post", { text:this.state.newAnswer,tname:this.state.nowtname,username,flag:3 });
  }

  async onCheck(tname) {   //?????????????????????
    console.log('now table:'+tname)
    let result = await this.callFunctionWithErrorHandling("todo", "get", undefined, {tname,flag:'checkhomework'});
    this.setState({
      table:result
    })
  }

  async refresh() {
    await this.getItems();
  }


  render() {

    const items = this.state.items?.map((item, index) =>
      <div key={index} className="item">
        <div className="title">
          <Text         //??????
            content={this.state.items[index].title}
          />
        </div>
        <div className="text">
          <Text         //??????
            content={this.state.items[index].description}
          />
        </div>
        <div className="action">
          <MenuButton
            trigger={<Button content="..." />}
            menu={[
              {
                content: '????????????',
                onClick: () => this.onDeleteItem(item.id,item.title)
              },
              {
                content: '????????????',
                onClick: () => this.setState({ isAddingAnswer: true,nowtname:item.title })
              },
              {
                content: '????????????',
                onClick: () => this.onCheck(item.title)
              }
            ]}
            on="click"
          />
        </div>
      </div>
    );

    const tables = this.state.table?.map((item, index) =>
          <div className='table'  >
              <Text
                content={this.state.table[index].username}
              />
              <Text
              content={this.state.table[index].text}
              />
          </div>
    );
    return (
      <div>
        {this.state.showLoginPage === false && <div className="flex-container">
          <div className="todo-col">
            <div className="todo">
              <div className="header">
                <div className="title">
                  <h2>Homework List</h2>
                  <span>{this.state.items.length} Homework{this.state.items.length === 1 ? "" : "s"}</span>
                </div>

                <div className="add-button">
                  <Button primary onClick={() => this.setState({ isAddingItem: true })}>+ Add Homework</Button>
                </div>
              </div>  
              {items}
              {tables}
              {this.state.isAddingAnswer && <div className="answer add">
                <Input
                autoFocus
                type="text"
                value={this.state.newAnswer}
                onChange={(e) => this.setState({ newAnswer: e.target.value })}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    this.onSub(this.state.username);
                    this.setState({
                        isAddingAnswer: false,
                      });
                  }
                }}
                onBlur={() => {
                  if (this.state.newAnswer) {
                    this.onSub(this.state.username);
                  }
                  this.setState({
                    isAddingAnswer: false,
                  });
                }}
                className="answer"
                />
                </div>
              }
              {this.state.isAddingItem && <div className="item add">
                <Form
                  onSubmit={() => {
                    const tit = document.getElementById("title")
                    const tex = document.getElementById("text")
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
export default Homework;
