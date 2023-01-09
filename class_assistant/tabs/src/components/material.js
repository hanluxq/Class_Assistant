// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
import './materiall.css'
import React from 'react';
import axios from 'axios';
import { TeamsFx } from "@microsoft/teamsfx";
import { Checkbox, Button, Input, MenuButton,Text,Form ,FormInput, FormCheckbox, FormButton} from "@fluentui/react-northstar"

class Material extends React.Component {

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
      fileurl:"",
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
    //let result = await this.callFunctionWithErrorHandling("todo", "get", undefined, { flag:'material' });
    let result = await this.callFunctionWithErrorHandling("todo", "get", undefined, { flag:'getfromazure' });
    console.log(result)
    if ("Error" === result) {
      throw new Error("todo Function failed, please check Azure Functions log for details!");
    } else {
      this.setState({
        items: result,
        initialized: true,
      });
    }
    //this.listblob()
  }

  async listblob(){   //下拉列表显示
    const fileList = document.getElementById("file-list");
    fileList.size = 0;
    fileList.innerHTML = "";
    
    const { BlobServiceClient } = require("@azure/storage-blob");
    const sasurl='https://todolist99bbc6api.blob.core.windows.net/?sv=2020-08-04&ss=bfqt&srt=sco&sp=rwdlacupitfx&se=2032-05-18T19:51:46Z&st=2022-05-18T11:51:46Z&spr=https&sig=X6ORB96vxu3P%2Fq8h%2FCEVzSX%2BsQzF8zeGngLo7XXwYgw%3D'
    const blobServiceClient = new BlobServiceClient(sasurl);
    const containerName = 'photos';
    const containerClient = blobServiceClient.getContainerClient(containerName);
    let iter = containerClient.listBlobsFlat();
    let blobItem = await iter.next();
    while (!blobItem.done) {
      fileList.size += 1;
      fileList.innerHTML += `<option>${blobItem.value.name}</option>`;
      blobItem = await iter.next();
      console.log(blobItem)
    }
    {/* <select id="file-list" multiple style={{height:'222px', width: '593px', overflow: 'scroll'}} /> */}
  }

  async onDelete(id,name){
    const newItems = this.state.items.filter(item => item.id !== id);
    this.setState({
      items: newItems
    })
    await this.callFunctionWithErrorHandling("todo", "delete", { name, flag:3 });
  }

  async uploadFiles(){
    console.log("Uploading files...");
    var file=document.getElementById('file').files[0];
    const { BlobServiceClient } = require("@azure/storage-blob");
    const sasurl='https://todolist99bbc6api.blob.core.windows.net/?sv=2020-08-04&ss=bfqt&srt=sco&sp=rwdlacupitfx&se=2032-05-18T19:51:46Z&st=2022-05-18T11:51:46Z&spr=https&sig=X6ORB96vxu3P%2Fq8h%2FCEVzSX%2BsQzF8zeGngLo7XXwYgw%3D'
    const blobServiceClient = new BlobServiceClient(sasurl);
    const containerName = 'photos';
    const containerClient = blobServiceClient.getContainerClient(containerName);
    const blockBlobClient = containerClient.getBlockBlobClient(file.name);
    blockBlobClient.uploadBrowserData(file)
    console.log("Uploaded success");
    this.refresh();
}

  async  downloadBlobToFile(blobName) {
    console.log("Downloading files...");
    const { BlobServiceClient } = require("@azure/storage-blob");
    const sasurl='https://todolist99bbc6api.blob.core.windows.net/?sv=2020-08-04&ss=bfqt&srt=sco&sp=rwdlacupitfx&se=2032-05-18T19:51:46Z&st=2022-05-18T11:51:46Z&spr=https&sig=X6ORB96vxu3P%2Fq8h%2FCEVzSX%2BsQzF8zeGngLo7XXwYgw%3D'
    const blobServiceClient = new BlobServiceClient(sasurl);
    const containerName = 'photos';
    const containerClient = blobServiceClient.getContainerClient(containerName);
    const blobClient = await containerClient.getBlobClient(blobName);
    const blob = await blobClient.downloadToFile();
    const url = blob._response.request.url;
    console.log(url)
    console.log(`download of ${blobName} success`);
    window.open(url, "_blank");
  }


  async refresh() {
    await this.getItems();
  }

  render() {
    const items = this.state.items?.map((item, index) =>
      <div key={index} className="item">
        <div classneme='textname'><Text  content={this.state.items[index].name}/></div>
        <div classneme='texttime'><Text  content={this.state.items[index].createdon}/></div>
        <div classneme='textsize'><Text  content={(this.state.items[index].size/1024).toFixed(2)+'  KB'} /></div>
        <Button className='downloadbutton' primary onClick={() => this.downloadBlobToFile(this.state.items[index].name) }>下载</Button>
      </div>
    );

    return (
      <div>
        {this.state.showLoginPage === false && 
        <div className="flex-container">
          <div className="todo-col">
            <div className="todo">

              <div className="header">
                <div className="title">
                  <h2>Material List</h2>
                  <span>{this.state.items.length} Material{this.state.items.length === 1 ? "" : "s"}</span>
                </div>
                
                <Input type={'file'} label='上传资料' id='file' onChange={()=>{this.setState({isAddingItem:true})}}></Input>
                {this.state.isAddingItem&&<Button onClick={()=>this.uploadFiles()} className='submitbutton'>提交</Button>}
                
              </div>  

              {/* {this.state.items.length > 0 && <div className="header-container">
                <div className="name">
                  <svg width="18" height="18" class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 20" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 6h16M4 10h16M4 14h16M4 18h16"></path></svg>
                  <span>name</span>
                </div>
                <div className="time">
                  <svg width="16" height="16" class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z"></path></svg>
                  <span>upload time</span>
                </div>
                <div className="size">
                <svg width="16" height="16" class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 21h10a2 2 0 002-2V9.414a1 1 0 00-.293-.707l-5.414-5.414A1 1 0 0012.586 3H7a2 2 0 00-2 2v14a2 2 0 002 2z"></path></svg>
                  <span>size</span>
                </div>
              </div>} */}

              {items}

            </div>
          </div>
        </div>}
      </div>
    );
    }
}
export default Material;
