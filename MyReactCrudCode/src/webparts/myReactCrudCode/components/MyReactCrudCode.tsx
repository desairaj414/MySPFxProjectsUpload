import * as React from 'react';
import { renderToString } from 'react-dom/server'
import styles from './MyReactCrudCode.module.scss';
import { IMyReactCrudCodeProps } from './IMyReactCrudCodeProps';
import { IMyReactCrudCodeStates } from './IMyReactCrudCodeStates';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPOperations } from '../Services/SPServices';
import { ICountryListItem } from '../models';
import { Dropdown, IDropdownOption, DefaultButton, Label, DatePicker, TextField } from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { MSGraphClientV3 } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import 'office-ui-fabric-react/dist/css/fabric.css';

export default class MyReactCrudCode extends React.Component<IMyReactCrudCodeProps, IMyReactCrudCodeStates, {}> {

  public _spOps : SPOperations;
  public selectedListTitle: string;

  constructor(props: IMyReactCrudCodeProps) {
    super(props);
    this.state = {
      listTitles: [], 
      countries: [], 
      status: "",
      // REACT Form CRUD States
      Items: [],
      EmployeeName: "",
      EmployeeNameId: 0,
      ID: 0,
      HireDate: null,
      JobDescription: "",
      HTML: (<table></table>) as JSX.Element,
      // Microsoft Graph API State
      MSGraphHTML: (<></>) as JSX.Element
    };
    this._spOps = new SPOperations();
  }

  public componentDidMount(){
    this._spOps.GetAllList().then((result:IDropdownOption[])=> {
      this.setState({listTitles: result})
    });
    //this.fetchData();
    this.getMails();
  }

  public render(): React.ReactElement<IMyReactCrudCodeProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <section className={`${styles.myReactCrudCode} ${hasTeamsContext ? styles.teams : ''}`}>

        {/* Introduction Part */}
        <div className="ms-Grid" dir="ltr">
          <div className="ms-Grid-row" style={{display: "flex",  alignItems: "center"}}>
            <div className="ms-Grid-col ms-sm3 ms-md3 ms-lg3">
              <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={`${styles.welcomeImage}`} />
            </div>
            <div className="ms-Grid-col ms-sm9 ms-md9 ms-lg9">
            <h2>My SharePoint SPFx CRUD OperationsCode using React!</h2>
              <div>{environmentMessage}</div>
              <p>Well done, {escape(userDisplayName)}!</p>
              <div className={styles.tryHiding}>
                <div>Web part property value: <strong>{escape(description)}</strong></div>
                <h2>Welcome to SharePoint Framework!</h2>
              </div>
            </div>
          </div>
        </div>
        <div className={styles.welcome}>
          {/* <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={`${styles.welcomeImage} ${styles.tryHiding}`} />
          <h2>My SharePoint SPFx CRUD OperationsCode using React!</h2>
          <div>{environmentMessage}</div>
          <p>Well done, {escape(userDisplayName)}!</p>
          <div className={styles.tryHiding}>
            <div>Web part property value: <strong>{escape(description)}</strong></div>
            <h2>Welcome to SharePoint Framework!</h2>
          </div> */}
        </div>
        

        {/* Heading of Normal CRUD */}
        <div className={styles.grey}>
          <hr/>
          <h3>
            CRUD Operations with Custom List Selection
          </h3>
          <hr/>
        </div>

        {/* Dropdown & Buttons of Normal CRUD */}
        <div>
          <br/>
          <Dropdown  
            className={styles.dropdown}
            options={this.state.listTitles}
            defaultSelectedKey="Countries"
            onChange={this.getListTitle}
          />
          <br/>
          <div className={styles.welcome}>
            <DefaultButton
              text="Get Items"
              className={styles.buttons}
              onClick={
                async ()=>{
                  const response: ICountryListItem[] = await this._spOps._getListItemsShort();
                  this.setState({countries: response});
                }
              }
            />
            <DefaultButton
              text="Clear Items"
              type="button"
              className={styles.buttons}
              onClick={()=>{this.setState({countries: []})}}
            />
            <DefaultButton
              text="Add Item"
              type="button"
              className={styles.buttons}
              onClick={
                async ()=>{
                  const response: ICountryListItem[] = await this._spOps._addListItemShort();
                  this.setState({countries: response});

                  // this._spOps.createListItem(this.selectedListTitle)
                  //   .then((result: string)=>{
                  //     this.setState({status: result});
                  //   })
                }
              }
            />
            <DefaultButton
              text="Update Item"
              type="button"
              className={styles.buttons}
              onClick={
                async ()=>{
                  const response: ICountryListItem[] = await this._spOps._updateListItemShort();
                  this.setState({countries: response});

                  // this._spOps.updateListItem(this.selectedListTitle)
                  //   .then((result: string)=>{
                  //     this.setState({status: result});
                  //   }) 
                }
              }
            />
            <DefaultButton
              text="Delete Item"
              type="button"
              className={styles.buttons}
              onClick={
                async ()=>{
                  const response: ICountryListItem[] = await this._spOps._deleteListItemShort();
                  this.setState({countries: response});

                  // this._spOps.deleteListItem(this.selectedListTitle)
                  //   .then((result: string)=>{
                  //     this.setState({status: result});
                  //   }) 
                }
              }
            />
          </div>
        </div>
        <br/>

        {/* Table & Status of Normal CRUD */}
        <div>
          <div className={styles.myStatusBar}>
            {this.state.status}
          </div>
          <div>
          {/* <div>
            <ul>
              {this.state.countries && this.state.countries.map((list) =>
                <li key={list.Id}>
                  <strong>Id:</strong> {list.Id}, <strong>Title:</strong> {list.Title}
                </li>
              )
              }
            </ul>
          </div> */}
          </div>
          <div>
            <table>
              {this.state.countries.length==0 
                ?<></>
                :<tr>
                <th>Id</th>
                <th>Title</th>
              </tr>
              }
              {this.state.countries && this.state.countries.map((list) =>
                <tr key={list.Id}>
                  <td>{list.Id}</td>
                  <td>{list.Title}</td>
                </tr>
              )}
            </table>
          </div>
        </div>
        <br/>

        {/* Heading of Form CRUD */}
        <div className={styles.grey}>
          <hr/>
          <h3>
            CRUD Operations with Custom Values Edit
          </h3>
          <hr/>
        </div>

        {/* Form of Form CRUD*/}
        <div>
          <form>
            <div>
              <Label>Employee Name</Label>
              <PeoplePicker
                context={SPOperations.getContext() as any}
                personSelectionLimit={1}
                // defaultSelectedUsers={this.state.EmployeeName===""?[]:this.state.EmployeeName}
                required={false}
                onChange={this._getPeoplePickerItems}
                defaultSelectedUsers={[this.state.EmployeeName?this.state.EmployeeName:""]}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}
                ensureUser={true}
              />
            </div>
            <div>
              <Label>Hire Date</Label>
              <DatePicker maxDate={new Date()} allowTextInput={false} strings={this._spOps.DatePickerStrings} value={this.state.HireDate} onSelectDate={(e) => { this.setState({ HireDate: e }); }} ariaLabel="Select a date" formatDate={this._spOps.FormatDate} />
            </div>
            <div>
              <Label>Job Description</Label>
              <TextField value={this.state.JobDescription} multiline onChange={this.onchange} />
            </div>
          </form>
        </div>
        <br/>

        {/*Buttons of Form CRUD */}
        <div className={styles.welcome}>
          <DefaultButton 
            className={styles.buttons}
            text="Get" 
            onClick={
              async() => {
                this.setState({EmployeeName:"",HireDate:null,JobDescription:""});
                this.fetchData();
              }
            }
          />
          <DefaultButton 
            className={styles.buttons}
            text="Clear" 
            onClick={() => {
                this.setState({EmployeeName:"",HireDate:null,JobDescription:"",HTML:(<table></table>) as JSX.Element});
              }
            }
          />
          <DefaultButton 
            className={styles.buttons}
            text="Create" 
            onClick={
              async() => {
                await this._spOps.SaveData(this.state.EmployeeNameId,this.state.HireDate,this.state.JobDescription);
                this.setState({EmployeeName:"",HireDate:null,JobDescription:""});
                this.fetchData();
              }
            }
          />
          <DefaultButton 
            className={styles.buttons} 
            text="Update" 
            onClick={
              async() => {
                await this._spOps.UpdateData(this.state.ID,this.state.EmployeeNameId,this.state.HireDate,this.state.JobDescription);
                this.setState({EmployeeName:"",HireDate:null,JobDescription:""});
                this.fetchData();
              }
            }
          />
          <DefaultButton 
            className={styles.buttons} 
            text="Delete" 
            onClick={
              async() => {
                await this._spOps.DeleteData(this.state.ID);
                this.setState({EmployeeName:"",HireDate:null,JobDescription:""});
                this.fetchData();
              }
            }
          />
        </div> 
        <br/>

        {/* Table of Form CRUD */}
        <div> 
          {this.state.HTML}
        </div>
        <br/>

        {/* Heading of Microsoft Graph API */}
        <div className={styles.grey}>
          <hr/>
          <h3>
            Microsoft Graph API Trial
          </h3>
          <hr/>
        </div>

        {/* Microft Graph API Implementation */}
        <div>
          {this.state.MSGraphHTML}
        </div>
        <br/><hr/>

      </section>
    );
  }

  public getListTitle = (event: any, data: any) => {
    this.selectedListTitle = data.text;
    SPOperations.setListTitle(data.text)
  }

  // Form CRUD Operations Functions

  public fetchData = async () => {
    this._spOps.fetchData().then((items:any)=>{
      this.setState({ Items: items });
      this.getHTML(items).then((html)=>{
        this.setState({ HTML: html });
      })
    })
  }

  public async getHTML(items: any[]) {
    var tabledata: JSX.Element = <table>
      <thead>
        <tr>
          <th>EmployeeName</th>
          <th>HireDate</th>
          <th>JobDescription</th>
        </tr>
      </thead>
      <tbody>
        {items && items.map((item, i) => {
          return [
            <tr key={i} onClick={()=>this.findData(item.ID)}>
              <td>{item.EmployeeName.Title}</td>
              <td>{this._spOps.FormatDate(item.HireDate)}</td>
              <td>{item.JobDescription}</td>
            </tr>
          ];
        })}
      </tbody>

    </table>;
    return await tabledata;
  }

  public findData = (id: any): void => {
    //this.fetchData();
    var itemID = id;
    var allitems = this.state.Items;
    var allitemsLength = allitems.length;
    if (allitemsLength > 0) {
        for (var i = 0; i < allitemsLength; i++) {
            if (itemID == allitems[i].Id) {
                this.setState({
                ID:itemID,
                EmployeeName:allitems[i].EmployeeName.Title,
                EmployeeNameId:allitems[i].EmployeeNameId,
                HireDate:new Date(allitems[i].HireDate),
                JobDescription:allitems[i].JobDescription
                });
            }
        }
    }
  }

  public onchange = (value: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, stateValue: string) => {
    this.setState({ JobDescription: stateValue.toString() });
  }

  public _getPeoplePickerItems = async (items: any[]) => {
    if (items.length > 0) {
        this.setState({ EmployeeName: items[0].text });
        this.setState({ EmployeeNameId: items[0].id });
    }
    else {
        //ID=0;
        this.setState({ EmployeeNameId: "" });
        this.setState({ EmployeeName: "" });
    }
  }

  public getMails(){
    SPOperations.getContext().msGraphClientFactory.getClient('3').then(
      (client: MSGraphClientV3): void => {
        // get information about the current user from the Microsoft Graph
        client
        .api('/me/messages')
        .top(5)
        .orderby("receivedDateTime desc")
        .get((_error: any, messages: any, rawResponse?: any) => {
          // List the latest emails based on what we got from the Graph
          console.log(messages.value)
          this._renderEmailList(messages.value);
        });
      }
    );
  }

  private _renderEmailList(messages: MicrosoftGraph.Message[]): void {
    let html:JSX.Element[] = []
    let html2: JSX.Element = 
    <div>
      {messages && messages.map((message)=>{
        <p className="${styles.welcome}">Email  - {message.subject}</p>
      })}
    </div>;
    console.log("Hiii" +renderToString(html2))
    for (let index = 0; index < messages.length; index++) {
      html.push(<p className="${styles.welcome}">Email {index + 1} - {messages[index].subject}</p>);
    }
    this.setState({ MSGraphHTML: html });
  }

}
