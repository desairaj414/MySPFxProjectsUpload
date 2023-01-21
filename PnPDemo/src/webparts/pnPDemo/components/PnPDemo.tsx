import * as React from 'react';
import styles from './PnPDemo.module.scss';
import { IPnPDemoProps } from './IPnPDemoProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IPnPDemoState } from './IPnPDemoState';
import { SPOperations } from "../Services/SPOps";
import { Button, Dropdown, IDropdownOption } from "office-ui-fabric-react"

export default class PnPDemo extends React.Component<IPnPDemoProps, IPnPDemoState, {}> {

  public _spOps : SPOperations;
  public selectedListTitle: string;

  constructor(props: IPnPDemoProps){
    super(props);
    this._spOps = new SPOperations();
    this.state = {listTitles:[], status:""}
  }

  public getListTitle = (event: any, data: any) => {
    this.selectedListTitle = data.text;
  }

  public componentDidMount(){
    this._spOps.getListTitles().then((result:IDropdownOption[])=> {
      this.setState({listTitles: result})
    });
  }

  public render(): React.ReactElement<IPnPDemoProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.pnPDemo} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
          <h2>Welcome to SharePoint Framework!</h2>
          <h3>SharePoint CRUD Operations using Pnp Js!</h3>
        </div>
        <div id ="dv_Parent" className={styles.myStyles}>
          <Dropdown  
            className={styles.dropdown} 
            options={this.state.listTitles}
            placeholder="**Select Your List**"
            onChange={this.getListTitle}
          ></Dropdown>
          <br/>
          <Button 
            className={styles.myButton} 
            text="Create List Item"
            onClick={() => {
              this._spOps
              .createListItem(this.selectedListTitle)
              .then((result: string)=>{
                this.setState({status: result});
              })
            }}
          ></Button>
          <Button 
            className={styles.myButton} 
            text="Update List Item"
            onClick={() => {
              this._spOps
              .updateListItem(this.selectedListTitle)
              .then((result: string)=>{
                this.setState({status: result});
              }) 
            }}
          ></Button>
          <Button 
            className={styles.myButton} 
            text="Delete List Item"
            onClick={() => {
              this._spOps
                .deleteListItem(this.selectedListTitle)
                .then((result: string)=>{
                  this.setState({status: result});
                }) 
            }}
          ></Button>
          <div className={styles.myStatusBar}>{this.state.status}</div>
        </div>
      </section>
    );
  }
}
