import * as React from 'react';
import styles from './ReactCrud.module.scss';
import { IReactCrudProps } from './IReactCrudProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IReactCrudState } from './IReactCrudState';
import { SPOperations } from "../../Services/SPServices";
import { Button, Dropdown, IDropdownOption } from "office-ui-fabric-react"

export default class ReactCrud extends React.Component<IReactCrudProps, IReactCrudState, {}> {

  public _spOps : SPOperations;
  public selectedListTitle: string;

  constructor(props: IReactCrudProps){
    super(props);
    this._spOps = new SPOperations();
    this.state = {listTitles:[], status:""}
  }

  public getListTitle = (event: any, data: any) => {
    this.selectedListTitle = data.text;
  }

  public componentDidMount(){
    this._spOps.GetAllList(this.props.context).then((result:IDropdownOption[])=> {
      this.setState({listTitles: result})
    });
  }

  public render(): React.ReactElement<IReactCrudProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.reactCrud} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
          <h2>Welcome to SharePoint Framework!</h2>
          <h3>SharePoint CRUD Operations using Rest API (spHTTPClient)!</h3>
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
            onClick={()=>
              this._spOps
                .createListItem(this.props.context, this.selectedListTitle)
                .then((result: string)=>{
                  this.setState({status: result});
                })
              }
          ></Button>
          <Button 
            className={styles.myButton} 
            text="Update List Item"
            onClick={()=>
              this._spOps
                .updateListItem(this.props.context, this.selectedListTitle)
                .then((result: string)=>{
                  this.setState({status: result});
                }) 
              }
          ></Button>
          <Button 
            className={styles.myButton} 
            text="Delete List Item"
            onClick={()=>
              this._spOps
                .deleteListItem(this.props.context, this.selectedListTitle)
                .then((result: string)=>{
                  this.setState({status: result});
                }) 
              }
          ></Button>
          <div className={styles.myStatusBar}>{this.state.status}</div>
        </div>
      </section>
    );
  }
}
