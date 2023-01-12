import * as React from 'react';
import styles from './MyReactCrudCode.module.scss';
import { IMyReactCrudCodeProps } from './IMyReactCrudCodeProps';
import { IMyReactCrudCodeStates } from './IMyReactCrudCodeStates';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPOperations } from '../Services/SPServices';
import { ICountryListItem } from '../models';
import { Dropdown, IDropdownOption, DefaultButton } from 'office-ui-fabric-react';

export default class MyReactCrudCode extends React.Component<IMyReactCrudCodeProps, IMyReactCrudCodeStates, {}> {

  public _spOps : SPOperations;
  public selectedListTitle: string;

  constructor(props: IMyReactCrudCodeProps) {
    super(props);
    this.state = {listTitles: [], countries: [], status: ""}
    this._spOps = new SPOperations();
  }

  public getListTitle = (event: any, data: any) => {
    this.selectedListTitle = data.text;
    SPOperations.setListTitle(data.text)
  }

  public componentDidMount(){
    this._spOps.GetAllList().then((result:IDropdownOption[])=> {
      this.setState({listTitles: result})
    });
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
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>My SharePoint SPFx CRUD Operations <br/>Code using React!</h2>
          <div>{environmentMessage}</div>
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div className={styles.tryHiding}>
            <div>Web part property value: <strong>{escape(description)}</strong></div>
            <h2>Welcome to SharePoint Framework!</h2>
          </div>
        </div>
        <div className={styles.buttons}>
          <br/>
          <Dropdown  
            className={styles.dropdown}
            options={this.state.listTitles}
            defaultSelectedKey="Countries"
            onChange={this.getListTitle}
          />
          <br/>
          <DefaultButton
            text="Get List Items"
            className={styles.buttons}
            onClick={
              async ()=>{
                const response: ICountryListItem[] = await this._spOps._onGetListItems();
                this.setState({countries: response});
              }
            }
          />
          <DefaultButton
            text="Add List Item"
            type="button"
            className={styles.buttons}
            onClick={
              async ()=>{
                const response: ICountryListItem[] = await this._spOps._onAddListItem();
                this.setState({countries: response});

                // this._spOps.createListItem(this.selectedListTitle)
                //   .then((result: string)=>{
                //     this.setState({status: result});
                //   })
              }
            }
          />
          <DefaultButton
            text="Update List Item"
            type="button"
            className={styles.buttons}
            onClick={
              async ()=>{
                const response: ICountryListItem[] = await this._spOps._onUpdateListItem();
                this.setState({countries: response});

                // this._spOps.updateListItem(this.selectedListTitle)
                //   .then((result: string)=>{
                //     this.setState({status: result});
                //   }) 
              }
            }
          />
          <DefaultButton
            text="Delete List Item"
            type="button"
            className={styles.buttons}
            onClick={
              async ()=>{
                const response: ICountryListItem[] = await this._spOps._onDeleteListItem();
                this.setState({countries: response});

                // this._spOps.deleteListItem(this.selectedListTitle)
                //   .then((result: string)=>{
                //     this.setState({status: result});
                //   }) 
              }
            }
          />
        </div>
        <br/>
        <div className={styles.myStatusBar}>
          {this.state.status}
        </div>
        <div>
          <ul>
            {this.state.countries && this.state.countries.map((list) =>
              <li key={list.Id}>
                <strong>Id:</strong> {list.Id}, <strong>Title:</strong> {list.Title}
              </li>
            )
            }
          </ul>
        </div>
      </section>
    );
  }
}
