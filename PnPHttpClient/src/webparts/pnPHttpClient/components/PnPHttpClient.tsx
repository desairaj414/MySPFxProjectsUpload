import * as React from 'react';
import styles from './PnPHttpClient.module.scss';
import { IPnPHttpClientProps } from './IPnPHttpClientProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class PnPHttpClient extends React.Component<IPnPHttpClientProps, {}> {
  public render(): React.ReactElement<IPnPHttpClientProps> {
    const {
      spListItems,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.pnPHttpClient} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
        </div>
        <div className={styles.buttons}>
          <button type="button" onClick={this.onGetListItemsClicked}>Get Countries</button>
          <button type="button" onClick={this.onAddListItemClicked}>Add List Item</button>
          <button type="button" onClick={this.onUpdateListItemClicked}>Update List Item</button>
          <button type="button" onClick={this.onDeleteListItemClicked}>Delete List Item</button>
        </div>
        <div>
          <table>
          {spListItems.length==0
            ?<> </>
            :<tr>
              <th>Id</th>
              <th>Title</th>
            </tr>
          }
          {spListItems && spListItems.map((list) =>
            <tr key={list.Id}>
              <td>{list.Id}</td>
              <td>{list.Title}</td>
            </tr>
          )}
          </table>
        </div>
      </section>
    );
  }

  private onGetListItemsClicked = (event: React.MouseEvent<HTMLButtonElement>): void => {
    event.preventDefault();
    this.props.onGetListItems();
  }

  private onAddListItemClicked = (event: React.MouseEvent<HTMLButtonElement>): void => {
    event.preventDefault();
    this.props.onAddListItem();
  }
  
  private onUpdateListItemClicked = (event: React.MouseEvent<HTMLButtonElement>): void => {
    event.preventDefault();
    this.props.onUpdateListItem();
  }
  
  private onDeleteListItemClicked = (event: React.MouseEvent<HTMLButtonElement>): void => {
    event.preventDefault();
    this.props.onDeleteListItem();
  }

}
