import * as React from 'react';
import styles from './NewsWebPart.module.scss';
import { INewsWebPartProps } from './INewsWebPartProps';
import { INewsWebPartStates } from './INewsWebPartStates';
import { SPOperations } from '../Services/SPServices';
import { DefaultButton } from 'office-ui-fabric-react';
import { INewsListItem } from '../models';
//import { escape } from '@microsoft/sp-lodash-subset';

export default class NewsWebPart extends React.Component<INewsWebPartProps, INewsWebPartStates ,{}> {

  public _spOps : SPOperations;

  constructor(props: INewsWebPartProps) {
    super(props);
    this.state = {
      newslist: [],
    };
    this._spOps = new SPOperations();
  }

  public render(): React.ReactElement<INewsWebPartProps> {
    const {
      hasTeamsContext,
    } = this.props;

    return (
      <section className={`${styles.newsWebPart} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
        <DefaultButton
              text="Get Items"
              className={styles.buttons}
              onClick={
                async ()=>{
                  const response: INewsListItem[] = await this._spOps.getNewsList();
                  this.setState({newslist: response});
                }
              }
            />
        </div>
        <div>
            <ul>
              {this.state.newslist && this.state.newslist.map((news) =>
                <li key={news.NewsTitle}>
                  <strong>NewsTitle:</strong> {news.NewsTitle},<br/> 
                  <strong> Publish Date:</strong> {news.PublishDate},<br/>
                  <strong> NewsLink:</strong> {news.NewsLink.Url},<br/>
                  <strong> Content:</strong> {news.NewsContent.slice(0,20)+"..."}
                </li>
              )}
            </ul>
        </div>
        <div style={{backgroundColor: "#F0F9FA",}}>
          <hr/>
          {this.state.newslist && this.state.newslist.map((news) =>
            <div key={news.NewsTitle}>
              <div className={styles.headerNews}>{news.PublishDate}: {news.NewsTitle}</div>
              <div className={styles.headerDetails}>
                <a href={news.NewsLink.Url}>{news.NewsContent.slice(0,60)+"..."} ▶</a>
              </div>
              <hr/>
            </div>
          )}
        </div>
      </section>
    );
  }
}
