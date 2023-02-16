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

  _getNewsList = async ()=>{
    const response: INewsListItem[] = await this._spOps.getNewsList(this.props.maxNews);
    this.setState({newslist: response});
  }

  public componentDidMount()
  {
    this._getNewsList();
  }

  public componentDidUpdate(prevProps : INewsWebPartProps, prevState : INewsWebPartStates) : void 
  {
    if(this.props.maxNews !== prevProps.maxNews)
    {
      this._getNewsList();
    }
  }

  public render(): React.ReactElement<INewsWebPartProps> {

    const {
      hasTeamsContext,
      description,
      maxCharacters,
      maxNews,
      toggle1
    } = this.props;

    return (
      <section className={`${styles.newsWebPart} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
        <DefaultButton
              text="Get Items"
              className={styles.buttons}
              onClick={this._getNewsList}
            />
        </div>
        <div>
            <ul>
              <li><strong>Description:</strong> {description}</li>
              <li><strong>maxCharacters:</strong> {maxCharacters}</li>
              <li><strong>maxNews:</strong> {maxNews}</li>
              <li><strong>toggle1:</strong> {toggle1.toString()}</li>
              <li><strong>API CALL RESULTS</strong></li>
              {this.state.newslist && this.state.newslist.map((news) =>
                <li key={news.NewsTitle}>
                  <strong>NewsTitle:</strong> {news.NewsTitle},<br/> 
                  <strong> Publish Date:</strong> {news.PublishDate},<br/>
                  <strong> Formated Date:</strong> {`${new Date(news.PublishDate).getDate()} ${new Date(news.PublishDate).toLocaleString('default', { month: 'short' })} ${new Date(news.PublishDate).getFullYear()}`.toUpperCase()},<br/>
                  <strong> NewsLink Url:</strong> {news.NewsLink.Url},<br/>
                  <strong> NewsLink Desc:</strong> {news.NewsLink.Description},<br/>
                  <strong> Content:</strong> {news.NewsContent.slice(0,maxCharacters)+"..."}
                </li>
              )}
            </ul>
        </div>
        <div style={{backgroundColor: "#F0F9FA",}}>
          <hr/>
          {this.state.newslist && this.state.newslist.map((news) =>
            <div key={news.NewsTitle}>
              <div className={styles.headerNews}>{`${new Date(news.PublishDate).getDate()} ${new Date(news.PublishDate).toLocaleString('default', { month: 'short' })} ${new Date(news.PublishDate).getFullYear()}`.toUpperCase()}: {news.NewsTitle}</div>
              <div className={styles.headerDetails}>
                <a href={news.NewsLink.Url}>{news.NewsContent.slice(0,maxCharacters)+"..."} ▶</a>
              </div>
              <hr/>
            </div>
          )}
        </div>
      </section>
    );
  }
}
