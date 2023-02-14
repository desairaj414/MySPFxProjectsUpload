import { INewsLink } from "./INewsLink";

export interface INewsListItem {
  NewsTitle: string;
  PublishDate: string;
  NewsContent: string;
  NewsLink: INewsLink;
}