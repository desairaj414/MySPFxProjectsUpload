import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IHelloWorldProps {
  description: string;
  context: WebPartContext;
  webURL:string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
