import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IHelloWorldProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  personName: string;
  listName: string;
  context:any;
  webURL:string;
}
