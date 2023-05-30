import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IKudoProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  siteURL:string;
  componentTitle:string;
  listName:string;
  emptyMessage:string;
  seeAllPageURL:any;
  webHeight:any;
  noofKudos:string;
  context:WebPartContext
  
}
