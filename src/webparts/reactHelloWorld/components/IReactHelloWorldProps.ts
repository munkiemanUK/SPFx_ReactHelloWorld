import { SPHttpClient } from '@microsoft/sp-http';

export interface IReactHelloWorldProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context:any;
  siteURL: string;
  spHttpClient: SPHttpClient;
}
