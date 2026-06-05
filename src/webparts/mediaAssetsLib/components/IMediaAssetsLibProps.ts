import { SPHttpClient } from '@microsoft/sp-http';

export interface IMediaAssetsLibProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  spHttpClient: SPHttpClient;
  siteUrl: string;
}