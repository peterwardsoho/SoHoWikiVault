import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ISohoWikiVaultProps {
  context: WebPartContext;
  // webPartTitle: string;
  setupComelete: boolean;
  wikiId: string;
}
