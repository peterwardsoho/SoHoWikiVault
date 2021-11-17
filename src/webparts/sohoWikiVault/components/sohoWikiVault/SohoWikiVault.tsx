import * as React from 'react';
import styles from './SohoWikiVault.module.scss';
import { ISohoWikiVaultProps } from './ISohoWikiVaultProps';
import { escape } from '@microsoft/sp-lodash-subset';

// Importing components
import { WelcomeScreen, MainCategory } from '../';

export class SohoWikiVault extends React.Component<ISohoWikiVaultProps, {}> {
  public render(): React.ReactElement<ISohoWikiVaultProps> {
    return (
      <div>
        {this.props.setupComelete ?
          <MainCategory context={this.props.context} wikiId={this.props.wikiId}></MainCategory> :
          <WelcomeScreen />}
      </div>
    );
  }
}
