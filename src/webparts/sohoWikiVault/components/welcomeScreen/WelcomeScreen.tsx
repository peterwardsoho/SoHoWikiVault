import * as React from 'react';
import styles from './welcomeScreen.module.scss';

export class WelcomeScreen extends React.Component<{}, {}> {
    public render() {
      return (
        <div className={styles.welcomeScreen}>
          <div className={styles.container}>
            <div className={styles.row}>
              <div className={styles.column}>
                <span className={styles.title}>Welcome to Soho Wiki Valut</span>
                <p className={styles.subTitle}>Please install settings from Webpart ProprtyPane.</p>
              </div>
            </div>
          </div>
        </div>
      );
    }
  }