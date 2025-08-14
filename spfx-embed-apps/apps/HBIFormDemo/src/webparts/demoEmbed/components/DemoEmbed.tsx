import * as React from 'react';
import styles from './DemoEmbed.module.scss';
import type { IDemoEmbedProps } from './IDemoEmbedProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { createSharePointClient, getCurrentUser, ensureListExists, createListItem } from '@hbi/sp-client';

export default class DemoEmbed extends React.Component<IDemoEmbedProps> {
  private handleClick = async () => {
    const sp = createSharePointClient({ spfxContext: this.props.context });
    const user = await getCurrentUser(sp);
    const listTitle = 'DemoTestList';
    const exists = await ensureListExists(sp, listTitle);
    if (exists) {
      await createListItem(sp, { listTitle, fields: { Title: `Hello ${user.displayName || user.email || user.loginName}` } });
      alert('Wrote item to DemoTestList');
    } else {
      alert(`List ${listTitle} does not exist. Read user: ${user.displayName || user.email || user.loginName}`);
    }
  };
  public render(): React.ReactElement<IDemoEmbedProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.demoEmbed} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <button onClick={this.handleClick}>Validate user and write demo item</button>
        </div>
      </section>
    );
  }
}
