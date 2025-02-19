import * as React from 'react';
import styles from './CollaborationEvents.module.scss';
import { ICollaborationEventsProps } from '../interfaces';

export default class CollaborationEvents extends React.Component<ICollaborationEventsProps> {
  public render(): React.ReactElement<ICollaborationEventsProps> {
    const {
      hasTeamsContext,
    } = this.props;

    return (
      <section className={`${styles.collaborationEvents} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          hiii
        </div>
      </section>
    );
  }
}
