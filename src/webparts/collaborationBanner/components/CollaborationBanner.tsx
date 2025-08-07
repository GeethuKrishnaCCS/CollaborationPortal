import * as React from 'react';
import styles from './CollaborationBanner.module.scss';
import type { ICollaborationBannerProps } from '../interfaces/ICollaborationBannerProps';


export default class CollaborationBanner extends React.Component<ICollaborationBannerProps> {
  public render(): React.ReactElement<ICollaborationBannerProps> {
    const {
      userDisplayName,
      hasTeamsContext,
      BannerImageUrl,
    } = this.props;

    return (
      <section className={`${styles.collaborationBanner} ${hasTeamsContext ? styles.teams : ''}`}>

        {this.props.selectDesign === 'ImageOverlap' && (
          <div>
            <div className={styles.background} style={{ backgroundImage: `url(${BannerImageUrl})` }}></div>
            <div className={styles.content} style={{ backgroundColor: this.props.imageOverlapBackgroundColor }}>
              <div className={styles.welcomecontainer}>
                <div className={styles.welcomeText} style={{ color: this.props.welcomeheadingfontcolor }}>{this.props.welcomeText}</div>
                <span className={styles.userName} style={{ color: this.props.usernameheadingfontcolor }}>{userDisplayName}!</span>
              </div>
              <div className={styles.missioncontainer}>
                <span className={styles.missionDescription} style={{ color: this.props.descriptionfontcolor }}>{this.props.missionDescription}</span>
              </div>

            </div>
          </div>
        )}

        {this.props.selectDesign === 'New' && (
          <div className={styles.newBanner}>
            <div className={styles.newLeftColumn}>
              <div className={styles.newTextContainer}>
                <span className={styles.newWelcomeBold} style={{ color: this.props.welcomeheadingfontcolor }}>{this.props.welcomeText}</span>
                <span className={styles.newUserHighlight} style={{ color: this.props.usernameheadingfontcolor }}>{userDisplayName}!</span>
                <div className={styles.newSubtitle} style={{ color: this.props.descriptionfontcolor }}>{this.props.missionDescription}</div>
              </div>
            </div>
            <div className={styles.newRightColumn}>
              <img src={BannerImageUrl} alt="Banner" className={styles.newBannerImage} />
            </div>
          </div>
        )}

      </section>
    );
  }
}
