// import * as React from 'react';
// import styles from './CollaborationOrgAnnouncement.module.scss';
// import type { ICollaborationOrgAnnouncementProps } from './ICollaborationOrgAnnouncementProps';
// import { escape } from '@microsoft/sp-lodash-subset';

// export default class CollaborationOrgAnnouncement extends React.Component<ICollaborationOrgAnnouncementProps> {
//   public render(): React.ReactElement<ICollaborationOrgAnnouncementProps> {
//     const {
//       description,
//       isDarkTheme,
//       environmentMessage,
//       hasTeamsContext,
//       userDisplayName
//     } = this.props;

//     return (
//       <section className={`${styles.collaborationOrgAnnouncement} ${hasTeamsContext ? styles.teams : ''}`}>
//         <div className={styles.welcome}>
//           <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
//           <h2>Well done, {escape(userDisplayName)}!</h2>
//           <div>{environmentMessage}</div>
//           <div>Web part property value: <strong>{escape(description)}</strong></div>
//         </div>
//         <div>
//           <h3>Welcome to SharePoint Framework!</h3>
//           <p>
//             The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
//           </p>
//           <h4>Learn more about SPFx development:</h4>
//           <ul className={styles.links}>
//             <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
//             <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
//             <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
//             <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
//             <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
//             <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
//             <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
//           </ul>
//         </div>
//       </section>
//     );
//   }
// }


import * as React from 'react';
import styles from './CollaborationOrgAnnouncement.module.scss';
import type { ICollaborationOrgAnnouncementProps, ICollaborationOrgAnnouncementState, IImageDetails } from './ICollaborationOrgAnnouncementProps';
import { PrimaryButton } from '@fluentui/react';
// import { Carousel } from '../../../shared/components/Carousel/Carousel';
import { Log } from '@microsoft/sp-core-library';
import { BaseService } from '../../../common/services/BaseService';
import { Carousel } from '../../../common/components/Carousel/Carousel';

const CAROUSEL: string = 'carousel';
// const LIGHTBOX: string = 'lightbox';
const CAROUSELNEW: string = 'carouselNew';

export default class CollaborationOrgAnnouncement extends React.Component<ICollaborationOrgAnnouncementProps, ICollaborationOrgAnnouncementState, {}> {
  
  private service: BaseService;/* To call the service file */
  constructor(props: ICollaborationOrgAnnouncementProps, state: ICollaborationOrgAnnouncementState) {
    super(props);

    this.state = {
      imageCount: 10,
      imageInfo: { info: [] }
    };
    this.service = new BaseService(this.props.context, this.props.webUrl);
    this.getItems = this.getItems.bind(this);
  }
  /**
   * Show the configure message, if ther webpart is not configured
   */
  private showConfigureMessge(): JSX.Element {
    return (
      <div className={styles.center}>
        <h2>Webpart is not configured</h2>
        <PrimaryButton iconProps={{ iconName: "Settings" }} onClick={() => this.props.propertyPane.open()}>Configure</PrimaryButton>
      </div>
    );
  }
  public async componentDidMount(): Promise<void> {
    await this.getItems();
  }
  /**
   * Gets the Images from the SharePoint Library to be displayed in the control
   */
  private async getItems(): Promise<void> {
    // initialize the image variable
    let imageDetails: IImageDetails = { info: [] };

    try {
      const queryURL = this.props.context.pageContext.web.serverRelativeUrl + "/" + this.props.listName;
      const selectquery = "*,FileRef,FileLeafRef"
      const imagedoc = await this.service.getImageItems(queryURL, selectquery);
      imagedoc.forEach((image: { [x: string]: any; }) => {
        imageDetails.info.push({ caption: image["Caption"], description: image["Description"], name: image["FileLeafRef"], path: image["FileRef"], redirectLink: image["RedirectURL"] !== null ? image["RedirectURL"]["Url"] : "#" });
      });
    }
    catch (e) {
      Log.error(this.props.loggerName, new Error(`Error occured in ImageGallery.getItems()`));
      Log.error(this.props.loggerName, e);
    }
    finally {
      this.setState({ imageInfo: imageDetails, imageCount: imageDetails.info.length });
    }
  }
  
  
  public render(): React.ReactElement<ICollaborationOrgAnnouncementProps> {
    const layoutType = this.props.layout === CAROUSELNEW ? 'imageStyle' : 'default';
    return (
    <div>
      <div className={styles.imageCarousel}>
        <div className={styles.container}>
          <h1 className={styles.pagetitle}>{this.props.WebpartTitle}</h1>

          {
            this.props.webUrl && this.props.listName ?

              this.props.layout === CAROUSEL ?
                <React.Suspense fallback={<div>Loading...</div>}>
                  <Carousel duration={this.props.duration}
                    images={this.state.imageInfo}
                    imagesCount={Number(this.state.imageCount)}
                    isAutoRotate={this.props.isAutorotate}
                    height={this.props.height}
                    width={this.props.width}
                    columnSection={this.props.ColumnSection}
                    showCaptions={true}
                    layoutType={layoutType}>
                    </Carousel>
                </React.Suspense>
                :
                this.props.layout === CAROUSELNEW ?
                  <React.Suspense fallback={<div>Loading...</div>}>
                    <Carousel duration={this.props.duration}
                    images={this.state.imageInfo}
                    imagesCount={Number(this.state.imageCount)}
                    isAutoRotate={this.props.isAutorotate}
                    height={this.props.height}
                    width={this.props.width}
                    columnSection={this.props.ColumnSection}
                    showCaptions={true}
                     layoutType={layoutType}>
                   
                    </Carousel>
                  </React.Suspense>
                  :
                  <React.Suspense fallback={<div>Loading...</div>}>
                    {/* <List imagesCount={this.state.imageCount} images={this.state.imageInfo}></List> */}
                  </React.Suspense>
              :
              this.showConfigureMessge()
          }

        </div>
      </div>
 {/* new */}
 {/* <div>
   <div className={styles.container}>
      <Carousel
         images={this.state.imageInfo}
        imagesCount={Number(this.state.imageCount)}
          height={this.props.height}
          width={this.props.width}
        isAutoRotate={true}
        duration={5000}
         columnSection={this.props.ColumnSection}
         showCaptions={true}
        >
         
        </Carousel>
    </div>
 </div> */}
      
</div>
    );
  }
}
