import * as React from 'react';
import styles from './CollaborationMediaGallery.module.scss';
import type { ICollaborationMediaGalleryProps, ICollaborationMediaGalleryState, IVideo } from './ICollaborationMediaGalleryProps';
import CollaborationMediaGalleryService from '../services/CollaborationMediaGalleryService';
import { MediaList } from './MediaList';
import { VideoPlayer } from './VideoPlayer';

export default class CollaborationMediaGallery extends React.Component<ICollaborationMediaGalleryProps,ICollaborationMediaGalleryState, {}> {
  constructor(props: ICollaborationMediaGalleryProps) {
    super(props);
    this.state = {
      items: [],
      selectedMedia: undefined,
      loading: false,
      error: ''
    };
  }

  public async componentDidMount(): Promise<void> {
    await this.getAllItems();
    window.addEventListener('resize', this.adjustDiv2Height); // Add the resize event listener

  }
  public componentWillUnmount(): void {
    window.removeEventListener('resize', this.adjustDiv2Height); // Clean up the event listener when unmounting
  }
  componentDidUpdate(): void {
    setTimeout(() => {
      this.adjustDiv2Height();
    }, 0);
  }
  // private async getAllItems(): Promise<void> {
  //   this.setState({ loading: true });
  //   try {
  //     const listId = this.props.listId;
  //     console.log('Fetching media items...');
  //     const items = await MediaGalleryService.getAllItems(listId, this.props.numberOfItems);

  //     if (items && items.length > 0) {
  //       const updatedItems = items.map((item, index) => ({
  //         ID: item.ID,
  //         title: item.Title,
  //         description: item.MediaDescription,
  //         serverRelativeUrl: item.File.ServerRelativeUrl,
  //         isSelected: index === 0 ? true : false
  //       }));
  //       console.log('Fetched items- transformed:', updatedItems);

  //       this.setState({ items: updatedItems, selectedMedia: updatedItems[0], loading: false });
  //     } else {
  //       console.warn('No items found in the list.');
  //       this.setState({ loading: false, error: 'No items found.' });
  //     }
  //   } catch (error) {
  //     console.error('Error fetching media items:', error);
  //     this.setState({ loading: false, error: 'Failed to fetch media items. Please try again.' });
  //   }
  // }
  private async getAllItems(): Promise<void> {
    this.setState({ loading: true });

    try {
      const listId = this.props.listId;
      console.log('Fetching media items...');

      // Fetch items from the list
      const items = await CollaborationMediaGalleryService.getAllItems(listId, this.props.numberOfItems);

      if (items && items.length > 0) {
        const updatedItems = await Promise.all(
          items.map(async (item:any, index:any) => {
            return ({
              ID: item.ID,
              title: item.Title,
              description: item.MediaDescription,
              serverRelativeUrl: item.File.ServerRelativeUrl,
              pURL: `${window.location.origin}/_api/v2.0/sharePoint:${encodeURIComponent(item.File.ServerRelativeUrl)}:/driveItem/thumbnails/0/large/content?preferNoRedirect=true`,
              thumbnailURL: '',
              isSelected: index === 0,
            });
          }
          )
        );

        console.log('Fetched items - transformed:', updatedItems);

        this.setState({
          items: updatedItems,
          selectedMedia: updatedItems[0],
          loading: false,
        });
      } else {
        console.warn('No items found in the list.');
        this.setState({ loading: false, error: 'No items found.' });
      }
    } catch (error) {
      console.error('Error fetching media items:', error);
      this.setState({ loading: false, error: 'Failed to fetch media items. Please try again.' });
    }
  }

  public handleVideoSelect = (selectedItem: IVideo): void => {
    console.log('Selected video:', selectedItem);

    const updatedItems = this.state.items.map((item) => ({
      ...item,
      isSelected: item.ID === selectedItem.ID
    }));

    this.setState({ items: updatedItems, selectedMedia: selectedItem });
  };

  private adjustDiv2Height() {
    const div1 = document.getElementById('videoPlayer');
    const div2 = document.getElementById('listWrapper');

    // Check if the elements exist
    if (!div1 || !div2) return;

    // Get the height of div1
    const div1Height = div1.offsetHeight;

    // Check for mobile screens
    const isMobile = window.innerWidth <= 1080; // Adjust the threshold as needed

    if (isMobile) {
      div2.style.height = '300px';
      div2.style.maxHeight = 'none';
    } else {
      div2.style.maxHeight = div1Height + 'px';
      div2.style.height = 'auto';
    }
    // Adjust height on load
    window.onload = this.adjustDiv2Height;

    // Adjust height on window resize
    window.onresize = this.adjustDiv2Height;
  }
  
  
  
  public render(): React.ReactElement<ICollaborationMediaGalleryProps> {
    const { hasTeamsContext } = this.props;
    const { items, selectedMedia, loading, error } = this.state;


    return (
      <section className={`${styles.collaborationMediaGallery} ${hasTeamsContext ? styles.teams : ''}`}>
        {loading && <p>Loading media items, please wait...</p>}
        {error && <p className={styles.errorMessage}>{error}</p>}

        {!loading && !error && items.length === 0 && (
          <p>No media items available at the moment.</p>
        )}

        {!loading && !error && items.length > 0 && (
          <div>
            {/* Heading */}
            <div className={styles.galleryHeading} style={{ color: this.props.headingfontcolor }}>{this.props.description}</div>
            
            <div className={styles.mediaContainer}>
              <div className={styles.playerContainer} id="videoWrapper">
                {selectedMedia && <VideoPlayer url={selectedMedia.serverRelativeUrl} pURL={selectedMedia.pURL} />}
              </div>
              <div className={styles.listContainer} id={"listWrapper"}>
                <MediaList items={items} onVideoSelect={this.handleVideoSelect} />
              </div>
            </div>
          </div>
        )}
      </section>
    );
  }
}
