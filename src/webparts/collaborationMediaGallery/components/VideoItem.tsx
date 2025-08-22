import * as React from 'react';
import styles from './VideoItem.module.scss'; 
import { IVideo } from './ICollaborationMediaGalleryProps';

interface IVideoItemProps {
  video: IVideo;
  onClick: () => void;
}

export const VideoItem: React.FunctionComponent<IVideoItemProps> = ({ video, onClick }) => {
  const thumbnailUrl= `${window.location.origin}/_api/v2.0/sharePoint:${encodeURIComponent(video.serverRelativeUrl)}:/driveItem/thumbnails/0/c240x240/content?preferNoRedirect=true`;
  
    return (
    <div
      className={`${styles.videoItem} ${video.isSelected ? styles.selected : ''}`} 
      onClick={onClick}
    >
      <div className={styles.thumbnailContainer}>
        <img src={thumbnailUrl} className={styles.thumbnail} />
      </div>
      <div className={styles.contentContainer}>
        <h4 className={styles.videoTitle}>{video.title}</h4>
        <p className={styles.videoDescription}>{video.description}</p>
      </div>
    </div>
  );
};
