import * as React from 'react';
import ReactPlayer from 'react-player';
import styles from './VideoPlayer.module.scss';

export interface IVideoPlayerProps {
  url: string;
  playing?: boolean;
  pURL?: string
}

export const VideoPlayer: React.FunctionComponent<IVideoPlayerProps> = ({ url, playing = true, pURL }) => {
  const playButton = (): JSX.Element => {
    return (
      <svg width="80" height="80" viewBox="0 0 80 80" fill="none" xmlns="http://www.w3.org/2000/svg">
        <circle cx="40" cy="40" r="40" fill="black" fill-opacity="0.6" />
        <path d="M49.325 38.5562L31.925 28.2C31.7125 28.075 31.4938 28 31.2438 28C30.5625 28 30.0062 28.5625 30.0062 29.25H30V50.75H30.0062C30.0062 51.4375 30.5625 52 31.2438 52C31.5 52 31.7125 51.9125 31.9437 51.7875L49.325 41.4438C49.7375 41.1 50 40.5813 50 40C50 39.4187 49.7375 38.9062 49.325 38.5562Z" fill="white" />
      </svg>
    )
  }

  return (
    <div className={styles.videoplayer} key={url} id={"videoPlayer"}>
      {url ? (
        <ReactPlayer
          light={pURL ? pURL : true}
          url={url}
          controls
          playing={playing}
          width="100%"
          height="100%"
          style={{ borderRadius: '15px', overflow: 'hidden' }}
          playIcon={playButton()}
        />
      ) : (
        <p>No video URL provided.</p>
      )}
    </div>
  );
};
