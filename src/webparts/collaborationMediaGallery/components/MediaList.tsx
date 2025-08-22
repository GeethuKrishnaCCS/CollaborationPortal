import * as React from 'react';
import { VideoItem } from './VideoItem';
import styles from './MediaList.module.scss';
import { IVideo } from './ICollaborationMediaGalleryProps';

export interface IMediaListProps {
    items: IVideo[];
    onVideoSelect: (item: IVideo) => void;
}

export const MediaList: React.FunctionComponent<IMediaListProps> = (props: React.PropsWithChildren<IMediaListProps>) => {
    const { items, onVideoSelect } = props;
    const [currentIndex, setCurrentIndex] = React.useState(0);

    const itemsPerPage = 4; // Show 4 items per page

    const visibleItems = items.slice(currentIndex, currentIndex + itemsPerPage);

    const handleVideoSelect = (item: IVideo): void => {
        onVideoSelect(item);
    };

    const handleNext = () => {
        if (currentIndex + itemsPerPage < items.length) {
            setCurrentIndex(currentIndex + itemsPerPage);
        }
    };

    const handlePrev = () => {
        if (currentIndex - itemsPerPage >= 0) {
            setCurrentIndex(currentIndex - itemsPerPage);
        }
    };

    const prevButton = () => {
        return (
            <svg xmlns="http://www.w3.org/2000/svg" height='30' width='30' viewBox="0 0 24 24" fill="currentColor" >
                <path fill-rule="evenodd" d="M11.47 7.72a.75.75 0 0 1 1.06 0l7.5 7.5a.75.75 0 1 1-1.06 1.06L12 9.31l-6.97 6.97a.75.75 0 0 1-1.06-1.06l7.5-7.5Z" clip-rule="evenodd" />
            </svg>
        )
    }
    const nextButton = () => {
        return (

            <svg xmlns="http://www.w3.org/2000/svg" height='30' width='30' viewBox="0 0 24 24" fill="currentColor" >
            <path fill-rule="evenodd" d="M12.53 16.28a.75.75 0 0 1-1.06 0l-7.5-7.5a.75.75 0 0 1 1.06-1.06L12 14.69l6.97-6.97a.75.75 0 1 1 1.06 1.06l-7.5 7.5Z" clip-rule="evenodd" />
          </svg>
        )
    }

    return (
        <div className={styles.mediaListContainer}>
            {/* Previous button */}
            <button
                className={styles.prevButton}
                onClick={handlePrev}
                disabled={currentIndex === 0}
            >
                {prevButton()}
            </button>

            <div className={styles.mediaList}>
                {visibleItems.map((item, index) => (
                    <VideoItem
                        key={index}
                        video={item}
                        onClick={() => handleVideoSelect(item)}
                    />
                ))}
            </div>

            {/* Next button */}
            <button
                className={styles.nextButton}
                onClick={handleNext}
                disabled={currentIndex + itemsPerPage >= items.length}
            >
                {nextButton()}
            </button>
        </div>
    );
};
