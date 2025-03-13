import * as React from 'react';
import styles from './FilmstripView.module.scss';

interface IFilmstripViewProps {
    events: any[];
    location?: string;
    stylesForImages: string;
}

const FilmstripView: React.FC<IFilmstripViewProps> = ({ events, location, stylesForImages }) => {
    return (
        <div className={styles.StackStyle}>
            <div className={styles.StackStyleContainer}>
                <div className={styles.BirthdaySlider}>
                    {events
                        .filter((event: any) => !location || event.Location === location)
                        .map((event: any) => {
                            const eventImageUrl = event.EventImage ? event.EventImage.split('?path=')[1] : "https://via.placeholder.com/150";
                            return (
                                <div className={styles.FilmstripCard} key={event.Id}>

                                    {/* image- circle */}
                                    {stylesForImages === "Circle" &&
                                        <>
                                            <div className={styles.Circledate}>
                                                {new Date(event.StartTime).toLocaleDateString("en-US", { month: "short", day: "numeric" })}
                                            </div>

                                            <div className={styles.imageContainer}>
                                                <img
                                                    className={`${styles.profileImage} ${styles.circleImageCard}`}
                                                    src={event.EventImage ? `https://ccsdev01.sharepoint.com${eventImageUrl}` : require('../assets/DefaultImage.png')}
                                                    alt={event.Title}
                                                />

                                            </div>
                                        </>
                                    }

                                    {/* image- rectangle */}
                                    {stylesForImages != "Circle" &&

                                        <div className={styles.imageContainer}>
                                            <img
                                                className={`${styles.profileImage} ${stylesForImages === "Square"
                                                    ? styles.squareImageCard
                                                    : stylesForImages === "Rectangle"
                                                        ? styles.rectangleImageCard
                                                        : ""
                                                    }`}
                                                src={event.EventImage ? `https://ccsdev01.sharepoint.com${eventImageUrl}` : require('../assets/DefaultImage.png')}
                                                alt={event.Title}
                                            />
                                            <div className={styles.date}>
                                                {new Date(event.StartTime).toLocaleDateString("en-US", { month: "short", day: "numeric" })}
                                            </div>
                                        </div>
                                    }


                                    <div className={styles.cardContent}>
                                        <span className={styles.category}>{event.Category}</span>
                                        <h3 className={styles.title}>{event.Title}</h3>
                                        <p className={styles.dateTime}>
                                            {new Date(event.StartTime).toLocaleString("en-US", {
                                                weekday: "short",
                                                month: "short",
                                                day: "numeric",
                                                hour: "2-digit",
                                                minute: "2-digit",
                                            })}
                                        </p>
                                        <p>{event.Location}</p>
                                        <div className={styles.actionButtons}>
                                            <img src={require('../assets/RSVP.png')} className={styles.rsvpButton} />
                                            <img src={require('../assets/Mail.png')} className={styles.mailIcon} />
                                            <img src={require('../assets/Comment.png')} className={styles.commentIcon} />
                                            <img src={require('../assets/Like.png')} className={styles.likeIcon} />
                                        </div>
                                    </div>
                                </div>
                            );
                        })}
                </div>
            </div>
        </div>
    );
};

export default FilmstripView;