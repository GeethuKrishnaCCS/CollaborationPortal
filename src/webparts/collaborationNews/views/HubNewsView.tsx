import * as React from "react";
import styles from "./HubNewsView.module.scss";

interface HubNewsViewProps {
  newsItems: any[];
  AuthorToggle: boolean;
  PublishedDateToggle: boolean;
}

export const HubNewsView: React.FC<HubNewsViewProps> = ({
  newsItems,
  AuthorToggle,
  PublishedDateToggle,
}) => {
  const mainStory = newsItems[0];
  const secondaryStories = newsItems.slice(1, 5);

  return (
    <div className={styles.hubNewsContainer}>
      {mainStory && (
        <div className={styles.mainStory}>
          <div className={styles.mainImageWrapper}>
            <img
              src={mainStory.imageUrl || "https://via.placeholder.com/600x400"}
              alt={mainStory.title}
              className={styles.mainImage}
              onError={(e) => (e.currentTarget.src = "https://via.placeholder.com/600x400")}
            />
          </div>
          <div className={styles.mainContent}>
            <h2 className={styles.mainTitle}>
              <a href={mainStory.url} target="_blank" rel="noopener noreferrer">
                {mainStory.title}
              </a>
            </h2>
            {mainStory.categories && mainStory.categories.length > 0 && (
              <div className={styles.newsCategory}>{mainStory.categories.join(", ")}</div>
            )}
            <div className={styles.mainMeta}>
              {AuthorToggle && <span>{mainStory.author}</span>}
              {PublishedDateToggle && (
                <span>{new Date(mainStory.publishedDate).toLocaleDateString()}</span>
              )}
              <span className={styles.views}>👁️ {mainStory.views} views</span>
            </div>
          </div>
        </div>
      )}
      <div className={styles.secondaryStories}>
        {secondaryStories.map((news) => (
          <div key={news.id} className={styles.secondaryItem}>
            <div className={styles.secondaryImageWrapper}>
              <img
                src={news.imageUrl || "https://via.placeholder.com/150"}
                alt={news.title}
                className={styles.secondaryImage}
                onError={(e) => (e.currentTarget.src = "https://via.placeholder.com/150")}
              />
            </div>
            <div className={styles.secondaryContent}>
              <h3 className={styles.secondaryTitle}>
                <a href={news.url} target="_blank" rel="noopener noreferrer">
                  {news.title}
                </a>
              </h3>
              {news.categories && news.categories.length > 0 && (
                <div className={styles.newsCategory}>{news.categories.join(", ")}</div>
              )}
              <div className={styles.secondaryMeta}>
                {AuthorToggle && <span>{news.author}</span>}
                {PublishedDateToggle && (
                  <span>{new Date(news.publishedDate).toLocaleDateString()}</span>
                )}
                <span className={styles.views}>👁️ {mainStory.views} views</span>
              </div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};