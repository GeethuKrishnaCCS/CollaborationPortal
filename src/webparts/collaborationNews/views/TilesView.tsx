import * as React from "react";
import styles from "./TilesView.module.scss";

interface TilesViewProps {
  newsItems: any[];
  numberOfNews: number;
}

export const TilesView: React.FC<TilesViewProps> = ({ newsItems, numberOfNews }) => {
  return (
    <div className={styles.tilesWrapper}>
      {newsItems.slice(0, numberOfNews).map((news) => (
        <div key={news.id} className={styles.tileCard}>
          <div className={styles.tileImageWrapper}>
            <img
              src={news.imageUrl || "https://via.placeholder.com/300"}
              alt={news.title}
              className={styles.tileImage}
              onError={(e) => (e.currentTarget.src = "https://via.placeholder.com/300")}
            />
          </div>
          <div className={styles.tileContent}>
            <h3 className={styles.tileTitle}>{news.title}</h3>
            <div className={styles.tileViews}>
              👁️ {news.views} views
            </div>
          </div>
        </div>
      ))}
    </div>
  );
};