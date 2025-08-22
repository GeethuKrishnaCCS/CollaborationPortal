import * as React from "react";
import styles from "./TopStoryView.module.scss";
import { TooltipHost } from '@fluentui/react';

interface TopStoryViewProps {
  newsItems: any[];
  AuthorToggle: boolean;
  PublishedDateToggle: boolean;
  numberOfNews: number;
  LikeToggle: boolean;
  likesMap: { [url: string]: number };
  likedMap: { [url: string]: boolean };
  likeEnabledMap: { [url: string]: boolean };
  toggleLike: (url: string) => void;
  isLikesLoading: boolean;
  newsContainerColor: string; 
  textColor?: string; 
}

export const TopStoryView: React.FC<TopStoryViewProps> = ({
  newsItems,
  AuthorToggle,
  PublishedDateToggle,
  numberOfNews,
  LikeToggle,
  likesMap,
  likedMap,
  likeEnabledMap,
  toggleLike,
  isLikesLoading,
  newsContainerColor,
  textColor 
}) => {
  const items = newsItems.slice(0, numberOfNews);
  

  if (items.length === 0) return null;

  const renderLike = (news: any) =>
    LikeToggle && likeEnabledMap[news.url] ? (
      <div className={styles.likeSection}>
        <button
          className={`${styles.likeButton} ${likedMap[news.url] ? styles.liked : ""}`}
          disabled={isLikesLoading}
          onClick={() => toggleLike(news.url)}
          aria-label={likedMap[news.url] ? "Unlike" : "Like"}
        >
          {likedMap[news.url] ? "❤️" : "🤍"}
        </button>
        <span>{likesMap[news.url] ?? 0}</span>
      </div>
    ) : null;

    const truncateText = (text: string, maxLength: number) => {
    if (!text) return "No description available";
    return text.length > maxLength ? text.substring(0, maxLength) + "..." : text;
  };

  return (
    <div
      className={styles.cardLayout}
      style={{ backgroundColor: newsContainerColor || "#fff" }}
    >
      <div className={styles.featuredNews}>
        <img className={styles.topimg}  src={items[0]?.imageUrl} alt={items[0]?.title} />
        <h3 className={styles.itemTitle}
        >
          <a href={items[0]?.url} target="_blank" rel="noopener noreferrer"  style={{ color: textColor|| "white"}}>
            {items[0]?.title}
          </a>
         <TooltipHost
            content={
              <div style={{ lineHeight: '1.5', color: textColor || 'black' }}>
                {/* <div>
                  <strong>Title:</strong> {items[0]?.title}
                </div> */}
                {items[0]?.Description && (
                  <div>
                    <strong>Description:</strong> {items[0]?.Description}
                  </div>
                )}
              </div>
            }
            id={`tooltip-${items[0]?.id}`}
            calloutProps={{ gapSpace: 0 }}
          >
            <h5 className={styles.description}>
              {truncateText(items[0]?.Description, 150)}
            </h5>
          </TooltipHost>
        </h3>
        <div className={styles.newsFooter}>
          {AuthorToggle && 
          <div className={styles.authorimage}>
            <img className={styles.authorIMG} src="https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRuCSSMIYXcxDjDaSU7ns9yod2mEQx9DT1jng&s" alt={items[0]?.author} />
             <div className={styles.authorInfo}>
          <span className={styles.authorname}>{items[0]?.author}</span>
          <div className={styles.newsViews}>{items[0]?.views} views</div>
        
          </div>
          </div>
          }
          {PublishedDateToggle && (
            <div className={styles.publishedDate}>
              {new Date(items[0]?.publishedDate).toLocaleDateString()}
            </div>
          )}
          {/* <div className={styles.newsViews}>{items[0]?.views} views</div> */}
          {renderLike(items[0])}
        </div>
      </div>
      <div className={styles.sideNews}>
        {items.slice(1, numberOfNews).map((news) => (
          <div key={news.id} className={styles.sideNewsItem}>
            <img className={styles.sideimg} src={news.imageUrl} alt={news.title} />
            <div className={styles.newsContent}>
              <div  className={styles.itemTitle}>
              <h3>
                <a href={news.url} target="_blank" rel="noopener noreferrer"  style={{ color: textColor|| "white"}}>
                  {news.title}
                </a>
               <TooltipHost
            content={
              <div style={{ lineHeight: '1.5', color: textColor || 'black' }}>
                {/* <div>
                  <strong>Title:</strong> {items[0]?.title}
                </div> */}
                {news?.Description && (
                  <div>
                    <strong>Description:</strong> {news?.Description}
                  </div>
                )}
              </div>
            }
            id={`tooltip-${news?.id}`}
            calloutProps={{ gapSpace: 0 }}
          >
            <h5 className={styles.description}>
              {truncateText(news?.Description, 60)}
            </h5>
          </TooltipHost>
              </h3>
              </div>
              <div className={styles.newsFooter}>
                {AuthorToggle &&
                <div className={styles.authorimage}>
                <img className={styles.authorIMG} src="https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRuCSSMIYXcxDjDaSU7ns9yod2mEQx9DT1jng&s" alt={news.author} />
                <div className={styles.authorInfo}>
                <>
                <span className={styles.authorname}>{news.author}</span>
                <div className={styles.newsViews}>{news.views} views</div>
                </>
                </div>
                </div>
                }
                {PublishedDateToggle && (
                  <span className={styles.publishedDate}>
                    {new Date(news.publishedDate).toLocaleDateString()}
                  </span>
                )}
                {/* <div className={styles.newsViews}>{news.views} views</div> */}
                {renderLike(news)}
              </div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};