// import * as React from "react";
// import { useState } from "react";
// import styles from "./SideBySideView.module.scss";

// interface SideBySideViewProps {
//   newsItems: any[];
//   numberOfNews: number;
//   AuthorToggle: boolean;
//   PublishedDateToggle: boolean;
//   LikeToggle: boolean;
//   likesMap: { [url: string]: number };
//   likedMap: { [url: string]: boolean };
//   likeEnabledMap: { [url: string]: boolean };
//   toggleLike: (url: string) => void;
//   isLikesLoading: boolean;
//   newsContainerColor: string;
// }

// export const SideBySideView: React.FC<SideBySideViewProps> = ({
//   newsItems,
//   numberOfNews,
//   AuthorToggle,
//   PublishedDateToggle,
//   LikeToggle,
//   likesMap,
//   likedMap,
//   likeEnabledMap,
//   toggleLike,
//   isLikesLoading,
//   newsContainerColor,
// }) => {
//   const [isLgLayout, setIsLgLayout] = useState(false); // State to track layout mode (2-column vs 3-column)
//   const items = newsItems.slice(0, numberOfNews);

//   // Toggle layout on click
//   const handleToggleLayout = () => {
//     setIsLgLayout((prev) => !prev);
//   };

//   return (
//     <div
//       className={`${styles.sideBySideWrapper} ${isLgLayout ? styles.lgSideBySide : ""}`}
//       style={{ backgroundColor: newsContainerColor || "#fff" }}
//       onClick={handleToggleLayout} // Add click handler
//     >
//       {items.map((news) => (
//         <div className={styles.sideBySideItem} key={news.id}>
//           <div className={styles.sideBySideImageWrapper}>
//             <img
//               src={news.imageUrl || "https://via.placeholder.com/150x150.png?text=No+Image+Available"}
//               alt={news.title}
//             />
//           </div>
//           <div className={styles.sideBySideContent}>
//             <h3 className={styles.sideBySideTitle}>
//               <a href={news.url} target="_blank" rel="noopener noreferrer">
//                 {news.title}
//               </a>
//             </h3>
//             <div className={styles.sideBySideFooter}>
//               {AuthorToggle && <span className={styles.authorname}>{news.author}</span>}
//               {PublishedDateToggle && (
//                 <span className={styles.publishedDate}>
//                   {new Date(news.publishedDate).toLocaleDateString()}
//                 </span>
//               )}
//               <div className={styles.sideBySideViews}>{news.views} views</div>
//               <div className={styles.likeSection}>
//                 <button
//                   className={`${styles.likeButton} ${likedMap[news.url] ? styles.liked : ""}`}
//                   disabled={isLikesLoading}
//                   onClick={(e) => {
//                     e.stopPropagation(); // Prevent like button click from toggling layout
//                     toggleLike(news.url);
//                   }}
//                   aria-label={likedMap[news.url] ? "Unlike" : "Like"}
//                 >
//                   {likedMap[news.url] ? "❤️" : "🤍"}
//                 </button>
//                 <span>{likesMap[news.url] ?? 0}</span>
//               </div>
//             </div>
//           </div>
//         </div>
//       ))}
//     </div>
//   );
// };

import * as React from "react";
import styles from "./SideBySideView.module.scss";

interface SideBySideViewProps {
  newsItems: any[];
  numberOfNews: number;
  AuthorToggle: boolean;
  PublishedDateToggle: boolean;
  LikeToggle: boolean;
  likesMap: { [url: string]: number };
  likedMap: { [url: string]: boolean };
  likeEnabledMap: { [url: string]: boolean };
  toggleLike: (url: string) => void;
  isLikesLoading: boolean;
  newsContainerColor: string;
}

export const SideBySideView: React.FC<SideBySideViewProps> = ({
  newsItems,
  numberOfNews,
  AuthorToggle,
  PublishedDateToggle,
  LikeToggle,
  likesMap,
  likedMap,
  likeEnabledMap,
  toggleLike,
  isLikesLoading,
  newsContainerColor,
}) => {
  const items = newsItems.slice(0, numberOfNews);

  return (
    <div
      className={styles.sideBySideWrapper}
      style={{ backgroundColor: newsContainerColor || "#fff" }}
    >
      {items.map((news) => (
        <div className={styles.sideBySideCard} key={news.id}>
          <div className={styles.sideBySideImageWrapper}>
            <img
              src={news.imageUrl || "https://via.placeholder.com/150x150.png?text=No+Image+Available"}
              alt={news.title}
            />
          </div>
          <div className={styles.sideBySideContent}>
            <h3 className={styles.sideBySideTitle}>
              <a href={news.url} target="_blank" rel="noopener noreferrer">
                {news.title}
              </a>
            </h3>
            <div className={styles.sideBySideDescription}>
              {news.description || "No description available."}
            </div>
            <div className={styles.sideBySideFooter}>
              {AuthorToggle && (
                <div className={styles.authorSection}>
                  <img
                    className={styles.authorAvatar}
                    src={news.authorAvatar || "https://ui-avatars.com/api/?name=" + encodeURIComponent(news.author)}
                    alt={news.author}
                  />
                  <span className={styles.authorname}>{news.author}</span>
                </div>
              )}
              {PublishedDateToggle && (
                <span className={styles.publishedDate}>
                  {new Date(news.publishedDate).toLocaleDateString()}
                </span>
              )}
              <div className={styles.sideBySideViews}>{news.views} views</div>
              <div className={styles.likeSection}>
                <button
                  className={`${styles.likeButton} ${likedMap[news.url] ? styles.liked : ""}`}
                  disabled={isLikesLoading}
                  onClick={(e) => {
                    e.stopPropagation();
                    toggleLike(news.url);
                  }}
                  aria-label={likedMap[news.url] ? "Unlike" : "Like"}
                >
                  {likedMap[news.url] ? "❤️" : "🤍"}
                </button>
                <span>{likesMap[news.url] ?? 0}</span>
              </div>
            </div>
          </div>
        </div>
      ))}
    </div>
  );
};