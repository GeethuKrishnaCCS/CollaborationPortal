import * as React from "react";
import { memo, useCallback } from "react";
import styles from "./ListView.module.scss";
import { CommentPopup } from "./CommentPopup";
import { CollaborationNewsModel} from "../models/CollaborationNewsModel";
import { NewsCommentItem } from "../services/CollaborationNewsService";

// Default fallback image using placehold.co
const DEFAULT_FALLBACK_IMAGE = "https://placehold.co/150x150?text=No+Image";

interface ListViewProps {
  newsItems: CollaborationNewsModel[];
  numberOfNews: number;
  AuthorToggle: boolean;
  PublishedDateToggle: boolean;
  LikeToggle: boolean;
  CommentToggle: boolean;
  likesMap: { [url: string]: number };
  likedMap: { [url: string]: boolean };
  likeEnabledMap: { [url: string]: boolean };
  commentsMap: { [url: string]: NewsCommentItem[] };
  commentInputs: { [url: string]: string };
  activePopupUrl: string | null;
  toggleLike: (url: string) => void;
  togglePopup: (url: string) => void;
  loadComments: (url: string) => void;
  setCommentInputs: React.Dispatch<React.SetStateAction<{ [url: string]: string }>>;
  handlePostComment: (url: string) => void;
  isLikesLoading: boolean;
  placeholderImageUrl: string;
  newsContainerColor: string; // Added for news items container background
}

// Memoized NewsItem component
const NewsItem: React.FC<{
  news: CollaborationNewsModel;
  AuthorToggle: boolean;
  PublishedDateToggle: boolean;
  LikeToggle: boolean;
  CommentToggle: boolean;
  likes: number;
  liked: boolean;
  likeEnabled: boolean;
  comments: NewsCommentItem[];
  commentInput: string;
  activePopup: boolean;
  toggleLike: (url: string) => void;
  togglePopup: (url: string) => void;
  loadComments: (url: string) => void;
  setCommentInput: (url: string, value: string) => void;
  handlePostComment: (url: string) => void;
  isLikesLoading: boolean;
  placeholderImageUrl: string;
}> = memo(
  ({
    news,
    AuthorToggle,
    PublishedDateToggle,
    LikeToggle,
    CommentToggle,
    likes,
    liked,
    likeEnabled,
    comments,
    commentInput,
    activePopup,
    toggleLike,
    togglePopup,
    loadComments,
    setCommentInput,
    handlePostComment,
    isLikesLoading,
    placeholderImageUrl,
  }) => {
    // Function to clean and validate image URLs
    const cleanImageUrl = (url: string | undefined): string => {
      if (!url) return placeholderImageUrl;

      // Remove extra quotes or malformed wrapping
      const cleanUrl = url.replace(/['"]/g, "");
      // Extract valid URL if embedded (e.g., 'http://...')
      const match = cleanUrl.match(/https?:\/\/[^\s<>"']+/);
      if (match) return match[0];

      console.warn(`Invalid image URL: ${url}`);
      return placeholderImageUrl;
    };

    // Handle image load errors
    const handleImageError = (e: React.SyntheticEvent<HTMLImageElement>, title: string) => {
      console.error(`Image failed to load for ${title}: ${e.currentTarget.src}`);
      if (e.currentTarget.src !== placeholderImageUrl) {
        e.currentTarget.src = placeholderImageUrl;
      } else {
        e.currentTarget.src = DEFAULT_FALLBACK_IMAGE;
      }
    };

    console.log(`NewsItem rendering: ${news.url}`); // Debug render

    return (
      <div className={styles.newsItem}>
        <div className={styles.newsImage}>
          <img
            src={cleanImageUrl(news.imageUrl)}
            alt={news.title}
            onError={(e) => handleImageError(e, news.title)}
            style={{ width: "150px", height: "150px" }}
          />
        </div>
        <div className={styles.newsContent}>
          <h3 className={styles.itemTitle}>
            <a href={news.url} target="_blank" rel="noopener noreferrer">
              {news.title}
            </a>
          </h3>
          {news.categories && news.categories.length > 0 && (
            <div className={styles.newsCategory}>{news.categories.join(", ")}</div>
          )}
          <div className={styles.newsFooter}>
            {AuthorToggle && <span className={styles.authorname}>{news.author}</span>}
            {PublishedDateToggle && (
              <span className={styles.publishedDate}>{new Date(news.publishedDate).toLocaleDateString()}</span>
            )}
            <div className={styles.newsViews}>{news.views} views</div>
            <div className={styles.actionsSection}>
              {LikeToggle && likeEnabled && (
                <div className={styles.likeSection}>
                  <button
                    className={liked ? styles.liked : styles.likeButton}
                    onClick={() => {
                      console.log(`Like button clicked for ${news.url}`);
                      toggleLike(news.url);
                    }}
                  >
                    {liked ? "❤️" : "🤍"}
                  </button>
                  <span>{likes ?? 0} Likes</span>
                </div>
              )}
              {CommentToggle && (
                <div className={styles.likeSection}>
                  <button
                    onClick={() => {
                      togglePopup(news.url);
                      if (!activePopup) {
                        loadComments(news.url);
                      }
                    }}
                    className={styles.commentButton}
                  >
                    💬
                  </button>
                </div>
              )}
            </div>
          </div>
          {CommentToggle && activePopup && (
            <CommentPopup
              newsUrl={news.url}
              comments={comments}
              commentInput={commentInput}
              setCommentInput={(value) => setCommentInput(news.url, value)}
              handlePostComment={handlePostComment}
              closePopup={() => togglePopup("")}
            />
          )}
        </div>
      </div>
    );
  },
  (prevProps, nextProps) => {
    const result =
      prevProps.news === nextProps.news &&
      prevProps.AuthorToggle === nextProps.AuthorToggle &&
      prevProps.PublishedDateToggle === nextProps.PublishedDateToggle &&
      prevProps.LikeToggle === nextProps.LikeToggle &&
      prevProps.CommentToggle === nextProps.CommentToggle &&
      prevProps.likes === nextProps.likes &&
      prevProps.liked === nextProps.liked &&
      prevProps.likeEnabled === nextProps.likeEnabled &&
      prevProps.comments === nextProps.comments &&
      prevProps.commentInput === nextProps.commentInput &&
      prevProps.activePopup === nextProps.activePopup &&
      prevProps.toggleLike === nextProps.toggleLike &&
      prevProps.togglePopup === nextProps.togglePopup &&
      prevProps.loadComments === nextProps.loadComments &&
      prevProps.setCommentInput === nextProps.setCommentInput &&
      prevProps.handlePostComment === nextProps.handlePostComment &&
      prevProps.isLikesLoading === nextProps.isLikesLoading &&
      prevProps.placeholderImageUrl === nextProps.placeholderImageUrl;

    console.log(`NewsItem memo check for ${prevProps.news.url}:`, { result });
    return result;
  }
);

export const ListView: React.FC<ListViewProps> = memo(
  ({
    newsItems,
    numberOfNews,
    AuthorToggle,
    PublishedDateToggle,
    LikeToggle,
    CommentToggle,
    likesMap,
    likedMap,
    likeEnabledMap,
    commentsMap,
    commentInputs,
    activePopupUrl,
    toggleLike,
    togglePopup,
    loadComments,
    setCommentInputs,
    handlePostComment,
    isLikesLoading,
    placeholderImageUrl,
    newsContainerColor,
  }) => {
    console.log("ListView rendering:", {
      LikeToggle,
      likeEnabledMap,
      likesMap,
      likedMap,
      isLikesLoading,
      newsUrls: newsItems.map((news) => news.url),
      newsContainerColor,
    });

    // Memoized handler for comment input changes
    const handleCommentInputChange = useCallback(
      (url: string, value: string) => {
        setCommentInputs((prev: { [url: string]: string }) => ({ ...prev, [url]: value }));
      },
      [setCommentInputs]
    );

    return (
      <div className={styles.listView} style={{ backgroundColor: newsContainerColor || "#ffffff" }}>
        {newsItems.slice(0, numberOfNews).map((news) => (
          <NewsItem
            key={news.url}
            news={news}
            AuthorToggle={AuthorToggle}
            PublishedDateToggle={PublishedDateToggle}
            LikeToggle={LikeToggle}
            CommentToggle={CommentToggle}
            likes={likesMap[news.url] || 0}
            liked={likedMap[news.url] || false}
            likeEnabled={likeEnabledMap[news.url] || false}
            comments={commentsMap[news.url] || []}
            commentInput={commentInputs[news.url] || ""}
            activePopup={activePopupUrl === news.url}
            toggleLike={toggleLike}
            togglePopup={togglePopup}
            loadComments={loadComments}
            setCommentInput={handleCommentInputChange}
            handlePostComment={handlePostComment}
            isLikesLoading={isLikesLoading}
            placeholderImageUrl={placeholderImageUrl}
          />
        ))}
      </div>
    );
  },
  (prevProps, nextProps) => {
    // Check if newsItems array has changed (reference or content)
    const newsItemsUnchanged =
      prevProps.newsItems === nextProps.newsItems ||
      (prevProps.newsItems.length === nextProps.newsItems.length &&
        prevProps.newsItems.every((item, index) => item.url === nextProps.newsItems[index].url));

    // Check if likesMap or likedMap changed for any news item
    const likesUnchanged = prevProps.newsItems.every((item) => {
      const url = item.url;
      return (
        prevProps.likesMap[url] === nextProps.likesMap[url] &&
        prevProps.likedMap[url] === nextProps.likedMap[url] &&
        prevProps.likeEnabledMap[url] === nextProps.likeEnabledMap[url]
      );
    });

    // Check if commentsMap or commentInputs changed for any news item
    const commentsUnchanged = prevProps.newsItems.every((item) => {
      const url = item.url;
      return (
        prevProps.commentsMap[url] === nextProps.commentsMap[url] &&
        prevProps.commentInputs[url] === nextProps.commentInputs[url]
      );
    });

    const result =
      newsItemsUnchanged &&
      likesUnchanged &&
      commentsUnchanged &&
      prevProps.numberOfNews === nextProps.numberOfNews &&
      prevProps.AuthorToggle === nextProps.AuthorToggle &&
      prevProps.PublishedDateToggle === nextProps.PublishedDateToggle &&
      prevProps.LikeToggle === nextProps.LikeToggle &&
      prevProps.CommentToggle === nextProps.CommentToggle &&
      prevProps.activePopupUrl === nextProps.activePopupUrl &&
      prevProps.toggleLike === nextProps.toggleLike &&
      prevProps.togglePopup === nextProps.togglePopup &&
      prevProps.loadComments === nextProps.loadComments &&
      prevProps.setCommentInputs === nextProps.setCommentInputs &&
      prevProps.handlePostComment === nextProps.handlePostComment &&
      prevProps.isLikesLoading === nextProps.isLikesLoading &&
      prevProps.placeholderImageUrl === nextProps.placeholderImageUrl &&
      prevProps.newsContainerColor === nextProps.newsContainerColor; // Added

    console.log("ListView memo check:", {
      newsItemsUnchanged,
      likesUnchanged,
      commentsUnchanged,
      result,
    });

    return result;
  }
);