import * as React from 'react';
import styles from './CollaborationNews.module.scss';
import type { ICollaborationNewsProps } from './ICollaborationNewsProps';
import { useState, useEffect, useMemo, memo } from "react";
import { CollaborationNewsService, NewsCommentItem, getSiteUrlFromNewsUrl } from "../services/CollaborationNewsService";
import { NewsHeader } from "../views/NewsHeader";
import { ListView } from "../views/ListView";
import { TopStoryView } from "../views/TopStoryView";
import { SideBySideView } from "../views/SideBySideView";
import { HubNewsView } from "../views/HubNewsView";
import { CarouselView } from "../views/CarouselView";
import { TilesView } from "../views/TilesView";

import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";
import { CollaborationNewsModel } from '../models/CollaborationNewsModel';

const handleAddNewsPost = (newsUrl: string): void => {
  const siteUrl = getSiteUrlFromNewsUrl(newsUrl);
  window.location.href = `${siteUrl}/_layouts/15/CreatePageFromTemplate.aspx?source=FromWebpart&promotedState=1`;
};

const handleAddNewsLink = (): void => {
  console.log("handleAddNewsLink: Clicked");
};

const debounce = <T extends (...args: any[]) => any>(func: T, wait: number): ((...args: Parameters<T>) => void) => {
  let timeout:number;
  return (...args: Parameters<T>) => {
    clearTimeout(timeout);
    timeout = setTimeout(() => func(...args), wait);
  };
};

const CollaborationNews : React.FC<ICollaborationNewsProps> =({
   StyleToggle = "list",
  AuthorToggle,
  PublishedDateToggle,
  LikeToggle,
  CommentToggle,
  selectedSites = [],
  selectedCategories = [],
  selectedGroups = [],
  numberOfNews,
  context,
  placeholderImageUrl = "https://via.placeholder.com/150x150.png?text=No+Image+Available",
  headingColor,
  backgroundColor,
  customTitle,
  // canAddNews = false,
  newsContainerColor,
  initialNews = [],
  buttonBackgroundColor,
  textColor 
 }) => {

  const [updatedNewsItems, setUpdatedNewsItems] = useState<CollaborationNewsModel[]>(initialNews);
  const [likesMap, setLikesMap] = useState<{ [url: string]: number }>({});
  const [likedMap, setLikedMap] = useState<{ [url: string]: boolean }>({});
  const [likeEnabledMap, setLikeEnabledMap] = useState<{ [url: string]: boolean }>({});
  const [activePopupUrl, setActivePopupUrl] = useState<string | null>(null);
  const [commentsMap, setCommentsMap] = useState<{ [url: string]: NewsCommentItem[] }>({});
  const [commentInputs, setCommentInputs] = useState<{ [url: string]: string }>({});
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);

  const newsService = useMemo(() => new CollaborationNewsService(context), [context]);
  const styleClass = StyleToggle.toLowerCase();

  const fetchData = async (): Promise<void> => {
    setIsLoading(true);
    setError(null);
    try {
      if (selectedSites.length === 0) {
        setError("No sites selected.");
        return;
      }
      console.log("fetchData: Fetching:", { selectedSites, selectedCategories, selectedGroups });
      const cacheKey = `news_${JSON.stringify(selectedSites)}_${context.pageContext.user.email?.toLowerCase()}_${JSON.stringify(
        selectedGroups.map((g) => g.id)
      )}_${JSON.stringify(selectedCategories)}`;
      const cacheTimestampKey = `${cacheKey}_timestamp`;
      const cachedNews = localStorage.getItem(cacheKey);
      const cachedTimestamp = localStorage.getItem(cacheTimestampKey);
      const cacheAge = cachedTimestamp ? Date.now() - parseInt(cachedTimestamp) : Infinity;
      const cacheExpiration = 5 * 60 * 1000; // 5 minutes

      if (cachedNews && cacheAge < cacheExpiration) {
        console.log("fetchData: Using cache:", JSON.parse(cachedNews).length);
        setUpdatedNewsItems(JSON.parse(cachedNews));
        setIsLoading(false);
        return;
      }

      const newsItems = await newsService.getNewsFromSites(selectedSites, selectedCategories, selectedGroups);
      console.log("Raw news items:", newsItems);
      const slicedItems = numberOfNews ? newsItems.slice(0, numberOfNews) : newsItems;
      const updatedItems = slicedItems.map((item:any) => ({
        ...item,
        imageUrl: item.imageUrl || placeholderImageUrl,
        
      }));
      setUpdatedNewsItems(updatedItems);
      localStorage.setItem(cacheKey, JSON.stringify(updatedItems));
      localStorage.setItem(cacheTimestampKey, Date.now().toString());
      console.log("fetchData: Fetched:", updatedItems.length);
    
    } catch (error) {
      console.error("fetchData: Error:", error);
      setError("Failed to load news.");
    } finally {
      setIsLoading(false);
    }
  };

  const fetchLikes = async (): Promise<void> => {
    try {
      const currentUserEmail = context.pageContext.user.email?.toLowerCase() || "";
      const { likes, liked, likeEnabled } = await newsService.fetchLikesForNews(updatedNewsItems, currentUserEmail);
      setLikesMap(likes);
      setLikedMap(liked);
      setLikeEnabledMap(likeEnabled);
    } catch (err) {
      console.error("fetchLikes: Error:", err);
      setError("Failed to load likes.");
      setTimeout(() => setError(null), 3000);
    }
  };

 useEffect(() => {
    console.log("useEffect: Props:", {
      initialNews: initialNews.length,
      selectedSites,
      selectedCategories,
      // canAddNews,
      hasHandlers: !!handleAddNewsPost && !!handleAddNewsLink,
    
    });
    if (initialNews.length > 0) {
      setUpdatedNewsItems(initialNews);
      setIsLoading(false);
    } else {
      void fetchData();
    }
  }, [newsService, selectedSites, selectedCategories, numberOfNews, placeholderImageUrl, selectedGroups,]);

  useEffect(() => {
    if (!LikeToggle || updatedNewsItems.length === 0) {
      setLikesMap({});
      setLikedMap({});
      setLikeEnabledMap({});
      return;
    }
    void fetchLikes();
  }, [LikeToggle, updatedNewsItems, newsService, context.pageContext.user.email]);

  const toggleLike = debounce(async (url: string): Promise<void> => {
    const currentLikeCount = likesMap[url] || 0;
    const currentLikedStatus = likedMap[url] || false;

    try {
      setLikesMap((prev) => ({ ...prev, [url]: currentLikedStatus ? currentLikeCount - 1 : currentLikeCount + 1 }));
      setLikedMap((prev) => ({ ...prev, [url]: !currentLikedStatus }));
      await newsService.toggleLike(url, currentLikedStatus);
      const updatedLikes = await newsService.fetchLikesForNews([{ url }], context.pageContext.user.email?.toLowerCase() || "");
      setLikesMap((prev) => ({ ...prev, [url]: updatedLikes.likes[url] }));
      setLikedMap((prev) => ({ ...prev, [url]: updatedLikes.liked[url] }));
    } catch (error) {
      console.error("toggleLike: Error:", error);
      setError("Failed to update like.");
      setTimeout(() => setError(null), 3000);
      setLikesMap((prev) => ({ ...prev, [url]: currentLikeCount }));
      setLikedMap((prev) => ({ ...prev, [url]: currentLikedStatus }));
    }
  }, 300);

  const loadComments = async (newsUrl: string): Promise<void> => {
    try {
      const comments = await newsService.fetchComments(newsUrl);
      setCommentsMap((prevMap) => ({ ...prevMap, [newsUrl]: comments }));
    } catch (error) {
      console.error("loadComments: Error:", error);
      setError("Failed to load comments.");
      setTimeout(() => setError(null), 3000);
    }
  };

  const togglePopup = (url: string): void => {
    setActivePopupUrl((prevUrl) => (prevUrl === url ? null : url));
    if (!commentsMap[url]) {
      void loadComments(url);
    }
  };

  const handlePostComment = async (newsUrl: string): Promise<void> => {
    const commentText = commentInputs[newsUrl];
    if (!commentText) {
      setError("Please write a comment.");
      setTimeout(() => setError(null), 3000);
      return;
    }
    try {
      await newsService.postComment(commentText, newsUrl);
      setCommentInputs((prevMap) => ({ ...prevMap, [newsUrl]: "" }));
      await loadComments(newsUrl);
    } catch (error) {
      console.error("handlePostComment: Error:", error);
      setError("Failed to post comment.");
      setTimeout(() => setError(null), 3000);
    }
  };

  if (isLoading) {
    return <div className={styles.newsContainer}>Loading...</div>;
  }

  if (error) {
    return <div className={styles.newsContainer}>Error: {error}</div>;
  }

      return (
    <div className={styles.newsContainer} style={{ backgroundColor: backgroundColor || "transparent" }}>
      <NewsHeader
        newsItems={updatedNewsItems}
        handleAddNewsPost={handleAddNewsPost}
        handleAddNewsLink={handleAddNewsLink}
        headingColor={headingColor}

        customTitle={customTitle}
        // canAddNews={canAddNews}
        context={context}
        selectedSites={selectedSites}
        buttonBackgroundColor={buttonBackgroundColor}
      />
      {updatedNewsItems.length === 0 && <p>No news available for selected sites or categories.</p>}
      {styleClass === "list" && updatedNewsItems.length > 0 && (
        <ListView
          newsItems={updatedNewsItems}
          numberOfNews={numberOfNews}
          AuthorToggle={AuthorToggle}
          PublishedDateToggle={PublishedDateToggle}
          LikeToggle={LikeToggle}
          CommentToggle={CommentToggle}
          likesMap={likesMap}
          likedMap={likedMap}
          likeEnabledMap={likeEnabledMap}
          commentsMap={commentsMap}
          commentInputs={commentInputs}
          activePopupUrl={activePopupUrl}
          toggleLike={toggleLike}
          togglePopup={togglePopup}
          loadComments={loadComments}
          setCommentInputs={setCommentInputs}
          handlePostComment={handlePostComment}
          isLikesLoading={false}
          placeholderImageUrl={placeholderImageUrl}
          newsContainerColor={newsContainerColor}
        />
      )}
      {styleClass === "topstory" && updatedNewsItems.length > 0 && (
        <TopStoryView
          newsItems={updatedNewsItems}
          numberOfNews={numberOfNews}
          AuthorToggle={AuthorToggle}
          PublishedDateToggle={PublishedDateToggle}
          LikeToggle={LikeToggle}
          likesMap={likesMap}
          likedMap={likedMap}
          likeEnabledMap={likeEnabledMap}
          toggleLike={toggleLike}
          isLikesLoading={false}
          newsContainerColor={newsContainerColor}
          textColor={textColor}
        />
      )}
      {styleClass === "sidebyside" && updatedNewsItems.length > 0 && (
        <SideBySideView
          newsItems={updatedNewsItems}
          numberOfNews={numberOfNews}
          AuthorToggle={AuthorToggle}
          PublishedDateToggle={PublishedDateToggle}
          LikeToggle={LikeToggle}
          likesMap={likesMap}
          likedMap={likedMap}
          likeEnabledMap={likeEnabledMap}
          toggleLike={toggleLike}
          isLikesLoading={false}
          newsContainerColor={newsContainerColor}
        />
      )}
      {styleClass === "hubnews" && updatedNewsItems.length > 0 && (
        <HubNewsView newsItems={updatedNewsItems} AuthorToggle={AuthorToggle} PublishedDateToggle={PublishedDateToggle} />
      )}
      {styleClass === "carousel" && updatedNewsItems.length > 0 && <CarouselView newsItems={updatedNewsItems} />}
      {styleClass === "tiles" && updatedNewsItems.length > 0 && (
        <TilesView newsItems={updatedNewsItems} numberOfNews={numberOfNews} />
      )}
    </div>
  );
  }
  const arePropsEqual = (prevProps: ICollaborationNewsProps, nextProps: ICollaborationNewsProps): boolean => {
  const compareArrays = (arr1: string[], arr2: string[]): boolean =>
    arr1.length === arr2.length && arr1.every((item, index) => item === arr2[index]);
  const compareGroups = (groups1: IPropertyFieldGroupOrPerson[], groups2: IPropertyFieldGroupOrPerson[]): boolean =>
    groups1.length === groups2.length &&
    groups1.every((group, index) => group.id === groups2[index].id && group.fullName === groups2[index].fullName);
  const compareNews = (news1: CollaborationNewsModel[], news2: CollaborationNewsModel[]): boolean =>
    news1.length === news2.length && news1.every((item, index) => item.id === news2[index].id && item.url === news2[index].url);

  return (
    prevProps.StyleToggle === nextProps.StyleToggle &&
    prevProps.AuthorToggle === nextProps.AuthorToggle &&
    prevProps.PublishedDateToggle === nextProps.PublishedDateToggle &&
    prevProps.LikeToggle === nextProps.LikeToggle &&
    prevProps.CommentToggle === nextProps.CommentToggle &&
    compareArrays(prevProps.selectedSites || [], nextProps.selectedSites || []) &&
    compareArrays(prevProps.selectedCategories || [], nextProps.selectedCategories || []) &&
    compareGroups(prevProps.selectedGroups || [], nextProps.selectedGroups || []) &&
    prevProps.numberOfNews === nextProps.numberOfNews &&
    prevProps.context === nextProps.context &&
    prevProps.placeholderImageUrl === nextProps.placeholderImageUrl &&
    prevProps.headingColor === nextProps.headingColor &&
    prevProps.backgroundColor === nextProps.backgroundColor &&
    prevProps.newsContainerColor === nextProps.newsContainerColor &&
    prevProps.customTitle === nextProps.customTitle &&
    // prevProps.canAddNews === nextProps.canAddNews &&
    compareNews(prevProps.initialNews || [], nextProps.initialNews || []) &&
    prevProps.buttonBackgroundColor === nextProps.buttonBackgroundColor
  );
}
export default memo(CollaborationNews, arePropsEqual);
