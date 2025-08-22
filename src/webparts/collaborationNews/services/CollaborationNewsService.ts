import { PnPConfig } from "./PnPConfig";
import { CollaborationNewsModel } from "../models/CollaborationNewsModel";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/behaviors/spfx";
import "@pnp/sp/search";

export interface NewsCommentItem {
  Title: string;
  NewsURL: string;
  CreatedOn: string;
  CreatedBy: { Title: string };
}

interface NewsView {
  title: string;
  url: string;
  views: number;
}

interface LikeInfo {
  likes: { [url: string]: number };
  liked: { [url: string]: boolean };
  likeEnabled: { [url: string]: boolean };
}

interface PageComment {
  text: string;
  createdDate: string;
  author: { name?: string; email?: string };
}

export const normalizeUrl = (url: string): string =>
  url.toLowerCase().split("?")[0].replace(/\/sitepages\/|\.aspx/gi, "").trim();

export const getSiteUrlFromNewsUrl = (url: string): string => {
  try {
    const urlObj = new URL(url);
    const parts = urlObj.pathname.split("/").filter((p) => p);
    const sitePath = parts.slice(0, 2).join("/");
    return `${urlObj.origin}/${sitePath}`;
  } catch (err) {
    console.error(`getSiteUrlFromNewsUrl: Invalid URL: ${url}`, err);
    return "";
  }
};

export const isExternalLink = (url: string, currentHost: string): boolean => {
  try {
    const linkHost = new URL(url).host;
    return currentHost !== linkHost;
  } catch (err) {
    console.warn(`isExternalLink: Invalid URL: ${url}`, err);
    return true;
  }
};

export const sanitizePageName = (url: string): string => {
  return url.replace(/\.aspx$/i, "").replace(/–/g, "-").replace(/[()]/g, "").replace(/\s+/g, "-").trim();
};

export class CollaborationNewsService {
  private pnpConfig: PnPConfig;
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
    this.pnpConfig = new PnPConfig(context);
  }

public async isUserInSelectedGroups(selectedGroups: IPropertyFieldGroupOrPerson[]): Promise<boolean> {
    const userEmail = this.context.pageContext.user.email?.toLowerCase() || "unknown";
    const cacheKey = `userGroups_${selectedGroups.map(g => g.id).sort().join("_")}_${userEmail}`;
    const cached = sessionStorage.getItem(cacheKey);
    if (cached !== null) {
      console.log(`isUserInSelectedGroups: Cached result for ${userEmail}: ${cached}`);
      return cached === "true";
    }

    if (userEmail === "unknown") {
      console.warn("isUserInSelectedGroups: User email is unknown, skipping group checks");
      sessionStorage.setItem(cacheKey, "false");
      return false;
    }

    if (!selectedGroups || selectedGroups.length === 0) {
      console.log(`isUserInSelectedGroups: No groups/users selected for ${userEmail}`);
      sessionStorage.setItem(cacheKey, "false");
      return false;
    }

    try {
      // Check if the current user is directly selected as an individual
      const selectedUsers = selectedGroups.filter(g => !g.id?.startsWith("c:0o.c|") && !g.id?.startsWith("c:0t.c|"));
      const isUserSelected = selectedUsers.some(user => user.email?.toLowerCase() === userEmail);
      if (isUserSelected) {
        console.log(`isUserInSelectedGroups: User ${userEmail} is directly selected to add news`);
        sessionStorage.setItem(cacheKey, "true");
        return true;
      }

      // Check AAD groups (security groups, mail-enabled security groups, distribution lists)
      const graphClient = await this.context.msGraphClientFactory.getClient("3");
      const me = await graphClient.api("/me").get();

      const aadGroups = selectedGroups.filter(g => g.id?.startsWith("c:0o.c|")).map(g => ({
        id: g.id?.split("|").pop(),
        displayName: g.fullName
      })).filter(g => g.id);
      if (aadGroups.length > 0) {
        const batchRequest = aadGroups.map(group => ({
          id: group.id,
          method: "GET",
          url: `/groups/${group.id}/members/$/microsoft.graph.user?$filter=id eq '${me.id}'`
        }));
        const batchResponse = await graphClient.api("/$batch").post({ requests: batchRequest });
        for (const response of batchResponse.responses) {
          if (response.status === 200 && response.body.value?.length > 0) {
            console.log(`isUserInSelectedGroups: User ${userEmail} in AAD group ${response.id} (${response.body.displayName || "unknown"})`);
            sessionStorage.setItem(cacheKey, "true");
            return true;
          }
        }
      }

      // Check SharePoint groups
      const sp = spfi(this.context.pageContext.web.absoluteUrl).using(SPFx(this.context));
      const spGroups = selectedGroups.filter(g => g.id?.startsWith("c:0t.c|"));
      for (const group of spGroups) {
        try {
          if (group.id !== undefined) {
            const groupId = parseInt(group.id.split("|").pop() || "0", 10);
            const users = await sp.web.siteGroups.getById(groupId).users();
            const isMember = users.some((user) => user.Email?.toLowerCase() === userEmail || user.LoginName?.toLowerCase().includes(userEmail));
            if (isMember) {
              console.log(`isUserInSelectedGroups: User ${userEmail} in SharePoint group ${groupId} (${group.fullName || "unknown"})`);
              sessionStorage.setItem(cacheKey, "true");
              return true;
            }
          }
        } catch (spGroupError) {
          console.warn(`isUserInSelectedGroups: Error checking SharePoint group ${group.id} (${group.fullName || "unknown"}) for ${userEmail}:`, spGroupError);
        }
      }

      console.log(`isUserInSelectedGroups: User ${userEmail} not in any group or directly selected`);
      sessionStorage.setItem(cacheKey, "false");
      return false;
    } catch (error) {
      console.error(`isUserInSelectedGroups: Fatal error for ${userEmail}:`, error);
      sessionStorage.setItem(cacheKey, "false");
      return false;
    }
  }
  

  public async getNewsFromSites(
    selectedSites: string[],
    selectedCategories?: string[],
    selectedGroups?: IPropertyFieldGroupOrPerson[]
  ): Promise<CollaborationNewsModel[]> {
    let newsItems: CollaborationNewsModel[] = [];
    let allCategories: Set<string> = new Set();

    for (const siteUrl of selectedSites) {
      try {
        console.log(`getNewsFromSites: Fetching from ${siteUrl}`, { selectedCategories });
        const sp = spfi(siteUrl).using(SPFx(this.context));
        console.log(`getNewsFromSites: SP instance for ${siteUrl}`, sp);
        
        const searchResults = await sp.search({
          Querytext: `Site:${siteUrl} PromotedState:2 IsContainer:false`,
          SelectProperties: [
            "Title",
            "Path",
            "ViewsLifeTime",
            "PictureThumbnailURL",
            "CreatedBy",
            "LastModifiedTime",
            "SPWebUrl",
            "Description",
          ],
          RowLimit: 50,
        });

        const rows = searchResults.PrimarySearchResults || [];
        console.log(`getNewsFromSites: Found ${rows.length} items for ${siteUrl}`);
         console.log(`getNewsFromSites: All fetched data for ${siteUrl}:`, rows);
        const siteNews = await Promise.all(
          rows.map(async (row: any) => {
            const pageUrl: string = row.Path || "";
            let validPageUrl = pageUrl;
            // const description: string = row.Description 

            // Ensure the URL ends with .aspx (and is not an external link)
            if (validPageUrl && !validPageUrl.endsWith(".aspx")) {
              // Try to extract the page name and append .aspx if missing
              const urlObj = new URL(validPageUrl, window.location.origin);
              const segments = urlObj.pathname.split("/");
              const lastSegment = segments[segments.length - 1];
              if (lastSegment && !lastSegment.endsWith(".aspx")) {
                segments[segments.length - 1] = lastSegment + ".aspx";
                urlObj.pathname = segments.join("/");
                validPageUrl = urlObj.toString();
              }
            }

            const siteWebUrl: string = row.SPWebUrl || "";
            const title: string = row.Title || "Untitled";
            const Description: string = row.Description || "No description available";
            const views: number = parseInt(row.ViewsLifeTime || "0") || 0;
            const publishedDate: Date = row.LastModifiedTime ? new Date(row.LastModifiedTime) : new Date();
            const imageUrlFromSearch: string | undefined = row.PictureThumbnailURL;
            const loginName: string | undefined = row.CreatedBy?.split("|").pop();
            // const createdBy: string = row.CreatedBy || "";

            let finalImageUrl: string | undefined = imageUrlFromSearch;
            let categories: string[] = [];

            if (siteWebUrl && validPageUrl) {
              const spItem = this.pnpConfig.getSP(siteWebUrl);
              const fileRef = decodeURIComponent(new URL(validPageUrl).pathname);

              try {
                const pageItem = await spItem.web.lists
                  .getByTitle("Site Pages")
                  .items.filter(`FileRef eq '${fileRef}'`)
                  .select("CanvasContent1", "BannerImageUrl", "Pagecategory")()
                  .then((res) => res[0]);

                if (!pageItem) {
                  console.warn(`getNewsFromSites: No page item for ${fileRef}`);
                  return null;
                }

                finalImageUrl =
                  this.extractFirstImage(pageItem?.CanvasContent1) ||
                  (typeof pageItem?.BannerImageUrl === "string" ? pageItem.BannerImageUrl : pageItem?.BannerImageUrl?.Url) ||
                  imageUrlFromSearch ||
                  "/sites/yoursite/Shared%20Documents/default-news-thumbnail.png";

                const rawCategories = pageItem?.Pagecategory;
                categories = Array.isArray(rawCategories)
                  ? rawCategories.filter((cat): cat is string => typeof cat === "string")
                  : typeof rawCategories === "string"
                  ? [rawCategories]
                  : [];
                console.log(`getNewsFromSites: "${title}" categories:`, categories);
                categories.forEach((cat) => allCategories.add(cat));
              } catch (err) {
                console.warn(`getNewsFromSites: Error for ${title}:`, err);
              }
            }

            let author: string = "Unknown";
            try {
              if (loginName) {
                const spUser = this.pnpConfig.getSP(siteUrl);
                const user = await spUser.web.ensureUser(loginName);
                author = user.data.Title||loginName


              }
              console.log("loginName:", loginName, "Author:", author);
              
            } catch (err) {
              console.warn("getNewsFromSites: Author error:", err);
            }

            return {
              id: Math.floor(Math.random() * 100000),
              title,
              url: validPageUrl,
              imageUrl: finalImageUrl,
              author,
              publishedDate,
              viewsLifeTime: views,
              categories,
              views,
              newsMeta: {},
              likes: 0,
              Description

            } as CollaborationNewsModel;
          })
        );

        let validSiteNews = siteNews.filter((item): item is CollaborationNewsModel => item !== null);
        console.log(`getNewsFromSites: Valid items for ${siteUrl}:`, validSiteNews.length);

        if (selectedCategories && selectedCategories.length > 0) {
          const filteredNews = validSiteNews.filter((item) => {
            if (!item.categories || item.categories.length === 0) {
              console.log(`getNewsFromSites: No categories for "${item.title}"`);
              return selectedCategories.includes(""); // Allow empty category if selected
            }
            const matches = item.categories.some((cat) =>
              selectedCategories.some((selCat) => selCat.toLowerCase().trim() === cat.toLowerCase().trim())
            );
            console.log(`getNewsFromSites: "${item.title}" categories: ${item.categories}, Match: ${matches}`);
            return matches;
          });
          console.log(`getNewsFromSites: Filtered by ${selectedCategories.join(", ")}:`, filteredNews.length);
          newsItems = [...newsItems, ...filteredNews];
        } else {
          newsItems = [...newsItems, ...validSiteNews];
        }
      } catch (error) {
        console.error(`getNewsFromSites: Error for ${siteUrl}:`, error);
      }
    }

    console.log("getNewsFromSites: All categories:", Array.from(allCategories));
    console.log("getNewsFromSites: Final items:", newsItems.length);
    return newsItems;
  }

  private extractFirstImage(htmlContent: string): string | null {
    if (!htmlContent) return null;
    const tempDiv = document.createElement("div");
    tempDiv.innerHTML = htmlContent;
    const imgElement = tempDiv.querySelector("img");
    return imgElement?.src || null;
  }

  public async fetchNewsViews(siteUrl: string): Promise<NewsView[]> {
    try {
      const sp = this.pnpConfig.getSP();
      const results = await sp.search({
        Querytext: "PromotedState:2",
        SelectProperties: ["Path", "Title", "ViewsLifeTime"],
        RowLimit: 50,
      });
      return results.PrimarySearchResults.map((row: any) => ({
        title: row.Title,
        url: normalizeUrl(row.Path),
        views: Number(row.ViewsLifeTime) || 0,
      }));
    } catch (error) {
      console.error("fetchNewsViews: Error:", error);
      return [];
    }
  }

  public async fetchLikesForNews(newsItems: { url: string }[], currentUserEmail: string): Promise<LikeInfo> {
    const likes: { [url: string]: number } = {};
    const liked: { [url: string]: boolean } = {};
    const likeEnabled: { [url: string]: boolean } = {};

    await Promise.all(
      newsItems.map(async (news) => {
        if (!news.url) {
          likes[news.url] = 0;
          liked[news.url] = false;
          likeEnabled[news.url] = false;
          return;
        }

        if (isExternalLink(news.url, window.location.host)) {
          likes[news.url] = 0;
          liked[news.url] = false;
          likeEnabled[news.url] = false;
          return;
        }

        try {
          const newsUrl = new URL(news.url);
          const siteAbsoluteUrl = `${newsUrl.origin}${newsUrl.pathname.split("/").slice(0, 3).join("/")}`;
          const sp = this.pnpConfig.getSP(siteAbsoluteUrl);
          const serverRelativePath = decodeURIComponent(newsUrl.pathname);

          if (!serverRelativePath.endsWith(".aspx")) {
            likes[news.url] = 0;
            liked[news.url] = false;
            likeEnabled[news.url] = false;
            return;
          }

          const page = await sp.web.loadClientsidePage(serverRelativePath);
          try {
            const likeInfo = await page.getLikedByInformation();
            likes[news.url] = parseInt(String(likeInfo?.likeCount ?? "0"), 10) || 0;
            liked[news.url] = likeInfo?.likedBy?.some((user) => user.email?.toLowerCase() === currentUserEmail) ?? false;
            likeEnabled[news.url] = true;
          } catch (error) {
            likes[news.url] = 0;
            liked[news.url] = false;
            likeEnabled[news.url] = false;
          }
        } catch (err) {
          likes[news.url] = 0;
          liked[news.url] = false;
          likeEnabled[news.url] = false;
        }
      })
    );

    return { likes, liked, likeEnabled };
  }

  public async toggleLike(url: string, isLiked: boolean): Promise<void> {
    if (!url) throw new Error("toggleLike: Invalid URL");
    try {
      const newsUrl = new URL(url);
      const siteAbsoluteUrl = `${newsUrl.origin}${newsUrl.pathname.split("/").slice(0, 3).join("/")}`;
      const serverRelativePath = decodeURIComponent(newsUrl.pathname);
      if (!serverRelativePath.endsWith(".aspx")) throw new Error("toggleLike: Invalid page path");

      const sp = this.pnpConfig.getSP(siteAbsoluteUrl);
      const page = await sp.web.loadClientsidePage(serverRelativePath);
      if (isLiked) await page.unlike();
      else await page.like();
    } catch (error) {
      console.error("toggleLike: Error:", error);
      throw error;
    }
  }

  public async fetchComments(newsUrl: string): Promise<NewsCommentItem[]> {
    if (!newsUrl) return [];
    try {
      const url = new URL(newsUrl);
      const siteAbsoluteUrl = `${url.origin}${url.pathname.split("/").slice(0, 3).join("/")}`;
      const pageName = url.pathname.split("/").pop()?.replace(".aspx", "") || "";
      if (!pageName) return [];

      const sp = this.pnpConfig.getSP(siteAbsoluteUrl);
      const pageItems = await sp.web.lists.getByTitle("Site Pages").items.filter(`FileLeafRef eq '${pageName}.aspx'`).select("Id")();
      if (!pageItems || pageItems.length === 0) return [];

      const pageId = pageItems[0].Id;
      const comments = await sp.web.lists
        .getByTitle("Site Pages")
        .items.getById(pageId)
        .comments.select("text", "createdDate", "author/name", "author/email")
        .expand("author")();

      return comments.map((comment: PageComment) => ({
        Title: comment.text || "No comment text",
        NewsURL: newsUrl,
        CreatedOn: comment.createdDate || "Unknown date",
        CreatedBy: { Title: comment.author?.name || comment.author?.email?.split("@")[0] || "Unknown User" },
      }));
    } catch (error) {
      console.error("fetchComments: Error:", error);
      return [];
    }
  }

  public async postComment(text: string, newsUrl: string): Promise<void> {
    if (!text || !newsUrl) throw new Error("postComment: Missing text or URL");
    try {
      const url = new URL(newsUrl);
      const siteAbsoluteUrl = `${url.origin}${url.pathname.split("/").slice(0, 3).join("/")}`;
      const pageName = url.pathname.split("/").pop()?.replace(".aspx", "") || "";
      if (!pageName) throw new Error("postComment: Invalid page name");

      const sp = this.pnpConfig.getSP(siteAbsoluteUrl);
      const pageItems = await sp.web.lists.getByTitle("Site Pages").items.filter(`FileLeafRef eq '${pageName}.aspx'`).select("Id")();
      if (!pageItems || pageItems.length === 0) throw new Error("postComment: Page not found");

      const pageId = pageItems[0].Id;
      await sp.web.lists.getByTitle("Site Pages").items.getById(pageId).comments.add(text);
    } catch (error) {
      console.error("postComment: Error:", error);
      throw error;
    }
  }
}