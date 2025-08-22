import * as React from "react";
import { useState, useEffect, useRef } from "react";
import styles from "./NewsHeader.module.scss";
import { CollaborationNewsModel } from "../models/CollaborationNewsModel";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { CollaborationAdminService} from "../services/CollaborationAdminService"; 

interface INewsHeaderProps {
  newsItems: CollaborationNewsModel[];
  handleAddNewsPost?: (siteUrl: string) => void;
  handleAddNewsLink?: (siteUrl: string) => void;
  headingColor?: string;
  customTitle: string;
  // canAddNews: boolean;
  context: WebPartContext;
  selectedSites?: string[];
  buttonBackgroundColor?: string;
}

export const NewsHeader: React.FC<INewsHeaderProps> = ({
  newsItems,
  handleAddNewsPost,
  handleAddNewsLink,
  headingColor,
  customTitle,
  // canAddNews,
  context,
  selectedSites = [],
  buttonBackgroundColor,
}) => {
  console.log("NewsHeader: Props:", {
    selectedSites,
    newsItemsLength: newsItems.length,
    headingColor,
    userEmail: context.pageContext.user.email,
    hasHandlers: !!handleAddNewsPost && !!handleAddNewsLink,
    buttonBackgroundColor,
  });

  // State for dropdowns
  const [isAddNewsOpen, setIsAddNewsOpen] = useState<boolean>(false);
  const [isSeeAllOpen, setIsSeeAllOpen] = useState<boolean>(false);
  const [openSubMenuIndex, setOpenSubMenuIndex] = useState<number | null>(null);
  const addNewsRef = useRef<HTMLDivElement>(null);
  const seeAllRef = useRef<HTMLDivElement>(null);
   const [isAdmin, setIsAdmin] = useState(false);

  // Toggle Add News dropdown
  const toggleAddNewsDropdown = () => {
    setIsAddNewsOpen((prev) => !prev);
    setIsSeeAllOpen(false); // Close See All dropdown
    if (isAddNewsOpen) {
      setOpenSubMenuIndex(null); // Close submenus
    }
  };

  // Toggle See All dropdown
  const toggleSeeAllDropdown = () => {
    setIsSeeAllOpen((prev) => !prev);
    setIsAddNewsOpen(false); // Close Add News dropdown
    setOpenSubMenuIndex(null); // Close submenus
  };

  // Toggle submenu for Add News
  const toggleSubMenu = (index: number) => {
    setOpenSubMenuIndex(openSubMenuIndex === index ? null : index);
    setIsAddNewsOpen(true); // Ensure Add News dropdown stays open
  };

   useEffect(() => {
    const checkAdmin = async () => {
      const result = await CollaborationAdminService.isCurrentUserAdmin(context);
      console.log("Admin check result:", result);
      setIsAdmin(result);
    };
    checkAdmin();
  }, [context]);

  // Close dropdowns on outside click
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      const target = event.target as Node;
      if (
        (addNewsRef.current && !addNewsRef.current.contains(target)) &&
        (seeAllRef.current && !seeAllRef.current.contains(target))
      ) {
        console.log("Clicked outside both dropdowns, closing");
        setIsAddNewsOpen(false);
        setIsSeeAllOpen(false);
        setOpenSubMenuIndex(null);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, []);

  // Handle keyboard accessibility for both buttons
  const handleKeyDown = (
    event: React.KeyboardEvent<HTMLButtonElement>,
    isAddNews: boolean
  ) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      if (isAddNews) {
        toggleAddNewsDropdown();
      } else {
        toggleSeeAllDropdown();
      }
    }
  };

  // Function to darken a hex color
  const darkenColor = (hex: string, percent: number): string => {
    let r = parseInt(hex.slice(1, 3), 16);
    let g = parseInt(hex.slice(3, 5), 16);
    let b = parseInt(hex.slice(5, 7), 16);

    r = Math.max(0, Math.floor(r * (1 - percent / 100)));
    g = Math.max(0, Math.floor(g * (1 - percent / 100)));
    b = Math.max(0, Math.floor(b * (1 - percent / 100)));

    return `#${r.toString(16).padStart(2, "0")}${g.toString(16).padStart(2, "0")}${b.toString(16).padStart(2, "0")}`;
  };

  // Inline styles for buttons
  const buttonStyle = buttonBackgroundColor
    ? {
        backgroundColor: buttonBackgroundColor,
        border: "none",
        height: "30px",
        width: "76px",
        // Override hover and active states
        ":hover": {
          backgroundColor: darkenColor(buttonBackgroundColor, 10),
        },
        ":active": {
          backgroundColor: darkenColor(buttonBackgroundColor, 20),
        },
      }
    : {
        border: "none",
      };

  return (
    <div className={styles.newsHeader}>
      <div className={styles["news-title"]} style={{ color: headingColor || "#000000" }}>
        {customTitle}
      </div>
      <div className={styles.actions}>
        {/* Add News Dropdown */}
        {selectedSites.length > 0 && (
          <div className={styles.dropdown} ref={addNewsRef}>
            {isAdmin && (
            <button
              className={`${styles["add-button"]} ${isAddNewsOpen ? styles.open : ""}`}
              onClick={toggleAddNewsDropdown}
              onKeyDown={(e) => handleKeyDown(e, true)}
              aria-expanded={isAddNewsOpen}
              aria-haspopup="true"
              style={buttonStyle} // Already applied
            >
              + Add <span className={styles["button-arrow"]}>{isAddNewsOpen}</span>
            </button>)}
            <div className={`${styles["dropdown-content"]} ${isAddNewsOpen ? styles.show : ""}`}>
              {selectedSites.map((siteUrl, index) => (
                <div key={`add-news-${index}`} className={styles["sub-menu"]}>
                  <div
                    className={styles["site-item"]}
                    onClick={() => toggleSubMenu(index)}
                  >
                    <span className={styles["site-name"]}>
                      {siteUrl.replace(/^https?:\/\/[^/]+\/sites\//, "")}
                    </span>
                    <span className={styles.arrow}>
                      {openSubMenuIndex === index ? "▼" : "▶"}
                    </span>
                  </div>
                  {openSubMenuIndex === index && (
                    <div className={styles["sub-menu-content"]}>
                      <span
                        className={handleAddNewsPost ? styles["sub-menu-item"] : styles.disabled}
                        onClick={() => {
                          if (handleAddNewsPost) {
                            handleAddNewsPost(siteUrl);
                            setIsAddNewsOpen(false);
                          } else {
                            console.warn("handleAddNewsPost is undefined for site:", siteUrl);
                          }
                        }}
                      >
                        📝 News Post
                      </span>
                      <span
                        className={handleAddNewsLink ? styles["sub-menu-item"] : styles.disabled}
                        onClick={() => {
                          if (handleAddNewsLink) {
                            handleAddNewsLink(siteUrl);
                            setIsAddNewsOpen(false);
                          } else {
                            console.warn("handleAddNewsLink is undefined for site:", siteUrl);
                          }
                        }}
                      >
                        🔗 News Link
                      </span>
                    </div>
                  )}
                </div>
              ))}
            </div>
          </div>
        )}
        {/* See All Dropdown */}
        {selectedSites.length > 0 && (
          <div className={styles.dropdown} ref={seeAllRef}>
            <span 
      className={styles.seeAllLink}
      onClick={toggleSeeAllDropdown}
      onKeyDown={(e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          toggleSeeAllDropdown();
        }
      }}
      role="button"
      tabIndex={0}
      aria-expanded={isSeeAllOpen}
      aria-haspopup="true"
    >
      See All
      <span className={styles.linkArrow}>{isSeeAllOpen}</span>
    </span>
            <div className={`${styles["dropdown-content"]} ${isSeeAllOpen ? styles.show : ""}`}>
              {selectedSites.map((siteUrl, index) => {
                const normalizedSiteUrl = siteUrl.replace(/\/+$/, "");
                const newsUrl = `${normalizedSiteUrl}/_layouts/15/news.aspx`;
                console.log(`See All: Site URL: ${normalizedSiteUrl}, News URL: ${newsUrl}`);
                return (
                  <a
                    key={`see-all-${index}`}
                    href={newsUrl}
                    className={styles["site-item"]}
                    onClick={(e) => {
                      e.preventDefault();
                      console.log(`Navigating to: ${newsUrl}`);
                      setIsSeeAllOpen(false);
                      window.location.assign(newsUrl);
                    }}
                  >
                    <span className={styles["site-name"]}>
                      {normalizedSiteUrl.replace(/^https?:\/\/[^/]+\/sites\//, "")}
                    </span>
                  </a>
                );
              })}
            </div>
          </div>
        )}
      </div>
    </div>
  );
};