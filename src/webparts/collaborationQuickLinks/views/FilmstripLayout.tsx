import * as React from 'react';
import { Icon, TooltipHost } from '@fluentui/react';
import Carousel from "react-multi-carousel";
import "react-multi-carousel/lib/styles.css";
import { CollaborationQuickLinksModel } from '../models/CollaborationQuickLinksModel';
import styles from './FilmstripLayout.module.scss';

interface IFilmStripLayoutProps {
  linkItems: CollaborationQuickLinksModel[];
  onManageLink: (id: string) => void;
  getBorderRadiusStyle: (link: CollaborationQuickLinksModel) => React.CSSProperties;
  getCardStyle: (link: CollaborationQuickLinksModel) => React.CSSProperties;
  getOpacityStyle: (link: CollaborationQuickLinksModel) => React.CSSProperties;
  getImageStyle: () => React.CSSProperties;
  getIconStyle: () => React.CSSProperties;
  onAddLink?: () => void;
  textColor: string;
  iconColor: string;
  cardColor: string;
 filmstripSize?: 'small' | 'medium' | 'large';
 cardAlignment?: 'left' | 'center' | 'right'; 
 cardHoverColor?: string;
}

const FilmStripLayout: React.FC<IFilmStripLayoutProps> = ({
  linkItems,
  onManageLink,
  getBorderRadiusStyle,
  getCardStyle,
  getOpacityStyle,
  getImageStyle,
  getIconStyle,
  textColor,
  iconColor,
  cardColor,
  filmstripSize,
   cardHoverColor,

}) => {
  const params = new URLSearchParams(typeof window !== 'undefined' ? window.location.search : '');
  const modeValue = params.get('Mode') || params.get('mode');
  const isEditMode = modeValue?.toLowerCase()?.trim() === 'edit';
  const [editIconVisible, setEditIconVisible] = React.useState<{ [key: string]: boolean }>({});

   const containerStyle: React.CSSProperties = {
      '--card-hover-color': cardHoverColor,
    } as React.CSSProperties;


  // Responsive configuration based on filmstripSize prop
  const responsive = React.useMemo(() => ({
    superLargeDesktop: { 
      breakpoint: { max: 3000, min: 2000 }, 
      items: filmstripSize === 'small' ? 1 : filmstripSize === 'medium' ? 2 : 4
    },
    desktop: { 
      breakpoint: { max: 2000, min: 1024 }, 
      items: filmstripSize === 'small' ? 1 : filmstripSize === 'medium' ? 2 : 4
    },
    tablet: { 
      breakpoint: { max: 1024, min: 464 }, 
      items: filmstripSize === 'small' ? 1 : filmstripSize === 'medium' ? 2 : 4
    },
    mobile: { 
      breakpoint: { max: 464, min: 0 }, 
      items: 1 
    },
  }), [filmstripSize]);

  const carouselSettings = {
    showDots: true,
    arrows: true,
    infinite: false,
    autoPlay: false,
    keyBoardControl: true,
    customTransition: "all .5",
    transitionDuration: 300,
    containerClass: "carousel-container",
    dotListClass: "custom-dot-list-style",
    itemClass: "carousel-item-padding-10-px",
    slidesToSlide: filmstripSize === 'small' ? 1 : filmstripSize === 'medium' ? 2 : 3,
    renderDotsOutside: true
  };

  const toggleEditIcon = (linkId: string) => {
    setEditIconVisible((prev) => ({
      ...prev,
      [linkId]: !prev[linkId],
    }));
  };

  const renderCardContent = (link: CollaborationQuickLinksModel) => {
    return (
      <div className={styles.cardContent}>
        {link.Icon ? (
          link.Icon.startsWith('http') ? (
            <img
              src={link.Icon}
              alt={link.Title}
              className={styles.cardImage}
              style={{ 
                ...getOpacityStyle(link), 
                ...getImageStyle() 
              }}
            />
          ) : (
            <Icon
              iconName={link.Icon?.trim() || "Globe"}
              className={styles.cardIcon}
              style={{ ...getIconStyle(), color: iconColor || "white" }}
            />
          )
        ) : null}
        <div
          className={styles.cardTitle}
          title={link.Title}
          style={{ 
            color: textColor || 'white',
            maxWidth: filmstripSize === 'small' ? '150px' : '120px'
          }}
        >
          {link.Title || "Title of link"}
        </div>
      </div>
    );
  };

  return (
    <div className={`${styles.filmstripContainer} ${filmstripSize ? styles[filmstripSize] : ''}`}
      style={{backgroundColor: cardColor || 'white',...containerStyle}}>
      <Carousel responsive={responsive} {...carouselSettings}>
        {linkItems.map((link) => (
          <div
            key={link.Id}
            className={styles.cardWrapper}
            style={{ 
              ...getBorderRadiusStyle(link), 
              ...getCardStyle(link),
              position: 'relative',
              height: filmstripSize === 'small' ? '160px' : '140px'
            }}
          >
            <TooltipHost
              content={
                <div style={{ lineHeight: '1.5', color: textColor || 'black' }}>
                  <div>
                    <strong>Title:</strong> {link.Title}
                  </div>
                  {link.Description && (
                    <div>
                      <strong>Description:</strong> {link.Description}
                    </div>
                  )}
                </div>
              }
              id={`tooltip-${link.Id}`}
              calloutProps={{ gapSpace: 0 }}
            >
              <a
                 href={!isEditMode && link.URL ? link.URL : undefined}
                 role={isEditMode || !link.URL ? 'button' : undefined}
                 tabIndex={0}
                target={link.Target || '_self'}
                rel={link.Target === '_blank' ? 'noopener noreferrer' : undefined}
                style={{ textDecoration: 'none' }}
                onClick={(e) => {
                  if (isEditMode) {
                    e.preventDefault();
                    e.stopPropagation();
                    toggleEditIcon(link.Id);
                  } else if (!link.URL) {
                    e.preventDefault();
                    onManageLink(link.Id);
                  }
                }}
              >
                {renderCardContent(link)}
              </a>
            </TooltipHost>
            {isEditMode && editIconVisible[link.Id] && (
              <Icon
                iconName="Edit"
                className={styles.editButton}
                onClick={(e) => {
                  e.preventDefault();
                  e.stopPropagation();
                  onManageLink(link.Id);
                  setEditIconVisible((prev) => ({
                    ...prev,
                    [link.Id]: false,
                  }));
                }}
                title="Edit Link"
                aria-label={`Edit link ${link.Title}`}
                style={{
                  color: iconColor || '#ffffff',
                  cursor: 'pointer',
                  position: 'absolute',
                  top: '5px',
                  right: '5px',
                }}
              />
            )}
          </div>
        ))}
      </Carousel>
    </div>
  );
};

export default FilmStripLayout;