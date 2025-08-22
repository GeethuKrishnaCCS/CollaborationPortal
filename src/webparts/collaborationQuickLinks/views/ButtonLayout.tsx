import * as React from 'react';
import { Icon, TooltipHost } from '@fluentui/react';
import { CollaborationQuickLinksModel } from '../models/CollaborationQuickLinksModel';
import styles from './ButtonLayout.module.scss';
import { ILayoutSettings } from '../models/CollaborationQuickLinksLayout';

interface IButtonLayoutProps {
  linkItems: CollaborationQuickLinksModel[];
  onManageLink: (id: string) => void;
  handleDragStart: (e: React.DragEvent<HTMLDivElement>, id: string, index: number) => void;
  handleDragEnd: (e: React.DragEvent<HTMLDivElement>) => void;
  handleDragOver: (e: React.DragEvent<HTMLDivElement>, id: string, index: number) => void;
  handleDrop: (e: React.DragEvent<HTMLDivElement>, id: string, index: number) => void;
  draggedItem: { id: string; index: number } | null;
  settings: ILayoutSettings;
  getBorderRadiusStyle: (link: CollaborationQuickLinksModel) => React.CSSProperties;
  getOpacityStyle: (link: CollaborationQuickLinksModel) => React.CSSProperties;
  getCardStyle: (link: CollaborationQuickLinksModel) => React.CSSProperties;
  getImageStyle: () => React.CSSProperties;
  getIconStyle: () => React.CSSProperties;
  textColor: string;
  iconColor: string;
  cardColor: string;
  cardAlignment?: 'left' | 'center' | 'right';
  cardHoverColor?: string; 
}

const ButtonLayout: React.FC<IButtonLayoutProps> = ({
  linkItems,
  onManageLink,
  handleDragStart,
  handleDragEnd,
  handleDragOver,
  handleDrop,
  settings,
  getBorderRadiusStyle,
  getOpacityStyle,
  getCardStyle,
  getImageStyle,
  getIconStyle,
  textColor,
  iconColor,
  cardColor,
  cardAlignment = 'center',
  cardHoverColor

}) => {
  const params = new URLSearchParams(typeof window !== 'undefined' ? window.location.search : '');
  const modeValue = params.get('Mode') || params.get('mode');
  const isEditMode = modeValue?.toLowerCase()?.trim() === 'edit';

  const [editIconVisible, setEditIconVisible] = React.useState<{ [key: string]: boolean }>({});

  const toggleEditIcon = (linkId: string) => {
    setEditIconVisible((prev) => ({
      ...prev,
      [linkId]: !prev[linkId],
    }));
  };

  // Get tile size from settings, with fallbacks
  const tileWidth = settings.tileSize?.tileWidth || 120;
  const tileHeight = settings.tileSize?.tileHeight || 50;
  const numColumns = settings.tileSize?.numColumns || 5;

  // Set CSS custom properties for dynamic values
  const containerStyle: React.CSSProperties = {
    '--tile-width': `${tileWidth}px`,
    '--tile-height': `${tileHeight}px`,
    '--num-columns': numColumns,
    '--card-hover-color': cardHoverColor,
    '--card-color': cardColor,
    '--text-color': textColor,
    '--icon-color': iconColor,
  } as React.CSSProperties;

  React.useEffect(() => {
  }, [modeValue, isEditMode, linkItems, editIconVisible, tileWidth, tileHeight, numColumns, settings]);

  return (
   <div className={`${styles.buttonLayoutContainer} ${styles[cardAlignment || 'center']}`} style={containerStyle}>
      {linkItems.map((link, index) => (
        <div
          key={link.Id}
          className={`${styles.buttonLayoutItem} ${!link.Icon ? styles.textOnly : ''}`}
          style={getCardStyle(link)}
          draggable={isEditMode}
          onDragStart={(e) => handleDragStart(e, link.Id, index)}
          onDragEnd={handleDragEnd}
          onDragOver={(e) => handleDragOver(e, link.Id, index)}
          onDrop={(e) => handleDrop(e, link.Id, index)}
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
            <div className={styles.buttonLinkContent}>
              <a
                href={!isEditMode && link.URL ? link.URL : undefined}
                role={isEditMode || !link.URL ? 'button' : undefined}
                tabIndex={0}
                target={link.Target || '_self'}
                rel={link.Target === '_blank' ? 'noopener noreferrer' : undefined}
                className={styles.buttonLinkAnchor}
                onClick={(e) => {
                  if (isEditMode) {
                    e.preventDefault();
                    e.stopPropagation();
                    toggleEditIcon(link.Id);
                  } else {
                    if (!link.URL) {
                      e.preventDefault();
                      onManageLink(link.Id);
                    }
                  }
                }}
              >
                {link.Icon ? (
                  <div className={styles.buttonLinkContent}>
                    <div className={styles.linkIconContainer} style={getBorderRadiusStyle(link)}>
                      {link.Icon.startsWith('http') ? (
                        <img
                          src={link.Icon}
                          alt={link.Title}
                          className={styles.buttonImageIcon}
                          style={{
                            ...getOpacityStyle(link),
                            ...getBorderRadiusStyle(link),
                            ...getImageStyle(),
                          }}
                        />
                      ) : (
                        <div className={styles.buttonIconWrapper} style={getCardStyle(link)}>
                          <Icon
                            iconName={link.Icon?.trim() || 'Globe'}
                            className={styles.buttonIcon}
                            style={getIconStyle()}
                          />
                        </div>
                      )}
                    </div>
                    <span className={styles.buttonTextLabel}>{link.Title}</span>
                  </div>
                ) : (
                  <span className={styles.buttonTextLabel}>{link.Title}</span>
                )}
              </a>
              {isEditMode && editIconVisible[link.Id] && (
                <Icon
                  iconName="Edit"
                  className={styles.editButton}
                  onClick={(e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    console.log('Edit icon clicked:', link.Id);
                    onManageLink(link.Id);
                    setEditIconVisible((prev) => ({
                      ...prev,
                      [link.Id]: false,
                    }));
                  }}
                  title="Edit Link"
                  aria-label={`Edit link ${link.Title}`}
                />
              )}
            </div>
          </TooltipHost>
        </div>
      ))}
    </div>
  );
};

export default ButtonLayout;