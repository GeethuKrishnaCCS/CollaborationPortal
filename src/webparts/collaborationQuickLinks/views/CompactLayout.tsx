import * as React from 'react';
import { Icon, TooltipHost } from '@fluentui/react';
import { CollaborationQuickLinksModel } from '../models/CollaborationQuickLinksModel';
import { ILayoutSettings } from '../models/CollaborationQuickLinksLayout';
import styles from './CompactLayout.module.scss';

interface ICompactLayoutProps {
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
  showImageBackground?: boolean;
}

const CompactLayout: React.FC<ICompactLayoutProps> = ({
  linkItems,
  onManageLink,
  handleDragStart,
  handleDragEnd,
  handleDragOver,
  handleDrop,
  draggedItem,
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
  cardHoverColor,
  showImageBackground 
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
  const tileWidth = settings.tileSize?.tileWidth || 180;
  const tileHeight = settings.tileSize?.tileHeight || 50;
  const numColumns = settings.tileSize?.numColumns || 5;

  // Set CSS custom properties for dynamic values
  const containerStyle: React.CSSProperties = {
    '--tile-width': `${tileWidth}px`,
    '--tile-height': `${tileHeight}px`,
    '--num-columns': numColumns,
    '--card-hover-color': cardHoverColor,
    '--card-color': cardColor || '#ffffff',
    '--text-color': textColor || '#087e86',
    '--icon-color': iconColor || 'gray',
  } as React.CSSProperties;

  React.useEffect(() => {
  }, [modeValue, isEditMode, linkItems, editIconVisible, tileWidth, tileHeight, numColumns, settings]);

  return (
    <div className={`${styles.compactLinksContainer} ${styles[cardAlignment || 'center']}`} style={containerStyle}>
      {linkItems.map((link, index) => (
        <div
          key={link.Id}
          className={`${styles.compactLinkItem} ${!link.Icon ? styles.textOnly : ''}`}
          style={getCardStyle(link)} // Apply link-specific card styles
          draggable={isEditMode}
          onDragStart={(e) => handleDragStart(e, link.Id, index)}
          onDragEnd={handleDragEnd}
          onDragOver={(e) => handleDragOver(e, link.Id, index)}
          onDrop={(e) => handleDrop(e, link.Id, index)}
        >
          <TooltipHost
            content={
              <div style={{ lineHeight: '1.5', color: textColor || '#087e86' }}>
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
            <div className={styles.linkContent}>
              <a
                href={!isEditMode && link.URL ? link.URL : undefined}
                role={isEditMode || !link.URL ? 'button' : undefined}
                tabIndex={0}
                target={link.Target || '_self'}
                rel={link.Target === '_blank' ? 'noopener noreferrer' : undefined}
                className={styles.linkproperties}
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
                  <div className={styles.linkIconContainer} style={getBorderRadiusStyle(link)}>
                    {link.Icon.startsWith('http') ? (
                      <div className={styles.compactImageBoxbackground} style={{ ...getBorderRadiusStyle(link) }}>
                      <img
                        src={link.Icon}
                        alt={link.Title}
                        className={styles.compactImageBox1}
                        style={{
                          ...getOpacityStyle(link),
                          ...getBorderRadiusStyle(link),
                          ...getImageStyle(),
                        }}
                      />
                      </div>
                    ) : (
                      <div className={styles.compacticonBox} style={{ ...getBorderRadiusStyle(link), ...getCardStyle(link) }}>
                        <div className={styles.compactIcon}
                          style={{ ...getBorderRadiusStyle(link) }}>
                          <Icon
                            iconName={link.Icon?.trim() || 'Globe'}
                            className={styles.linkIcon}
                            style={getIconStyle()}
                          />
                        </div>
                      </div>
                    )}
                  </div>
                ) : null}
                <span className={styles.linkText} title={link.Title}>
                  {link.Title}
                </span>
              </a>
            </div>
          </TooltipHost>
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
      ))}
    </div>
  );
};

export default CompactLayout;