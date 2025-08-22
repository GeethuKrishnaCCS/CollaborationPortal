import * as React from 'react';
import type { ICollaborationQuickLinksProps } from './ICollaborationQuickLinksProps';
import "react-multi-carousel/lib/styles.css";
import { CollaborationQuickLinksModel } from '../models/CollaborationQuickLinksModel';
import { LayoutType, IconSize } from '../models/CollaborationQuickLinksLayout';
import { PrimaryButton, DefaultButton, TextField, Dropdown, IDropdownOption, Icon } from '@fluentui/react';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { FilePicker } from "@pnp/spfx-controls-react/lib/FilePicker";
import { IconPicker } from "@pnp/spfx-controls-react/lib/IconPicker";
import { CollaborationAudienceTargetService } from '../services/CollaborationAudienceTargetService';
import TilesLayout from '../views/TilesLayout';
import CompactLayout from '../views/CompactLayout';
import ImageWithTitleLayout from '../views/ImageWithTitle';
import GridLayout from '../views/GridLayout';
import ButtonLayout from '../views/ButtonLayout';
import FilmstripLayout from '../views/FilmstripLayout';
import ListLayout from '../views/ListLayout';
import "isomorphic-fetch";
import * as strings from 'CollaborationQuickLinksWebPartStrings';
import { IFilePickerResult } from '@pnp/spfx-property-controls';
import styles from './CollaborationQuickLinks.module.scss';



// export default class CollaborationQuickLinks extends React.Component<ICollaborationQuickLinksProps> {
//   public render(): React.ReactElement<ICollaborationQuickLinksProps> {
//     const {

//     } = this.props;

//     return (
//       <section>
//         <div>hii</div>
//       </section>
//     );
//   }
// }

const CollaborationQuickLinks: React.FC<ICollaborationQuickLinksProps> = ({
  service, title, links, onManageLink, context, layoutType, defaultBorderRadius, imageOpacity, color,
  enableAudienceTargeting, layoutSettings, quickAccessHeader, quickAccessBackgroundColor, iconColor,
  textColor, cardColor, webPartInstanceId, filmstripSize, TitleColor, BottonColor, cardAlignment,
  cardHoverColor, TitleFontSize, quickAccessHeaderAlignment
}) => {
  const [linkItems, setLinkItems] = React.useState<CollaborationQuickLinksModel[]>(links);
  const [hideDialog, setHideDialog] = React.useState(true);
  const [newLink, setNewLink] = React.useState<CollaborationQuickLinksModel>({
    Title: '',
    URL: '',
    Icon: '',
    Id: '',
    Description: '',
    SortWeight: 0,
    TargetAudiences: [],
    Target: '_self',
    WebPartInstanceId: webPartInstanceId
  });
  const [imageOption, setImageOption] = React.useState<'default' | 'file' | 'icon' | 'text-only'>('default');
  const [draggedItem, setDraggedItem] = React.useState<{ id: string, index: number } | null>(null);
  const [dragOverItem, setDragOverItem] = React.useState<{ id: string, index: number } | null>(null);
  const [isAdmin, setIsAdmin] = React.useState<boolean>(false);

  const settings = layoutSettings[layoutType] || {
    hideText: false,
    iconSize: IconSize.Medium,
    imageSize: { width: 100, height: 100 },
  };

  const imageOptions: IDropdownOption[] = [
    { key: 'default', text: strings.UseDefaultIcon },
    { key: 'file', text: strings.SelectFromFiles },
    { key: 'icon', text: strings.SelectFromIcons },
    { key: 'text-only', text: strings.TextOnly }
  ];

  React.useEffect(() => {
    const checkAdminAndFetchLinks = async (): Promise<void> => {
      try {
        const allLinks = await service.getLinks(webPartInstanceId);
        const adminGroups = await CollaborationAudienceTargetService.fetchAdmins(context);
        const currentUserLogin = context.pageContext.user.loginName?.toLowerCase();
        const currentUserEmail = context.pageContext.user.email?.toLowerCase();
        // const allMembers = adminGroups.flatMap((group: any) => group.members);

        const allMembers = adminGroups.reduce((acc: any[], group: any) => acc.concat(group.members), []);

        const isAdminResult = allMembers.some(
          (admin: any) =>
            admin.loginName?.toLowerCase() === currentUserLogin ||
            admin.email?.toLowerCase() === currentUserEmail
        );

        setIsAdmin(isAdminResult);

        if (!enableAudienceTargeting) {
          setLinkItems(allLinks);
          return;
        }
        if (isAdminResult) {
          setLinkItems(allLinks);
        } else {
          const filteredLinks = allLinks.filter(link => {
            if (!link.TargetAudiences || link.TargetAudiences.length === 0) return true;
            return link.TargetAudiences.some((aud: any) =>
              (aud.login && aud.login.toLowerCase() === currentUserLogin) ||
              (aud.email && aud.email.toLowerCase() === currentUserEmail)
            );
          });
          setLinkItems(filteredLinks);
        }
      } catch (error) {
        console.error("Error checking admin status or fetching links:", error);
        setIsAdmin(false);
        setLinkItems([]);
      }
    };

    checkAdminAndFetchLinks();
  }, [context, service, links, enableAudienceTargeting, webPartInstanceId]);

  const getBorderRadiusStyle = (link: CollaborationQuickLinksModel): React.CSSProperties => {
    const radius = link.BorderRadius ?? defaultBorderRadius ?? 4;
    return {
      borderRadius: `${radius}px`
    };
  };

  const getIconStyle = (): React.CSSProperties => {
    const size = settings.iconSize;
    return {
      fontSize: `${size}px`,
      width: `${size}px`,
      height: `${size}px`,
      lineHeight: `${size}px`,
    };
  };

  const getImageStyle = (): React.CSSProperties => {
    const imageSize = settings.imageSize || { width: 100, height: 100 };
    return {
      width: `${imageSize.width}px`,
      height: `${imageSize.height}px`,
      objectFit: 'cover',
    };
  };

  const getOpacityStyle = (link: CollaborationQuickLinksModel): React.CSSProperties => {
    const rawOpacity = link.ImageOpacity ?? imageOpacity ?? 100;
    const parsedOpacity = typeof rawOpacity === 'string' ? parseFloat(rawOpacity) : rawOpacity;
    const finalOpacity = Math.max(0, Math.min(parsedOpacity / 100, 1));
    return {
      opacity: finalOpacity
    };
  };

  const getCardStyle = (link: CollaborationQuickLinksModel): React.CSSProperties => {
    const isDraggedOver = dragOverItem?.id === link.Id;
    const fillColor = isDraggedOver ? '#f3f2f1' : (link.color ?? color ?? '#03787c');
    return {
      backgroundColor: fillColor,
      ['--card-hover-color' as any]: cardHoverColor,
      transform: isDraggedOver ? 'scale(1.02)' : 'none',
      border: isDraggedOver ? '2px dashed #03787c' : 'none',
      transition: 'all 0.2s ease'
    };
  };

  const handleDragStart = (e: React.DragEvent<HTMLDivElement>, id: string, index: number) => {
    e.dataTransfer.setData('text/plain', id);
    setDraggedItem({ id, index });
    e.currentTarget.style.opacity = '0.5';
  };

  const handleDragEnd = (e: React.DragEvent<HTMLDivElement>) => {
    e.currentTarget.style.opacity = '1';
    setDraggedItem(null);
    setDragOverItem(null);
  };

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>, id: string, index: number) => {
    e.preventDefault();
    if (draggedItem && draggedItem.id !== id) {
      setDragOverItem({ id, index });
    }
  };

  const handleDrop = async (e: React.DragEvent<HTMLDivElement>, targetId: string, targetIndex: number) => {
    e.preventDefault();
    if (!draggedItem || draggedItem.id === targetId) {
      return;
    }
    const updatedLinks = [...linkItems];
    const sourceIndex = draggedItem.index;
    const targetIndexFinal = targetIndex;
    [updatedLinks[sourceIndex], updatedLinks[targetIndexFinal]] =
      [updatedLinks[targetIndexFinal], updatedLinks[sourceIndex]];
    const withUpdatedWeights = updatedLinks.map((item, index) => ({
      ...item,
      SortWeight: index
    }));
    try {
      setLinkItems(withUpdatedWeights);
      await service.updateLinkOrder(withUpdatedWeights);
      const refreshedLinks = await service.getLinks(webPartInstanceId);
      setLinkItems(refreshedLinks);
    } catch (error) {
      console.error('Error updating sort order:', error);
      setLinkItems(linkItems);
    } finally {
      setDraggedItem(null);
      setDragOverItem(null);
    }
  };

  const handleAddLink = async (): Promise<void> => {
    try {
      let finalIcon = newLink.Icon;
      if (imageOption !== 'text-only') {
        if (imageOption === 'default' && !finalIcon) {
          finalIcon = 'Globe';
        }
      }
      const linkToSave = {
        ...newLink,
        Icon: finalIcon,
        SortWeight: linkItems.length,
        TargetAudiences: []
      };
      const savedLink = await service.addLink(linkToSave);
      if (!savedLink?.Id) throw new Error("Saved link has no ID");
      const updatedLinks = await service.getLinks(webPartInstanceId);
      setLinkItems(updatedLinks);
      onManageLink(savedLink.Id);
      setNewLink({
        Title: '',
        URL: '',
        Icon: imageOption === 'default' ? 'Globe' : '',
        Id: '',
        Description: '',
        SortWeight: 0,
        TargetAudiences: [],
        Target: '_self',
        WebPartInstanceId: webPartInstanceId
      });
      setHideDialog(true);
    } catch (error) {
      console.error("Error adding link:", error);
    }
  };

  const containerWidth = settings.containerWidth ? `${settings.containerWidth}px` : '800px';
  const containerHeight = settings.containerHeight ? `${settings.containerHeight}px` : '400px';

  return (
    <div className={styles.quickLinks}>
      {linkItems.length === 0 && (
        <div
          className={styles.quickAccessWrapper1}
          style={{
            background: quickAccessBackgroundColor ||
              'linear-gradient(to right, #a0eacf, #6bd9f7)',
            width: containerWidth,
            height: containerHeight
          }}
        >
          <div
            className={styles.quickAccessHeader1}
            style={{
              display: 'flex',
              justifyContent: quickAccessHeaderAlignment || 'left',
              alignItems: 'left',
              flexWrap: 'wrap',
            }}
          >
            <h2
              className={styles.heading}
              style={{
                color: TitleColor || "black",
                fontSize: TitleFontSize
              }}
            >
              {quickAccessHeader}
            </h2>
            <div className={styles.actions}>
              {isAdmin && (
                <button
                  className={styles.addButton}
                  onClick={() => setHideDialog(false)}
                  style={{ backgroundColor: BottonColor || "black", color: "black" }}
                >
                  {strings.AddButtonLabel}
                </button>
              )}
            </div>
          </div>
          <p>{strings.NolinksMessage}</p>
        </div>
      )}
      {linkItems.length > 0 && layoutType === LayoutType.Tiles && (
        <div
          className={styles.quickAccessWrapper1}
          style={{
            background: quickAccessBackgroundColor ||
              'linear-gradient(to right, #a0eacf, #6bd9f7)',
            width: containerWidth,
            height: containerHeight
          }}
        >
          <div
            className={styles.quickAccessHeader1}
            style={{
              display: 'flex',
              justifyContent: quickAccessHeaderAlignment || 'left',
              alignItems: 'left',
              flexWrap: 'wrap',
            }}
          >
            <h2
              className={styles.heading}
              style={{
                color: TitleColor || "black",
                fontSize: TitleFontSize
              }}
            >
              {quickAccessHeader}
            </h2>
            <div className={styles.actions}>
              {isAdmin && (
                <button
                  className={styles.addButton}
                  onClick={() => setHideDialog(false)}
                  style={{ backgroundColor: BottonColor || "black", color: "black" }}
                >
                  {strings.AddButtonLabel}
                </button>
              )}
            </div>
          </div>
          <TilesLayout
            linkItems={linkItems}
            onManageLink={onManageLink}
            handleDragStart={handleDragStart}
            handleDragEnd={handleDragEnd}
            handleDragOver={handleDragOver}
            handleDrop={handleDrop}
            draggedItem={draggedItem}
            settings={settings}
            getBorderRadiusStyle={getBorderRadiusStyle}
            getCardStyle={getCardStyle}
            getOpacityStyle={getOpacityStyle}
            getImageStyle={getImageStyle}
            getIconStyle={getIconStyle}
            textColor={textColor}
            cardColor={cardColor}
            iconColor={iconColor}
            cardAlignment={cardAlignment}
            cardHoverColor={cardHoverColor}
          />
        </div>
      )}
      {linkItems.length > 0 && layoutType === LayoutType.ImageWithTitle && (
        <div
          className={styles.quickAccessWrapper1}
          style={{
            background: quickAccessBackgroundColor ||
              'linear-gradient(to right, #a0eacf, #6bd9f7)',
            width: containerWidth,
            height: containerHeight
          }}
        >
          <div
            className={styles.quickAccessHeader1}
            style={{
              display: 'flex',
              justifyContent: quickAccessHeaderAlignment || 'left',
              alignItems: 'left',
              flexWrap: 'wrap',
            }}
          >
            <h2
              className={styles.heading}
              style={{
                color: TitleColor || "black",
                fontSize: TitleFontSize
              }}
            >
              {quickAccessHeader}
            </h2>
            <div className={styles.actions}>
              {isAdmin && (
                <button
                  className={styles.addButton}
                  onClick={() => setHideDialog(false)}
                  style={{ backgroundColor: BottonColor || "black", color: "black" }}
                >
                  {strings.AddButtonLabel}
                </button>
              )}
            </div>
          </div>
          <ImageWithTitleLayout
            linkItems={linkItems}
            onManageLink={onManageLink}
            handleDragStart={handleDragStart}
            handleDragEnd={handleDragEnd}
            handleDragOver={handleDragOver}
            handleDrop={handleDrop}
            draggedItem={draggedItem}
            settings={settings}
            getBorderRadiusStyle={getBorderRadiusStyle}
            getCardStyle={getCardStyle}
            getOpacityStyle={getOpacityStyle}
            getImageStyle={getImageStyle}
            getIconStyle={getIconStyle}
            textColor={textColor}
            iconColor={iconColor}
            cardColor={cardColor}
            cardAlignment={cardAlignment}
            cardHoverColor={cardHoverColor}


          />
        </div>
      )}
      {linkItems.length > 0 && layoutType === LayoutType.Compact && (
        <div
          className={styles.quickAccessWrapper1}
          style={{
            background: quickAccessBackgroundColor ||
              'linear-gradient(to right, #a0eacf, #6bd9f7)',
            width: containerWidth,
            height: containerHeight
          }}
        >
          <div
            className={styles.quickAccessHeader1}
            style={{
              display: 'flex',
              justifyContent: quickAccessHeaderAlignment || 'left',
              alignItems: 'left',
              flexWrap: 'wrap',
            }}
          >
            <h2
              className={styles.heading}
              style={{
                color: TitleColor || "black",
                fontSize: TitleFontSize
              }}
            >
              {quickAccessHeader}
            </h2>
            <div className={styles.actions}>
              {isAdmin && (
                <button
                  className={styles.addButton}
                  onClick={() => setHideDialog(false)}
                  style={{ backgroundColor: BottonColor || "black", color: "black" }}
                >
                  {strings.AddButtonLabel}
                </button>
              )}
            </div>
          </div>
          <CompactLayout
            linkItems={linkItems}
            onManageLink={onManageLink}
            handleDragStart={handleDragStart}
            handleDragEnd={handleDragEnd}
            handleDragOver={handleDragOver}
            handleDrop={handleDrop}
            draggedItem={draggedItem}
            settings={settings}
            getBorderRadiusStyle={getBorderRadiusStyle}
            getCardStyle={getCardStyle}
            getOpacityStyle={getOpacityStyle}
            getImageStyle={getImageStyle}
            getIconStyle={getIconStyle}
            textColor={textColor}
            iconColor={iconColor}
            cardColor={cardColor}
            cardAlignment={cardAlignment}
            cardHoverColor={cardHoverColor}
          />
        </div>
      )}
      {linkItems.length > 0 && layoutType === LayoutType.FilmStrip && (
        <div
          className={styles.quickAccessWrapper1}
          style={{
            background: quickAccessBackgroundColor ||
              'linear-gradient(to right, #a0eacf, #6bd9f7)',
            width: containerWidth,
            height: containerHeight
          }}
        >
          <div
            className={styles.quickAccessHeader1}
            style={{
              display: 'flex',
              justifyContent: quickAccessHeaderAlignment || 'center',
              // alignItems: 'center',
              flexWrap: 'wrap',
            }}
          >
            <h2
              className={styles.heading}
              style={{
                color: TitleColor || "black",
                fontSize: TitleFontSize
              }}
            >
              {quickAccessHeader}
            </h2>
            <div className={styles.actions}>
              {isAdmin && (
                <button
                  className={styles.addButton}
                  onClick={() => setHideDialog(false)}
                  style={{ backgroundColor: BottonColor || "black", color: "black" }}
                >
                  {strings.AddButtonLabel}
                </button>
              )}
            </div>
          </div>
          <FilmstripLayout
            linkItems={linkItems}
            onManageLink={onManageLink}
            getBorderRadiusStyle={getBorderRadiusStyle}
            getCardStyle={getCardStyle}
            getOpacityStyle={getOpacityStyle}
            getImageStyle={getImageStyle}
            getIconStyle={getIconStyle}
            filmstripSize={filmstripSize}
            textColor={textColor}
            iconColor={iconColor}
            cardColor={cardColor}
            cardHoverColor={cardHoverColor}
          />
        </div>
      )}
      {linkItems.length > 0 && layoutType === LayoutType.Grid && (
        <div
          className={styles.quickAccessWrapper1}
          style={{
            background: quickAccessBackgroundColor ||
              'linear-gradient(to right, #a0eacf, #6bd9f7)',
            width: containerWidth,
            height: containerHeight
          }}
        >
          <div
            className={styles.quickAccessHeader1}
            style={{
              display: 'flex',
              justifyContent: quickAccessHeaderAlignment || 'left',
              alignItems: 'left',
              flexWrap: 'wrap',
            }}
          >
            <h2
              className={styles.heading}
              style={{
                color: TitleColor || "black",
                fontSize: TitleFontSize
              }}
            >
              {quickAccessHeader}
            </h2>
            <div className={styles.actions}>
              {isAdmin && (
                <button
                  className={styles.addButton}
                  onClick={() => setHideDialog(false)}
                  style={{ backgroundColor: BottonColor || "black", color: "black" }}
                >
                  {strings.AddButtonLabel}
                </button>
              )}
            </div>
          </div>
          <GridLayout
            linkItems={linkItems}
            onManageLink={onManageLink}
            handleDragStart={handleDragStart}
            handleDragEnd={handleDragEnd}
            handleDragOver={handleDragOver}
            handleDrop={handleDrop}
            draggedItem={draggedItem}
            settings={settings}
            getBorderRadiusStyle={getBorderRadiusStyle}
            getCardStyle={getCardStyle}
            getOpacityStyle={getOpacityStyle}
            getImageStyle={getImageStyle}
            getIconStyle={getIconStyle}
            textColor={textColor}
            iconColor={iconColor}
            cardColor={cardColor}
            cardAlignment={cardAlignment}
            cardHoverColor={cardHoverColor}
          />
        </div>
      )}
      {linkItems.length > 0 && layoutType === LayoutType.Button && (
        <div
          className={styles.quickAccessWrapper1}
          style={{
            background: quickAccessBackgroundColor ||
              'linear-gradient(to right, #a0eacf, #6bd9f7)',
            width: containerWidth,
            height: containerHeight
          }}
        >
          <div
            className={styles.quickAccessHeader1}
            style={{
              display: 'flex',
              justifyContent: quickAccessHeaderAlignment || 'left',
              alignItems: 'left',
              flexWrap: 'wrap',
            }}
          >
            <h2
              className={styles.heading}
              style={{
                color: TitleColor || "black",
                fontSize: TitleFontSize
              }}
            >
              {quickAccessHeader}
            </h2>
            <div className={styles.actions}>
              {isAdmin && (
                <button
                  className={styles.addButton}
                  onClick={() => setHideDialog(false)}
                  style={{ backgroundColor: BottonColor || "black", color: "black" }}
                >
                  {strings.AddButtonLabel}
                </button>
              )}
            </div>
          </div>
          <ButtonLayout
            linkItems={linkItems}
            onManageLink={onManageLink}
            handleDragStart={handleDragStart}
            handleDragEnd={handleDragEnd}
            handleDragOver={handleDragOver}
            handleDrop={handleDrop}
            draggedItem={draggedItem}
            settings={settings}
            getBorderRadiusStyle={getBorderRadiusStyle}
            getCardStyle={getCardStyle}
            getOpacityStyle={getOpacityStyle}
            getImageStyle={getImageStyle}
            getIconStyle={getIconStyle}
            textColor={textColor}
            iconColor={iconColor}
            cardColor={cardColor}
            cardAlignment={cardAlignment}
            cardHoverColor={cardHoverColor}
          />
        </div>
      )}
      {linkItems.length > 0 && layoutType === LayoutType.List && (
        <div
          className={styles.quickAccessWrapper1}
          style={{
            background: quickAccessBackgroundColor ||
              'linear-gradient(to right, #a0eacf, #6bd9f7)',
            width: containerWidth,
            height: containerHeight
          }}
        >
          <div
            className={styles.quickAccessHeader1}
            style={{
              display: 'flex',
              justifyContent: quickAccessHeaderAlignment || 'left',
              alignItems: 'left',
              flexWrap: 'wrap',
            }}
          >
            <h2
              className={styles.heading}
              style={{
                color: TitleColor || "black",
                fontSize: TitleFontSize
              }}
            >
              {quickAccessHeader}
            </h2>
            <div className={styles.actions}>
              {isAdmin && (
                <button
                  className={styles.addButton}
                  onClick={() => setHideDialog(false)}
                  style={{ backgroundColor: BottonColor || "black", color: "black" }}
                >
                  {strings.AddButtonLabel}
                </button>
              )}
            </div>
          </div>
          <ListLayout
            linkItems={linkItems}
            onManageLink={onManageLink}
            handleDragStart={handleDragStart}
            handleDragEnd={handleDragEnd}
            handleDragOver={handleDragOver}
            handleDrop={handleDrop}
            draggedItem={draggedItem}
            settings={settings}
            getBorderRadiusStyle={getBorderRadiusStyle}
            getCardStyle={getCardStyle}
            getOpacityStyle={getOpacityStyle}
            getImageStyle={getImageStyle}
            getIconStyle={getIconStyle}
            textColor={textColor}
            iconColor={iconColor}
            cardColor={cardColor}
            cardAlignment={cardAlignment}
            cardHoverColor={cardHoverColor}
          />
        </div>
      )}
      <Dialog
        hidden={hideDialog}
        onDismiss={() => setHideDialog(true)}
        dialogContentProps={{
          type: DialogType.normal,
          title: strings.AddNewLinkTitle,
        }}
      >
        <TextField
          label="Title"
          value={newLink.Title}
          onChange={(e, val) => setNewLink({ ...newLink, Title: val || '' })}
          required
        />
        <TextField
          label="URL"
          value={newLink.URL}
          onChange={(e, val) => setNewLink({ ...newLink, URL: val || '' })}
          required
        />
        <Dropdown
          label="Thumbnail Option"
          options={imageOptions}
          selectedKey={imageOption}
          onChange={(e, option) => {
            if (option) {
              const newOption = option.key as 'default' | 'file' | 'icon' | 'text-only';
              setImageOption(newOption);
              if (newOption === 'text-only') {
                setNewLink({ ...newLink, Icon: '' });
              } else if (newOption === 'default' && !newLink.Icon) {
                setNewLink({ ...newLink, Icon: 'Globe' });
              }
            }
          }}
        />
        {imageOption === 'default' && (
          <div>
            <Icon
              iconName={newLink.Icon || 'Globe'}
              style={{ ...getIconStyle(), color: '#fff' }}
            />
          </div>
        )}
        {imageOption === 'file' && (
          <FilePicker
            bingAPIKey="<BING API KEY>"
            accepts={[".gif", ".jpg", ".jpeg", ".bmp", ".dib", ".tif", ".tiff", ".ico", ".png", ".jxr", ".svg"]}
            buttonIcon="FileImage"
            buttonLabel="Select Image"
            context={context}
            onSave={async (filePickerResults: IFilePickerResult[]) => {
              if (filePickerResults && filePickerResults.length > 0) {
                const file = filePickerResults[0];
                let imageUrl = file.fileAbsoluteUrl;
                if (!imageUrl && file.downloadFileContent) {
                  const blob = await file.downloadFileContent();
                  imageUrl = URL.createObjectURL(blob);
                }
                if (imageUrl) {
                  setNewLink({ ...newLink, Icon: imageUrl });
                } else {
                  alert("Selected file does not have a valid URL. Please choose another image.");
                }
              }
            }}
          />
        )}
        {imageOption === 'icon' && (
          <div className={styles.iconButton}>
            <IconPicker
              buttonLabel={'Icon'}
              onChange={(iconName: string) => setNewLink({ ...newLink, Icon: iconName })}
              onSave={(iconName: string) => setNewLink({ ...newLink, Icon: iconName })}
            />
          </div>
        )}
        {imageOption === 'text-only' && (
          <div>
            {newLink.Title}
          </div>
        )}
        <DialogFooter>
          <PrimaryButton
            onClick={handleAddLink}
            text="Add Link"
            disabled={!newLink.Title || !newLink.URL}
          />
          <DefaultButton onClick={() => setHideDialog(true)} text="Cancel" />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default CollaborationQuickLinks;