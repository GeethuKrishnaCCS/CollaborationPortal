import * as React from 'react';
import styles from './Filmstrip.module.scss';
import dialogStyles from './FilmstripDialog.module.scss';
import { IIconProps, Icon, Dialog, DialogFooter, DialogType, TextField, PrimaryButton, DefaultButton } from '@fluentui/react';
import * as moment from 'moment';
import { ProductAnniversaryService } from '../services/ProductAnniversaryService';

interface IFilmstripViewProps {
  greetings: Array<{
    Id: number;
    type: string;
    Birthday?: string;
    Location?: string;
    JoiningDate?: string;
    WeddingDate?: string;
    Employee: {
      Title: string;
      JobTitle: string;
    };
    employeeImage?: string;
    BirthdayLikedBy?: string;
    BirthdayLikesCount?: number;
    WeddingAnniversaryLikedBy?: string;
    WeddingAnniversaryLikesCount?: number;
    WorkAnniversaryLikedBy?: string;
    WorkAnniversaryLikesCount?: number;
    BirthdayComments?: string;
    WeddingAnniversaryComments?: string;
    WorkAnniversaryComments?: string;
  }>;
  backicon: IIconProps;
  nexticon: IIconProps;
  headingfontcolor: string;
  contentfontcolor: string;
  headingFontFamily: string;
  contentFontfamily: string;
  bdayGreetingWish: string;
  WorkGreetingWish: string;
  weddingGreetingWish: string;
  context: any;
  ShapeForImages: string;
  displayListFields: Array<{
    StaticName: string;
    ListFields: string;
  }>;
  AnniversaryListUrl: string;
}

const FilmstripView: React.FC<IFilmstripViewProps> = ({
  greetings,
  headingfontcolor,
  contentfontcolor,
  ShapeForImages,
  displayListFields,
  bdayGreetingWish,
  WorkGreetingWish,
  weddingGreetingWish,
  context,
  AnniversaryListUrl,
}) => {
  const itemsPerPage = 4;
  const [renderedGreetings, setRenderedGreetings] = React.useState<IFilmstripViewProps['greetings']>([]);
  const [likedStatus, setLikedStatus] = React.useState<{ [key: number]: boolean }>({});
  const [showCommentDialog, setShowCommentDialog] = React.useState(false);
  const [selectedGreeting, setSelectedGreeting] = React.useState<IFilmstripViewProps['greetings'][0] | null>(null);
  const [newComment, setNewComment] = React.useState('');
  const [existingComments, setExistingComments] = React.useState<string[]>([]);
  const service = new ProductAnniversaryService(context, context.pageContext.web.serverRelativeUrl);

  React.useEffect(() => {
    const initializeLikedStatus = async () => {
      try {
        const currentUser = await service.getCurrentUser();
        const userEmail = currentUser.Email;

        const initialLikedStatus: { [key: number]: boolean } = {};
        greetings.forEach(greeting => {
          let likedBy: string[] = [];
          switch (greeting.type) {
            case 'Birthday':
              likedBy = greeting.BirthdayLikedBy ? greeting.BirthdayLikedBy.split(';').filter(email => email) : [];
              break;
            case 'Wedding Anniversary':
              likedBy = greeting.WeddingAnniversaryLikedBy ? greeting.WeddingAnniversaryLikedBy.split(';').filter(email => email) : [];
              break;
            case 'Work Anniversary':
              likedBy = greeting.WorkAnniversaryLikedBy ? greeting.WorkAnniversaryLikedBy.split(';').filter(email => email) : [];
              break;
          }
          initialLikedStatus[greeting.Id] = likedBy.includes(userEmail);
        });

        setLikedStatus(initialLikedStatus);

        if (greetings.length > 0) {
          const today = moment().format("MM-DD");
          const filteredGreetings = greetings.filter((greeting) => {
            const birthdayMatch = greeting.Birthday && moment(greeting.Birthday).format("MM-DD") === today;
            const weddingMatch = greeting.WeddingDate && moment(greeting.WeddingDate).format("MM-DD") === today;
            const workAnniversaryMatch = greeting.JoiningDate && moment(greeting.JoiningDate).format("MM-DD") === today;
            return birthdayMatch || weddingMatch || workAnniversaryMatch;
          });
          setRenderedGreetings(filteredGreetings.slice(0, itemsPerPage));
        }
      } catch (error) {
        console.error('Error initializing liked status:', error);
      }
    };

    initializeLikedStatus();
  }, [greetings, context]);

  const getFieldValue = (greeting: any, staticName: string): string => {
    const fieldMap = displayListFields.find((f: any) => f.StaticName === staticName);
    if (fieldMap) {
      const fieldKey = fieldMap.ListFields;
      if (fieldKey.includes('.')) {
        const [fieldName, subField] = fieldKey.split('.');
        if (greeting[fieldName] && greeting[fieldName][subField]) {
          return greeting[fieldName][subField];
        }
      } else if (greeting[fieldKey]) {
        return greeting[fieldKey];
      }
    }
    return "";
  };

  const handleLikeToggle = async (greeting: IFilmstripViewProps['greetings'][0]) => {
    try {
      const currentUser = await service.getCurrentUser();
      const userEmail = currentUser.Email;
      const isCurrentlyLiked = likedStatus[greeting.Id];

      await service.toggleLike(AnniversaryListUrl, greeting.Id, userEmail, greeting.type as 'Birthday' | 'Wedding Anniversary' | 'Work Anniversary', isCurrentlyLiked);

      setLikedStatus(prev => ({
        ...prev,
        [greeting.Id]: !isCurrentlyLiked
      }));

      setRenderedGreetings(prev => prev.map(g => {
        if (g.Id === greeting.Id) {
          const updatedGreeting = { ...g };
          switch (greeting.type) {
            case 'Birthday':
              updatedGreeting.BirthdayLikedBy = isCurrentlyLiked
                ? g.BirthdayLikedBy?.split(';').filter(email => email !== userEmail).join(';') || ''
                : `${g.BirthdayLikedBy ? g.BirthdayLikedBy + ';' : ''}${userEmail}`;
              updatedGreeting.BirthdayLikesCount = (g.BirthdayLikesCount || 0) + (isCurrentlyLiked ? -1 : 1);
              break;
            case 'Wedding Anniversary':
              updatedGreeting.WeddingAnniversaryLikedBy = isCurrentlyLiked
                ? g.WeddingAnniversaryLikedBy?.split(';').filter(email => email !== userEmail).join(';') || ''
                : `${g.WeddingAnniversaryLikedBy ? g.WeddingAnniversaryLikedBy + ';' : ''}${userEmail}`;
              updatedGreeting.WeddingAnniversaryLikesCount = (g.WeddingAnniversaryLikesCount || 0) + (isCurrentlyLiked ? -1 : 1);
              break;
            case 'Work Anniversary':
              updatedGreeting.WorkAnniversaryLikedBy = isCurrentlyLiked
                ? g.WorkAnniversaryLikedBy?.split(';').filter(email => email !== userEmail).join(';') || ''
                : `${g.WorkAnniversaryLikedBy ? g.WorkAnniversaryLikedBy + ';' : ''}${userEmail}`;
              updatedGreeting.WorkAnniversaryLikesCount = (g.WorkAnniversaryLikesCount || 0) + (isCurrentlyLiked ? -1 : 1);
              break;
          }
          return updatedGreeting;
        }
        return g;
      }));
    } catch (error) {
      console.error('Error toggling like:', error);
    }
  };

  const handleCommentClick = async (greeting: IFilmstripViewProps['greetings'][0]) => {
    setSelectedGreeting(greeting);
    try {
      const comments = await service.getComments(AnniversaryListUrl, greeting.Id, greeting.type as 'Birthday' | 'Wedding Anniversary' | 'Work Anniversary');
      console.log('comments: ', comments);
      setExistingComments(comments);
    } catch (error) {
      console.error('Error fetching comments:', error);
      setExistingComments([]);
    }
    setNewComment('');
    setShowCommentDialog(true);
  };

  const handleSaveComment = async () => {
    if (!selectedGreeting || !newComment.trim()) return;

    try {
      const currentUser = await service.getCurrentUser();
      const userName = currentUser.Title;
      const commentWithMetadata = `${moment().format('DD-MM-YYYY')} - ${userName}: ${newComment}`;
      // const userEmail = currentUser.Email;
      // const commentWithMetadata = `${moment().format('YYYY-MM-DD HH:mm:ss')} - ${userEmail}: ${newComment}`;

      await service.addComment(
        AnniversaryListUrl,
        selectedGreeting.Id,
        selectedGreeting.type as 'Birthday' | 'Wedding Anniversary' | 'Work Anniversary',
        commentWithMetadata
      );

      setExistingComments(prev => [...prev, commentWithMetadata]);
      setRenderedGreetings(prev => prev.map(g => {
        if (g.Id === selectedGreeting.Id) {
          const updatedGreeting = { ...g };
          switch (selectedGreeting.type) {
            case 'Birthday':
              updatedGreeting.BirthdayComments = `${g.BirthdayComments ? g.BirthdayComments + '\n' : ''}${commentWithMetadata}`;
              break;
            case 'Wedding Anniversary':
              updatedGreeting.WeddingAnniversaryComments = `${g.WeddingAnniversaryComments ? g.WeddingAnniversaryComments + '\n' : ''}${commentWithMetadata}`;
              break;
            case 'Work Anniversary':
              updatedGreeting.WorkAnniversaryComments = `${g.WorkAnniversaryComments ? g.WorkAnniversaryComments + '\n' : ''}${commentWithMetadata}`;
              break;
          }
          return updatedGreeting;
        }
        return g;
      }));

      setNewComment('');
      setShowCommentDialog(false);
    } catch (error) {
      console.error('Error saving comment:', error);
    }
  };

  const handleCancelComment = () => {
    setShowCommentDialog(false);
    setNewComment('');
    setSelectedGreeting(null);
    setExistingComments([]);
  };

  const dialogContentProps = {
    type: DialogType.normal,
    title: `Comments for ${selectedGreeting?.type || ''} - ${selectedGreeting?.Employee.Title || ''}`,
    // subText: 'View existing comments and add a new one below.',
  };

  // Utility function to strip HTML tags and decode HTML entities
  function stripHtmlAndDecodeEntities(html: string): string {
    // Remove HTML tags
    const tmp = document.createElement('div');
    tmp.innerHTML = html;
    // Get text content (decodes entities automatically)
    return tmp.textContent || tmp.innerText || '';
  }

  return (
    <div className={styles.filmstripContainer}>
      {renderedGreetings.length > 0 ? (
        <div className={styles.cardContainer}>
          {renderedGreetings.map((greeting, index) => (
            <div className={styles.card} key={index}>
              <div className={styles.dateSection}>
                {greeting.type === "Birthday" && (
                  <>
                    <span className={styles.month}>{moment(greeting.Birthday).format("MMM")}</span>
                    <span className={styles.day}>{moment(greeting.Birthday).format("D")}</span>
                  </>
                )}
                {greeting.type === "Wedding Anniversary" && (
                  <>
                    <span className={styles.month}>{moment(greeting.WeddingDate).format("MMM")}</span>
                    <span className={styles.day}>{moment(greeting.WeddingDate).format("D")}</span>
                  </>
                )}
                {greeting.type === "Work Anniversary" && (
                  <>
                    <span className={styles.month}>{moment(greeting.JoiningDate).format("MMM")}</span>
                    <span className={styles.day}>{moment(greeting.JoiningDate).format("D")}</span>
                  </>
                )}
              </div>

              <div className={styles.imageContainer}>
                <img
                  className={`${styles.profileImage} ${ShapeForImages === "Square"
                    ? styles.squareImageCard
                    : ShapeForImages === "Circle"
                      ? styles.circleImageCard
                      : ShapeForImages === "Rectangle"
                        ? styles.rectangleImageCard
                        : ""
                    }`}
                  src={greeting.employeeImage || ""}
                  alt={`Profile picture of ${greeting.Employee.Title}`}
                />
              </div>

              <div>
                {greeting.type === "Birthday" && (<div className={styles.category}>{bdayGreetingWish}</div>)}
                {greeting.type === "Wedding Anniversary" && (<div className={styles.category}>{weddingGreetingWish}</div>)}
                {greeting.type === "Work Anniversary" && (<div className={styles.category}>{WorkGreetingWish}</div>)}

                <div className={styles.name} style={{ color: contentfontcolor }}>
                  {getFieldValue(greeting, "EmployeeName")}
                </div>
                <div className={styles.jobTitle} style={{ color: contentfontcolor }}>
                  {getFieldValue(greeting, "Designation")}
                </div>
                <div className={styles.location}>
                  {getFieldValue(greeting, "Location")}
                </div>
              </div>

              <div className={styles.footerIcons}>
                <Icon iconName="Mail" className={styles.icon} />
                <Icon
                  iconName="Comment"
                  className={styles.icon}
                  onClick={() => handleCommentClick(greeting)}
                  style={{ cursor: 'pointer' }}
                />
                <Icon
                  iconName={likedStatus[greeting.Id] ? "HeartFill" : "Heart"}
                  className={`${styles.icon} ${likedStatus[greeting.Id] ? styles.liked : ''}`}
                  onClick={() => handleLikeToggle(greeting)}
                  style={{ cursor: 'pointer' }}
                />
                <span className={styles.likesCount}>
                  {greeting.type === "Birthday" ? (greeting.BirthdayLikesCount || 0) :
                    greeting.type === "Wedding Anniversary" ? (greeting.WeddingAnniversaryLikesCount || 0) :
                      (greeting.WorkAnniversaryLikesCount || 0)}
                </span>
              </div>
            </div>
          ))}
        </div>
      ) : (
        <div className={styles.noDataMessage}>No Anniversary</div>
      )}

      <Dialog
        hidden={!showCommentDialog}
        onDismiss={handleCancelComment}
        dialogContentProps={dialogContentProps}
        modalProps={{
          containerClassName: dialogStyles.dialogContainer,
        }}
        minWidth={400}
      >
        <div className={dialogStyles.commentsContainer}>
          {existingComments.length > 0 ? (
            existingComments.map((comment, index) => (
              <div key={index} className={dialogStyles.commentItem}>
                {stripHtmlAndDecodeEntities(comment)}
              </div>
            ))
          ) : (
            <div className={dialogStyles.noComments}>No comments yet.</div>
          )}
        </div>
        <TextField
          className={dialogStyles.textArea}
          label="Add a comment"
          multiline
          rows={4}
          value={newComment}
          onChange={(_ev, newValue) => setNewComment(newValue || '')}
        />
        <DialogFooter className={dialogStyles.buttonContainer}>
          <PrimaryButton
            className={dialogStyles.postButton}
            text="Save"
            onClick={handleSaveComment}
            disabled={!newComment.trim()}
          />
          <DefaultButton
            className={dialogStyles.cancelButton}
            text="Cancel"
            onClick={handleCancelComment}
          />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default FilmstripView;