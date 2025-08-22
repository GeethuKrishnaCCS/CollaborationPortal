import * as React from "react";
import styles from "./CommentPopup.module.scss";
import { NewsCommentItem } from "../services/CollaborationNewsService";

interface CommentPopupProps {
  newsUrl: string;
  comments: NewsCommentItem[];
  commentInput: string;
  setCommentInput: (value: string) => void;
  handlePostComment: (newsUrl: string) => void;
  closePopup: () => void;
}

export const CommentPopup: React.FC<CommentPopupProps> = ({
  newsUrl,
  comments,
  commentInput,
  setCommentInput,
  handlePostComment,
  closePopup,
}) => {
  return (
    <div className={styles.popup}>
      <h4>Comments</h4>
      <ul>
        {comments?.length > 0 ? (
          comments.map((comment, idx) => (
            <li key={idx}>
              <strong>{comment.CreatedBy?.Title || "Anonymous"}</strong> (
              {new Date(comment.CreatedOn).toLocaleString()}): {comment.Title}
            </li>
          ))
        ) : (
          <li>No comments available.</li>
        )}
      </ul>
      <div>
        <textarea
          value={commentInput}
          onChange={(e) => setCommentInput(e.target.value)}
          placeholder="Write a comment..."
        />
        <div className={styles.buttonContainer}>
          <button
            onClick={() => handlePostComment(newsUrl)}
            className={styles.postButton}
          >
            Post Comment
          </button>
          <button
            className={styles.cancelButton}
            onClick={closePopup}
          >
            Cancel
          </button>
        </div>
      </div>
    </div>
  );
};