import * as React from 'react';
import styles from './CompactView.module.scss';
import * as moment from 'moment';

interface ICompactViewProps {
  greetings: any[];
  headingfontcolor: string;
  contentfontcolor: string;
  bdayGreetingWish: string;
  WorkGreetingWish: string;
  weddingGreetingWish: string;
  context: any;
  // displayName: boolean;
  // displayJobTitle: boolean;
  // displayEmail: boolean;
  // displayType: boolean;
  // displayDOB: boolean;
  // displayJoiningDate: boolean;
  // displayWeddingDate: boolean;
}

const CompactView: React.FC<ICompactViewProps> = (props) => {
  console.log('greetings: ', props.greetings);
  return (
    <>
      {props.greetings.length > 0 ? (
      <div className={styles.CompactGrid}>
        {props.greetings.map((greeting: any) => (
          <div className={styles.CompactCard} key={greeting.Id}>
            <div className={styles.cardImage}>
              <img src={greeting.employeeImage || ''} alt={`Profile picture of ${greeting.Employee.Title}`} />
            </div>
            <div className={styles.cardDetails}>
              { <div>{greeting.Employee.Title}</div>}
              {<div>{greeting.Employee.JobTitle}</div>}
              { <div className={styles.cardRole}>{greeting.type}</div>}

              {greeting.type === "Birthday" && (
                <div className={styles.cardDate}>
                  {"On " + moment(greeting.Birthday).format("DD/MM")}
                </div>
              )}
              {greeting.type === "Work Anniversary" && (
                <div className={styles.cardDate}>
                  {"On " + moment(greeting.JoiningDate).format("DD/MM")}
                </div>
              )}
              { greeting.type === "Wedding Anniversary" && (
                <div className={styles.cardDate}>
                  {"On " + moment(greeting.WeddingDate).format("DD/MM")}
                </div>
              )}
            </div>
          </div>
        ))}
      </div>
      ) : (
      <div className={styles.noData}>
        <p>No Anniversary</p>
      </div>
      )}
    </>
  );
};

export default CompactView;