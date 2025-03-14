import * as React from 'react';
import styles from './Filmstrip.module.scss';
import { IIconProps, Icon } from '@fluentui/react';
import * as moment from 'moment';

interface IFilmstripViewProps {
  greetings: Array<{
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
  }>;
  backicon: IIconProps;
  nexticon: IIconProps;
  headingfontcolor: string;
  contentfontcolor: string;
  bdayGreetingWish: string;
  WorkGreetingWish: string;
  weddingGreetingWish: string;
  context: any;
  
  ShapeForImages: string;
  selectedFields: string[];
}

const FilmstripView: React.FC<IFilmstripViewProps> = ({
  greetings,
  headingfontcolor,
  contentfontcolor,

  ShapeForImages,
  selectedFields
}) => {
  const itemsPerPage = 4;
  const [renderedGreetings, setRenderedGreetings] = React.useState<IFilmstripViewProps['greetings']>([]);

  React.useEffect(() => {
    if (greetings.length > 0) {
      setRenderedGreetings(greetings.slice(0, itemsPerPage));
    }
  }, [greetings]);

  const getFormattedDate = React.useCallback((greeting: IFilmstripViewProps['greetings'][0]): string => {
    const date = greeting.type === "Birthday"
      ? greeting.Birthday
      : greeting.type === "Work Anniversary"
        ? greeting.JoiningDate
        : greeting.WeddingDate;
    return moment(date).format("MMM D"); // Example: "Feb 19"
  }, []);

  return (
    <div className={styles.filmstripContainer}>
      {renderedGreetings.length > 0 ? (
        <div className={styles.cardContainer}>
          {renderedGreetings.map((greeting, index) => (
            <div className={styles.card} key={index}>
              {/* Date Section */}
              <div className={styles.dateSection}>
                <span className={styles.month}>{getFormattedDate(greeting).split(" ")[0]}</span>
                <span className={styles.day}>{getFormattedDate(greeting).split(" ")[1]}</span>
              </div>

              {/* Employee Image */}
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

              {/* type */}
              { <div className={styles.category}>{greeting.type}</div>}

              {/* Employee Details */}
              <div className={styles.details}>

                { <div className={styles.name} style={{ color: contentfontcolor }}>
                  {greeting.Employee.Title}
                </div>}

                { <div className={styles.jobTitle} style={{ color: contentfontcolor }}>
                  {greeting.Employee.JobTitle}
                </div>}

                {<div className={styles.location}>{greeting.Location}</div>}

              </div>

              {/* Footer Icons */}
              <div className={styles.footerIcons}>
                <Icon iconName="Mail" className={styles.icon} />
                <Icon iconName="Comment" className={styles.icon} />
                <Icon iconName="Heart" className={styles.icon} />
              </div>
            </div>
          ))}
        </div>
      ) : (
        <div className={styles.noDataMessage}>No Anniversary</div>
      )}
    </div>
   
  );
};

export default FilmstripView;
