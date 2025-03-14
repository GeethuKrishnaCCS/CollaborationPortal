import * as React from "react";
import styles from "./CardView.module.scss";
import { IIconProps, Icon } from "@fluentui/react";
import * as moment from "moment";

interface ICardViewProps {
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

}

const CardView: React.FC<ICardViewProps> = ({
    greetings,
    headingfontcolor,
    contentfontcolor,
   
}) => {
    const getFormattedDate = (greeting: ICardViewProps["greetings"][0]): string => {
        const date =
            greeting.type === "Birthday"
                ? greeting.Birthday
                : greeting.type === "Work Anniversary"
                    ? greeting.JoiningDate
                    : greeting.WeddingDate;

        return moment(date).format("MMM D"); // Example: "Feb 19"
    };

    return (
    
        <div className={styles.gridContainer}>
            {greetings.length > 0 ? (
                greetings.map((greeting, index) => (
                    <div key={index} className={styles.card}>
                        {/* Date Section */}
                        <div className={styles.dateSection}>
                            <span className={styles.month}>{getFormattedDate(greeting).split(" ")[0]}</span>
                            <span className={styles.day}>{getFormattedDate(greeting).split(" ")[1]}</span>
                            <span className={styles.eventType}>{greeting.type}</span>
                        </div>

                        {/* Employee Image */}
                        <div className={styles.imageContainer}>
                            <img
                                src={greeting.employeeImage || ""}
                                alt={`Profile picture of ${greeting.Employee.Title}`}
                                className={styles.profileImage}
                            />
                        </div>

                        {/* Employee Details */}
                        <div className={styles.details}>
                            <div className={styles.name} style={{ color: headingfontcolor }}>
                                {greeting.Employee.Title}
                            </div>
                            <div className={styles.jobTitle} style={{ color: contentfontcolor }}>
                                {greeting.Employee.JobTitle}
                            </div>
                            <div className={styles.location}>{greeting.Location}</div>
                        </div>

                        {/* Footer Icons */}
                        <div className={styles.footerIcons}>
                            <Icon iconName="Mail" className={styles.icon} />
                            <Icon iconName="Search" className={styles.icon} />
                            <Icon iconName="Heart" className={styles.icon} />
                        </div>
                    </div>
                ))
            ) : (
                <div>No Anniversary</div>
            )}
        </div>
    );
};

export default CardView;
