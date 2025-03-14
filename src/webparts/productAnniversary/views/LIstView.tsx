import * as React from 'react';
import styles from './ListView.module.scss';
import { IIconProps, IconButton } from '@fluentui/react';
import * as moment from 'moment';
import { WebPartContext } from '@microsoft/sp-webpart-base';

interface IListViewProps {
    greetings: any[];
    headingfontcolor: string;
    bdayGreetingWish: string;
    WorkGreetingWish: string;
    weddingGreetingWish: string;
    AnniversayNoOfItemDisplay: string;
    context: WebPartContext;
    contentfontcolor: string;
}

const ListView: React.FC<IListViewProps> = (props) => {

    const { greetings, AnniversayNoOfItemDisplay, contentfontcolor } = props;
    const itemsPerPage = parseInt(AnniversayNoOfItemDisplay) || 3;

    const [currentIndex, setCurrentIndex] = React.useState(0);

    const handleScrollUp = () => {
        setCurrentIndex((prevIndex) => Math.max(prevIndex - itemsPerPage, 0));
    };

    const handleScrollDown = () => {
        setCurrentIndex((prevIndex) => Math.min(prevIndex + itemsPerPage, greetings.length - 1));
    };

    const ChevronUp: IIconProps = { iconName: 'ChevronUp' };
    const ChevronDown: IIconProps = { iconName: 'ChevronDown' };

    const displayedItems = greetings.slice(currentIndex, currentIndex + itemsPerPage);
    console.log('displayedItems: ', displayedItems);

    return (
        <div className={styles.box}>

            <div className={styles.uparrow}>
                <IconButton
                    iconProps={ChevronUp}
                    ariaLabel="Scroll up"
                    onClick={handleScrollUp}
                    disabled={currentIndex === 0}
                    className={styles.customIconButton}
                />
            </div>

            <div className={styles.employee}>
                {displayedItems.length > 0 ? (
                    displayedItems.map((item: any, index: number) => (
                        <div className={styles.persondiv} key={index}>
                            <div className={styles.Profilecard}>
                                <div className={styles.ImgContainer}>
                                    <img
                                        src={item.employeeImage || ''}
                                        className={styles.Image}
                                        alt={`Profile picture of ${item.Employee.Title}`}
                                    />
                                </div>
                                <div className={styles.secondarycard}>
                                    <div  style={{ color: contentfontcolor }} className={styles.Namecard}>{item.Employee.Title}</div>
                                    <div className={styles.secondarytextstyle}>
                                        {item.type === 'Birthday'
                                            ? `${props.bdayGreetingWish} on ${moment(item.DateOfBirth).format('MMM DD')}`
                                            : item.type === 'Work Anniversary'
                                                ? `${props.WorkGreetingWish} on ${moment(item.DateOfJoining).format('MMM DD')}`
                                                : item.type === 'Wedding Anniversary'
                                                    ? `${props.weddingGreetingWish} on ${moment(item.DateOfWedding).format('MMM DD')}`
                                                    : ''}
                                    </div>
                                </div>
                            </div>
                        </div>
                    ))
                ) : (
                    <div>No Anniversaries Today</div>
                )}
            </div>

            <div className={styles.downarrow}>
                <IconButton
                    iconProps={ChevronDown}
                    ariaLabel="Scroll down"
                    onClick={handleScrollDown}
                    disabled={currentIndex >= greetings.length - parseInt(props.AnniversayNoOfItemDisplay)}
                    className={styles.customIconButton}
                />
            </div>
        </div>
    );
};

export default ListView;
