import { Birthday } from '../types';
import * as React from 'react';
import styles from './BirthdayView.module.scss';

export interface BirthdayViewProps {
    birthdays: Birthday[];
    userDisplayName: string;
    layout: 'one-column' | 'two-columns' | 'three-columns';
    description: string;
    environmentMessage: string;
  }
  
  export const BirthdayView: React.FC<BirthdayViewProps> = ({ birthdays,userDisplayName, layout ,description,environmentMessage}) => {
    return (
      <div>
        <div className={styles[layout as keyof typeof styles]}>
          {birthdays.map(birthday => (
            <div key={birthday.name}>
              {birthday.name} - {birthday.date.toDateString()}
            </div>
          ))}
        </div>
        <div className={styles[layout as keyof typeof styles]}>
        <h2>{description}</h2>
        <p>{environmentMessage}</p>
        <p>{`Hello, ${userDisplayName}!`}</p>
        </div>
      </div>
    );
  };