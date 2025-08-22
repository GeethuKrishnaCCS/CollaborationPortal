import * as React from 'react';
import { ToolbarProps, View, NavigateAction } from 'react-big-calendar';
import styles from './CollaborationCalendar.module.scss';

type CalendarEvent = {
  title: string;
  start: Date;
  end: Date;
  allDay?: boolean;
  resource?: any;
};

type CustomView = View | 'year';

interface CustomToolbarProps extends ToolbarProps<CalendarEvent, object> {
  onView: (view: CustomView) => void;
}

const CustomToolbar: React.FC<CustomToolbarProps> = ({
  label,
  onView,
  view,
  onNavigate,
  date
}) => {
  const handleViewChange = (v: CustomView) => {
    onView(v);
  };

  const handleNavigate = (action: NavigateAction) => {
    onNavigate(action);
  };

  const shouldShowNavigation = view === 'day' || view === 'week' || view === 'month';

  const formatWeekRange = (currentDate: Date): string => {
    const start = new Date(currentDate);
    const day = start.getDay(); 
    const diffToSunday = -day; 
    start.setDate(start.getDate() + diffToSunday);

    const end = new Date(start);
    end.setDate(start.getDate() + 6); 

    const startStr = start.toLocaleDateString('default', {
      month: 'short',
      day: 'numeric',
      year: 'numeric'
    });

    const endStr = end.toLocaleDateString('default', {
      month: 'short',
      day: 'numeric',
      year: 'numeric'
    });

    return `${startStr} – ${endStr}`;
  };

  return (
    <div className={styles.toolbarContainer}>
      {shouldShowNavigation && (
        <div className={styles.navControls}>
          {view === 'day' && (
            <>
              <span className={styles.dayLabel}>
                {date.toLocaleDateString('default', {
                  weekday: 'long',
                  day: 'numeric',
                  month: 'long',
                  year: 'numeric',
                })}
              </span>
              <div className={styles.navButtons}>
                <button onClick={() => handleNavigate('PREV')} className={styles.navButton} aria-label='Previous'>
                    <svg width="8" height="15" viewBox="0 0 8 15" fill="none" xmlns="http://www.w3.org/2000/svg">
                      <path d="M6.65685 1.92213L1 7.57898L6.65685 13.2358" stroke="#333333" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
                    </svg>
                </button>
                <button onClick={() => handleNavigate('NEXT')} className={styles.navButton} aria-label='Next'>
                    <svg width="9" height="15" viewBox="0 0 9 15" fill="none" xmlns="http://www.w3.org/2000/svg">
                      <path d="M1.97059 1.92213L7.62744 7.57898L1.97059 13.2358" stroke="#333333" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
                    </svg>
                </button>
              </div>
            </>
          )}
          {view === 'week' && (
            <>
              <span className={styles.weekLabel}>{formatWeekRange(date)}</span>
              <div className={styles.navButtons}>
                <button onClick={() => handleNavigate('PREV')} className={styles.navButton} aria-label='Previous'>
                    <svg width="8" height="15" viewBox="0 0 8 15" fill="none" xmlns="http://www.w3.org/2000/svg">
                      <path d="M6.65685 1.92213L1 7.57898L6.65685 13.2358" stroke="#333333" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
                    </svg>
                </button>
                <button onClick={() => handleNavigate('NEXT')} className={styles.navButton} aria-label='Next'>
                    <svg width="9" height="15" viewBox="0 0 9 15" fill="none" xmlns="http://www.w3.org/2000/svg">
                      <path d="M1.97059 1.92213L7.62744 7.57898L1.97059 13.2358" stroke="#333333" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
                    </svg>
                </button>
              </div>
            </>
          )}
          {view === 'month' && (
            <>
              <span className={styles.dateLabel}>
                {date.toLocaleString('default', {
                  month: 'long',
                 year: 'numeric'
                })}
              </span>
              <div className={styles.navButtons}>
                <button onClick={() => handleNavigate('PREV')} className={styles.navButton} aria-label='Previous'>
                    <svg width="8" height="15" viewBox="0 0 8 15" fill="none" xmlns="http://www.w3.org/2000/svg">
                      <path d="M6.65685 1.92213L1 7.57898L6.65685 13.2358" stroke="#333333" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
                    </svg>
                </button>
                <button onClick={() => handleNavigate('NEXT')} className={styles.navButton} aria-label='Next'>
                    <svg width="9" height="15" viewBox="0 0 9 15" fill="none" xmlns="http://www.w3.org/2000/svg">
                      <path d="M1.97059 1.92213L7.62744 7.57898L1.97059 13.2358" stroke="#333333" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
                    </svg>
                </button>
              </div>
            </>
          )}
        </div>
      )}

      {!shouldShowNavigation && (
        <span className={styles.dateLabel}>{label}</span>
      )}

      <div className={styles.viewSelector}>
        {['day', 'week', 'month', 'year'].map((v) => (
          <button
            key={v}
            className={`${styles.viewButton} ${view === v ? styles.active : ''}`}
            onClick={() => handleViewChange(v as CustomView)}
          >
            {v.charAt(0).toUpperCase() + v.slice(1)}
          </button>
        ))}
      </div>
    </div>
  );
};

export default CustomToolbar;