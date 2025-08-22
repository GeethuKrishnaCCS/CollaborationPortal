import * as React from 'react';

import {
  Calendar,
  momentLocalizer,
  View,
  DateLocalizer,
} from 'react-big-calendar';
import 'react-big-calendar/lib/css/react-big-calendar.css';
import styles from './CustomYearView.module.scss';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome'; // Import FontAwesomeIcon
import { faCalendar } from '@fortawesome/free-solid-svg-icons';
import * as moment from 'moment';
type CustomView = View | 'year';

interface CalendarEvent {
  title: string;
  start: Date;
  end: Date;
  allDay?: boolean;
  resource?: any;
}

interface CustomYearViewProps {
  events: CalendarEvent[];
  onViewChange: (view: CustomView) => void;
  onMouseEnter?: (event: CalendarEvent, e: React.MouseEvent<HTMLElement, MouseEvent>) => void;
  onMouseLeave?: () => void;
  calendarBackgroundColor?: string;
}

const localizer: DateLocalizer = momentLocalizer(moment);

const CustomYearView: React.FC<CustomYearViewProps> = ({
  events,
  onViewChange,
  onMouseEnter,
  onMouseLeave,
  calendarBackgroundColor,
}) => {
  // Current date for reference, used in renderUpcomingView
  const today = moment().startOf('day'); 

  // Log events for debugging
  React.useEffect(() => {
    console.log('Events:', events);
  }, [events]);

  // Precompute a map of dates to events for faster lookup
  const eventDateMap = React.useMemo(() => {
    const map = new Map<string, CalendarEvent[]>();
    events.forEach(event => {
      const dateKey = moment(event.start).format('YYYY-MM-DD');
      const existingEvents = map.get(dateKey) || [];
      map.set(dateKey, [...existingEvents, event]);
    });
    return map;
  }, [events]);

  const [currentYear, setCurrentYear] = React.useState(moment().year());
  const currentMonth = moment().month();
  const initialIndex = currentMonth === 11 ? 10 : currentMonth;
  const [currentIndex, setCurrentIndex] = React.useState(initialIndex);
  const [selectedMonth, setSelectedMonth] = React.useState<moment.Moment | null>(null);
  const [monthlyEvents, setMonthlyEvents] = React.useState<CalendarEvent[]>([]);
  const [isEventView, setIsEventView] = React.useState(false);
  const [isFutureView, setIsFutureView] = React.useState(false);
  const [isUpcomingView, setIsUpcomingView] = React.useState(false);
  const [isMonthUpcomingView, setIsMonthUpcomingView] = React.useState(false);
  const [upcomingStartDate, setUpcomingStartDate] = React.useState<Date | null>(null);
  const [popupPosition, setPopupPosition] = React.useState<{ top: string; left: string }>({ top: '0px', left: '0px' });
  const [isPopupVisible, setIsPopupVisible] = React.useState(false);
  const [dayEvents, setDayEvents] = React.useState<CalendarEvent[]>([]);
  const [futureEvents, setFutureEvents] = React.useState<CalendarEvent[]>([]);
  const [monthPopupEvents, setMonthPopupEvents] = React.useState<CalendarEvent[]>([]);
  const [isMonthPopupVisible, setIsMonthPopupVisible] = React.useState(false);
  const [monthPopupTitle, setMonthPopupTitle] = React.useState('');
  const [hoverPopupPosition, setHoverPopupPosition] = React.useState<{ top: string; left: string }>({ top: '0px', left: '0px' });
  const [isHoverPopupVisible, setIsHoverPopupVisible] = React.useState(false);
  const [hoverDateEvents, setHoverDateEvents] = React.useState<CalendarEvent[]>([]);

  
  const [hoverDate, setHoverDate] = React.useState<string>('');

  const months = React.useMemo(
    () => Array.from({ length: 12 }, (_, i) => moment().year(currentYear).month(i).startOf('month')),
    [currentYear]
  );

  const visibleMonths = months.slice(currentIndex, currentIndex + 2);

  const handlePrev = () => {
    if (currentIndex > 0) {
      setCurrentIndex(currentIndex - 2);
    }
  };

  const handleNext = () => {
    if (currentIndex < months.length - 2) {
      setCurrentIndex(currentIndex + 2);
    }
  };

  const handlePrevYear = () => {
    setCurrentYear((prev: any) => prev - 1);
    setCurrentIndex(0);
  };

  const handleNextYear = () => {
    setCurrentYear((prev: any) => prev + 1);
    setCurrentIndex(0);
  };

  const hasEventOnDate = (date: Date): boolean => {
    const dateKey = moment(date).format('YYYY-MM-DD');
    const hasEvent = !!eventDateMap.get(dateKey);
    console.log(`Date: ${dateKey}, Has Event: ${hasEvent}`);
    return hasEvent;
  };

  const showEventPopup = (
    eventsOnDay: CalendarEvent[],
    position: { top: string; left: string },
    date: Date
  ) => {
    setPopupPosition(position);
    setIsPopupVisible(true);
    setDayEvents(eventsOnDay);
    setHoverDate(moment(date).format('MMMM DD, YYYY'));
  };

  const handleEventClick = (
    event: CalendarEvent,
    e: React.MouseEvent<HTMLElement, MouseEvent>
  ) => {
    const sourceEvents = isUpcomingView || isMonthUpcomingView ? events : monthlyEvents;
    const selectedDate = moment(event.start).startOf('day');
    const eventsOnSameDay = sourceEvents.filter(ev =>
      moment(ev.start).isSame(selectedDate, 'day')
    );
    const rect = e.currentTarget.getBoundingClientRect();
    const popupWidth = 250;
    const popupHeight = 200;
    const spacing = 10;

    let top = rect.bottom + window.scrollY + spacing;
    let left = rect.left + window.scrollX;

    if (top + popupHeight > window.innerHeight + window.scrollY) {
      top = rect.top + window.scrollY - popupHeight - spacing;
    }
    if (left + popupWidth > window.innerWidth + window.scrollX) {
      left = window.innerWidth + window.scrollX - popupWidth - spacing;
    }

    const position = {
      top: `${Math.max(top, 0)}px`,
      left: `${Math.max(left, 0)}px`,
    };

    showEventPopup(eventsOnSameDay, position, event.start);
  };

  const handleDateClick = (slotInfo: { start: Date; end: Date; slots: Date[] }) => {
    const date = slotInfo.start;
    if (hasEventOnDate(date)) {
      const dateKey = moment(date).format('YYYY-MM-DD');
      const dateEvents = eventDateMap.get(dateKey) || [];
      const element = document.querySelector(`[data-date="${dateKey}"]`);
      let position = { top: '0px', left: '0px' };

      if (element) {
        const rect = element.getBoundingClientRect();
        const popupWidth = 250;
        const popupHeight = 200;
        const spacing = 10;

        let top = rect.bottom + window.scrollY + spacing;
        let left = rect.left + window.scrollX;

        if (top + popupHeight > window.innerHeight + window.scrollY) {
          top = rect.top + window.scrollY - popupHeight - spacing;
        }
        if (left + popupWidth > window.innerWidth + window.scrollX) {
          left = window.innerWidth + window.scrollX - popupWidth - spacing;
        }

        position = {
          top: `${Math.max(top, 0)}px`,
          left: `${Math.max(left, 0)}px`,
        };
      }

      showEventPopup(dateEvents, position, date);
    }
  };

  const closePopup = () => {
    setIsPopupVisible(false);
    setDayEvents([]);
    setHoverDate('');
  };

  const closeEventView = () => {
    setIsEventView(false);
    setSelectedMonth(null);
    setMonthlyEvents([]);
  };

  const handleYearClick = () => {
    const yearEvents = events.filter(event =>
      moment(event.start).year() === currentYear
    );
    setFutureEvents(yearEvents);
    setIsFutureView(true);
  };

  const handleBackToYear = () => {
    setIsFutureView(false);
  };

  const toggleUpcomingView = () => {
    setIsUpcomingView(!isUpcomingView);
    setIsMonthUpcomingView(false);
    setUpcomingStartDate(null);
  };

  const toggleMonthUpcomingView = (month: moment.Moment) => {
    setIsMonthUpcomingView(true);
    setUpcomingStartDate(month.toDate());
  };

  const closeMonthUpcomingView = () => {
    setIsMonthUpcomingView(false);
    setUpcomingStartDate(null);
  };

  const handleDateHover = (date: Date, e: React.MouseEvent<HTMLElement, MouseEvent>) => {
    console.log(`Hovering date: ${moment(date).format('YYYY-MM-DD')}`);
    if (hasEventOnDate(date)) {
      console.log('Events found for date');
      const rect = e.currentTarget.getBoundingClientRect();
      const popupWidth = 250;
      const popupHeight = 200;
      const spacing = 10;

      let top = rect.bottom + window.scrollY + spacing;
      let left = rect.left + window.scrollX;

      if (top + popupHeight > window.innerHeight + window.scrollY) {
        top = rect.top + window.scrollY - popupHeight - spacing;
      }
      if (left + popupWidth > window.innerWidth + window.scrollX) {
        left = window.innerWidth + window.scrollX - popupWidth - spacing;
      }

      const position = {
        top: `${Math.max(top, 0)}px`,
        left: `${Math.max(left, 0)}px`,
      };

      const dateKey = moment(date).format('YYYY-MM-DD');
      const dateEvents = eventDateMap.get(dateKey) || [];
      console.log('Date Events:', dateEvents);

      setHoverPopupPosition(position);
      setHoverDateEvents(dateEvents);
      setHoverDate(moment(date).format('MMMM DD, YYYY'));
      setIsHoverPopupVisible(true);
    }
  };

  const handleDateLeave = () => {
    console.log('Mouse left date');
    setIsHoverPopupVisible(false);
    setHoverDateEvents([]);
    setHoverDate('');
  };

  const renderUpcomingView = (startDate: Date | null = null) => {
    // For month view, start from the later of today or the month's start; for global view, start from today
    const monthStart = startDate ? moment(startDate).startOf('month') : null;
    const referenceDate = monthStart && monthStart.isAfter(today, 'day') ? monthStart : today;
    const monthEnd = startDate ? moment(startDate).endOf('month') : null;

    const upcomingEvents = events
      .filter(event => {
        const eventDate = moment(event.start);
        // Include events on or after referenceDate
        const isAfterOrOnReference = eventDate.isSameOrAfter(referenceDate, 'day');
        // For month-specific view, ensure events are within the month
        if (monthEnd) {
          return isAfterOrOnReference && eventDate.isSameOrBefore(monthEnd, 'day');
        }
        // For global view, no month restriction
        return isAfterOrOnReference;
      })
      .sort((a, b) => moment(a.start).valueOf() - moment(b.start).valueOf());

    const title = startDate
      ? `Upcoming Events for ${moment(startDate).format('MMMM YYYY')}`
      : 'Upcoming Events';

    return (
      <div className={styles.upcomingEventsBox}>
        <div className={styles.eventHeader}>
            <h3>{title}</h3>
            <button
              className={styles.backButton}
              onClick={startDate ? closeMonthUpcomingView : toggleUpcomingView}
            >
              ← Back
            </button>
        </div>
        {upcomingEvents.length > 0 ? (
          <ul className={styles.eventListScrollable}>
            {upcomingEvents.map((event, index) => (
              <li
                key={index}
                className={styles.eventItem}
                onClick={(e) => handleEventClick(event, e)}
                onMouseEnter={(e) => onMouseEnter?.(event, e)}
                onMouseLeave={() => onMouseLeave?.()}
              >
                <strong>{event.title}</strong>
                <p>{moment(event.start).format('MMM DD, hh:mm A')} - {moment(event.end).format('hh:mm A')}</p>
              </li>
            ))}
          </ul>
        ) : (
          <p>No upcoming events.</p>
        )}
      </div>
    );
  };

  const renderFutureView = () => {
    const monthsOfYear = Array.from({ length: 12 }, (_, i) =>
      moment().year(currentYear).month(i).startOf('month')
    );

    return (
      <div className={styles.futureViewContainer}>
        <div className={styles.headerSection}>
          <h3>Events in {currentYear}</h3>
          <button className={styles.backButton} onClick={handleBackToYear}>← Back to Year View</button>
        </div>
        <div className={styles.monthGrid}>
          {monthsOfYear.map((month, idx) => {
            const monthly = futureEvents
              .filter(event => moment(event.start).isSame(month, 'month'))
              .sort((a, b) => moment(a.start).valueOf() - moment(b.start).valueOf());

            const visibleEvents = monthly.slice(0, 4);
            const remainingCount = monthly.length - 4;

            return (
              <div key={idx} className={styles.futureMonthBox}>
                <h4>{month.format('MMMM')}</h4>
                {monthly.length > 0 ? (
                  <ul className={styles.eventListScrollable}>
                    {visibleEvents.map((event, i) => (
                      <li key={i} className={styles.eventItem}>
                        <strong>{event.title}</strong>
                        <p>{moment(event.start).format('MMM DD, hh:mm A')} - {moment(event.end).format('hh:mm A')}</p>
                      </li>
                    ))}
                    {remainingCount > 0 && (
                      <li
                        className={styles.moreItem}
                        onClick={(e) => {
                          e.stopPropagation();
                          setMonthPopupEvents(monthly.slice(4));
                          setMonthPopupTitle(month.format('MMMM YYYY'));
                          const monthBox = e.currentTarget.closest(`.${styles.futureMonthBox}`) as HTMLElement;
                          if (monthBox) {
                            const rect = monthBox.getBoundingClientRect();
                            const popupWidth = 250;
                            const popupHeight = 200;
                            const spacing = 10;

                            let top = rect.bottom + window.scrollY + spacing;
                            let left = rect.left + window.scrollX;

                            if (top + popupHeight > window.innerHeight + window.scrollY) {
                              top = rect.top + window.scrollY - popupHeight - spacing;
                            }
                            if (left + popupWidth > window.innerWidth + window.scrollX) {
                              left = window.innerWidth + window.scrollX - popupWidth - spacing;
                            }

                            setPopupPosition({ top: `${top}px`, left: `${left}px` });
                          }
                          setIsMonthPopupVisible(true);
                        }}
                      >
                        +{remainingCount} more
                      </li>
                    )}
                  </ul>
                ) : (
                  <p>No events</p>
                )}
              </div>
            );
          })}
        </div>
      </div>
    );
  };

  const firstEventPerDayMap = new Map<string, CalendarEvent>();
  monthlyEvents.forEach((event) => {
    const dateKey = moment(event.start).format('YYYY-MM-DD');
    if (!firstEventPerDayMap.has(dateKey)) {
      firstEventPerDayMap.set(dateKey, event);
    }
  });

  const allEvents = Array.from(firstEventPerDayMap.values());


  const headerStyles = calendarBackgroundColor
    ? `
        .${styles.headerSection} {
          background: ${calendarBackgroundColor} !important;
        }
      `
    : '';
  return (
    <div className={styles.yearViewContainer}>
      {headerStyles && <style>{headerStyles}</style>}
      {isUpcomingView ? (
        renderUpcomingView()
      ) : isMonthUpcomingView && upcomingStartDate ? (
        renderUpcomingView(upcomingStartDate)
      ) : isFutureView ? (
        renderFutureView()
      ) : (
        <>
          <div className={styles.headerSection}>
            <div className={styles.leftSection}>
              <div className={styles.monthNav}>
                <button onClick={handlePrev} disabled={currentIndex === 0} className={styles.navButton}>‹</button>
                <span className={styles.monthTitle}>
                  {visibleMonths.map(m => m.format('MMMM')).join(' – ')}
                </span>
                <button onClick={handleNext} disabled={currentIndex >= months.length - 2} className={styles.navButton}>›</button>
              </div>
              <div className={styles.yearNav}>
                <button onClick={handlePrevYear} className={styles.navButton}>«</button>
                <span
                  className={styles.yearTitle}
                  onClick={handleYearClick}
                  style={{ cursor: 'pointer', textDecoration: 'underline', margin: '0 8px' }}
                >
                  {currentYear}
                </span>
                <button onClick={handleNextYear} className={styles.navButton}>»</button>
              </div>
            </div>
            <div className={styles.rightSection}>
              <button className={styles.viewButton} onClick={() => onViewChange('day')}>Day</button>
              <button className={styles.viewButton} onClick={() => onViewChange('week')}>Week</button>
              <button className={styles.viewButton} onClick={() => onViewChange('month')}>Month</button>
              <button className={`${styles.viewButton} ${styles.active}`} onClick={() => onViewChange('year')}>Year</button>
            </div>
          </div>

          {!isEventView && (
            <div className={styles.monthGrid}>
              {visibleMonths.map((month) => (
                <div key={`${month.year()}-${month.month()}`} className={styles.monthContainer}>
                  <div className={styles.monthHeader}>
                    <h4 className={styles.monthTitle}>{month.format('MMMM YYYY')}</h4>
                    <button
                      className={styles.monthUpcomingIcon}
                      onClick={() => toggleMonthUpcomingView(month)}
                      title={`Show Upcoming Events for ${month.format('MMMM YYYY')}`}
                      aria-label={`Show upcoming events for ${month.format('MMMM YYYY')}`}
                    >
                      <FontAwesomeIcon icon={faCalendar} />
                    </button>
                  </div>
                  <Calendar
                    localizer={localizer}
                    events={[]}
                    startAccessor="start"
                    endAccessor="end"
                    views={['month']}
                    defaultDate={month.toDate()}
                    style={{ height: '250px', width: '100%' }}
                    toolbar={false}
                    selectable={true}
                    onSelectSlot={handleDateClick}
                    dayPropGetter={(date) => {
                      const isEventDay = hasEventOnDate(date);
                      const dateKey = moment(date).format('YYYY-MM-DD');
                      return {
                        className: isEventDay ? styles.eventDay : '',
                        'data-date': dateKey,
                        style: isEventDay
                          ? { backgroundColor: '#d4f0ff', cursor: 'pointer', position: 'relative' }
                          : {},
                        onMouseEnter: isEventDay
                          ? (e: React.MouseEvent<HTMLElement, MouseEvent>) => handleDateHover(date, e)
                          : undefined,
                        onMouseLeave: isEventDay ? handleDateLeave : undefined,
                      };
                    }}
                  />
                </div>
              ))}
            </div>
          )}

          {isEventView && selectedMonth && (
            <div className={styles.upcomingEventsBox}>
              <h3>Events for {selectedMonth.format('MMMM YYYY')}</h3>
              <button className={styles.backButton} onClick={closeEventView}>← Back</button>
              {allEvents.length > 0 ? (
                <ul className={styles.eventListScrollable}>
                  {allEvents.map((event, index) => (
                    <li
                      key={index}
                      className={styles.eventItem}
                      onClick={(e) => handleEventClick(event, e)}
                      onMouseEnter={(e) => onMouseEnter?.(event, e)}
                      onMouseLeave={() => onMouseLeave?.()}
                    >
                      <strong>{event.title}</strong>
                      <p>{moment(event.start).format('MMM DD, hh:mm A')} - {moment(event.end).format('hh:mm A')}</p>
                    </li>
                  ))}
                </ul>
              ) : (
                <p>No events this month.</p>
              )}
            </div>
          )}
        </>
      )}

      {isPopupVisible && (
        <div className={styles.hoverPopup} style={{ top: popupPosition.top, left: popupPosition.left }}>
          {console.log('Rendering popup with events:', dayEvents)}
          <h3>{hoverDate}</h3>
          <ul className={styles.eventListScrollable}>
            {dayEvents.map((ev, idx) => (
              <li key={idx}>
                <strong>{ev.title}</strong>
                <p>{moment(ev.start).format('hh:mm A')} - {moment(ev.end).format('hh:mm A')}</p>
              </li>
            ))}
          </ul>
          <button onClick={closePopup} className={styles.closePopupButton} aria-label="Close popup">
            <svg width="15" height="16" viewBox="0 0 15 16" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M14.4247 0.817749C14.4637 0.833352 14.5027 0.852856 14.5417 0.868458C15.0059 1.05179 15.158 1.6174 14.846 2.00747C14.7992 2.06598 14.7485 2.11669 14.6938 2.1674C12.6889 4.17238 10.6839 6.18125 8.67502 8.18232C8.5775 8.27984 8.5697 8.32665 8.67112 8.42807C10.6956 10.4447 12.7123 12.4653 14.7328 14.482C14.9396 14.6887 15.0566 14.9228 14.9864 15.2192C14.8616 15.7692 14.218 15.9799 13.7928 15.6093C13.7499 15.5703 13.707 15.5274 13.6679 15.4884C11.6591 13.4795 9.6502 11.4706 7.64133 9.45786C7.53991 9.35644 7.4931 9.36424 7.39558 9.46176C5.375 11.4901 3.35052 13.5107 1.32995 15.5352C0.943776 15.9214 0.389873 15.8746 0.136326 15.4338C0.0973182 15.3636 0.0661124 15.2855 0.0349066 15.2114C0.0349066 15.1139 0.0349066 15.0164 0.0349066 14.9189C0.10512 14.7004 0.257248 14.5366 0.417178 14.3767C2.40655 12.3912 4.38812 10.4057 6.37749 8.42417C6.47111 8.33055 6.47111 8.28374 6.37749 8.19012C4.38032 6.20465 2.39095 4.21138 0.401575 2.22591C0.245546 2.06988 0.101219 1.90995 0.0310059 1.69541C0.0310059 1.59789 0.0310059 1.50038 0.0310059 1.40286C0.124623 1.1064 0.31966 0.911366 0.616115 0.817749C0.713633 0.817749 0.811151 0.817749 0.90867 0.817749C1.12321 0.884061 1.28314 1.03229 1.43917 1.18832C3.42464 3.17769 5.41791 5.16706 7.40338 7.16033C7.497 7.25395 7.54381 7.25395 7.63742 7.16033C9.6229 5.16706 11.6162 3.17769 13.6016 1.18832C13.7577 1.03229 13.9176 0.887962 14.1321 0.817749C14.2297 0.817749 14.3272 0.817749 14.4247 0.817749Z" fill="currentcolor"/>
            </svg>
          </button>
        </div>
      )}

      {isMonthPopupVisible && monthPopupEvents.length > 0 && (
        <div className={styles.hoverPopup} style={{ top: popupPosition.top, left: popupPosition.left }}>
          <h3>{monthPopupTitle}</h3>
          <ul className={styles.eventListScrollable}>
            {monthPopupEvents.map((event, idx) => (
              <li key={idx}>
                <strong>{event.title}</strong>
                <p>{moment(event.start).format('MMM DD, hh:mm A')} – {moment(event.end).format('hh:mm A')}</p>
              </li>
            ))}
          </ul>
          <button onClick={() => setIsMonthPopupVisible(false)} className={styles.closePopupButton} aria-label='Close Popup'>
            <svg width="15" height="16" viewBox="0 0 15 16" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M14.4247 0.817749C14.4637 0.833352 14.5027 0.852856 14.5417 0.868458C15.0059 1.05179 15.158 1.6174 14.846 2.00747C14.7992 2.06598 14.7485 2.11669 14.6938 2.1674C12.6889 4.17238 10.6839 6.18125 8.67502 8.18232C8.5775 8.27984 8.5697 8.32665 8.67112 8.42807C10.6956 10.4447 12.7123 12.4653 14.7328 14.482C14.9396 14.6887 15.0566 14.9228 14.9864 15.2192C14.8616 15.7692 14.218 15.9799 13.7928 15.6093C13.7499 15.5703 13.707 15.5274 13.6679 15.4884C11.6591 13.4795 9.6502 11.4706 7.64133 9.45786C7.53991 9.35644 7.4931 9.36424 7.39558 9.46176C5.375 11.4901 3.35052 13.5107 1.32995 15.5352C0.943776 15.9214 0.389873 15.8746 0.136326 15.4338C0.0973182 15.3636 0.0661124 15.2855 0.0349066 15.2114C0.0349066 15.1139 0.0349066 15.0164 0.0349066 14.9189C0.10512 14.7004 0.257248 14.5366 0.417178 14.3767C2.40655 12.3912 4.38812 10.4057 6.37749 8.42417C6.47111 8.33055 6.47111 8.28374 6.37749 8.19012C4.38032 6.20465 2.39095 4.21138 0.401575 2.22591C0.245546 2.06988 0.101219 1.90995 0.0310059 1.69541C0.0310059 1.59789 0.0310059 1.50038 0.0310059 1.40286C0.124623 1.1064 0.31966 0.911366 0.616115 0.817749C0.713633 0.817749 0.811151 0.817749 0.90867 0.817749C1.12321 0.884061 1.28314 1.03229 1.43917 1.18832C3.42464 3.17769 5.41791 5.16706 7.40338 7.16033C7.497 7.25395 7.54381 7.25395 7.63742 7.16033C9.6229 5.16706 11.6162 3.17769 13.6016 1.18832C13.7577 1.03229 13.9176 0.887962 14.1321 0.817749C14.2297 0.817749 14.3272 0.817749 14.4247 0.817749Z" fill="currentcolor"/>
            </svg>
          </button>
        </div>
      )}

      {isHoverPopupVisible && hoverDateEvents.length > 0 && (
        <div className={styles.dateHoverPopup} style={{ top: hoverPopupPosition.top, left: popupPosition.left }}>
          <h3>{hoverDate}</h3>
          <ul className={styles.eventListScrollable}>
            {hoverDateEvents.map((event, idx) => (
              <li key={idx}>
                <strong>{event.title}</strong>
                <p>{moment(event.start).format('hh:mm A')} – {moment(event.end).format('hh:mm A')}</p>
              </li>
            ))}
          </ul>
        </div>
      )}
    </div>
  );
};

export default CustomYearView;