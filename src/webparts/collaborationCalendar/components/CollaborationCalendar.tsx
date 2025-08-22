import * as React from 'react';
import { useEffect, useState } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
//import { setupPnP } from '../../../common/PnP/PnPConfig';
import { setupPnP } from '../pnpjsConfig';
import {
  Calendar as BigCalendar,
  momentLocalizer,
  ToolbarProps,
  Views,
  View,
  EventWrapperProps,
  SlotInfo,
  HeaderProps
} from 'react-big-calendar';
import 'react-big-calendar/lib/css/react-big-calendar.css';
import styles from './CollaborationCalendar.module.scss';
import CustomToolbar from './CustomToolbar';
import CustomYearView from './CustomYearView';
import * as moment from 'moment-timezone';

import { IEventItem } from '../models/IEventItem';
import {
  getEventsFromList,
  saveEventToList,
  updateEventInList,
  saveEventToOutlook
} from '../services/SPService';
import { PeoplePicker, PrincipalType, IPeoplePickerContext } from '@pnp/spfx-controls-react/lib/PeoplePicker';

const localizer = momentLocalizer(moment);

export interface ICollaborationCalendarProps {
  title: string;
  siteUrl: string;
  listName: string;
  context: WebPartContext;
  showOutlookEvents: boolean;
  showSpEventsOnly: boolean;
  outlookEventColor: string;
  spEventColor: string;
  calendarBackgroundColor: string;
  allowedGroups: any[];
}

type CalendarEvent = {
  id?: number;
  title: string;
  start: Date;
  end: Date;
  resource?: IEventItem;
};

type CustomView = View | "year";

const CollaborationCalendar: React.FC<ICollaborationCalendarProps> = ({
  listName,
  context,
  showOutlookEvents,
  showSpEventsOnly,
  outlookEventColor,
  spEventColor,
  calendarBackgroundColor,
  allowedGroups
}) => {
  useEffect(() => {
    setupPnP(context);
  }, [context]);

  const [calendarEvents, setCalendarEvents] = useState<CalendarEvent[]>([]);
  const [currentView, setCurrentView] = useState<CustomView>(Views.DAY);
  const [hoveredEvent, setHoveredEvent] = useState<IEventItem | null>(null);
  const [tooltipVisible, setTooltipVisible] = useState(false);
  const [tooltipPosition, setTooltipPosition] = useState<{ x: number; y: number }>({ x: 0, y: 0 });
  const [eventPopup, setEventPopup] = useState({
    show: false,
    events: [] as CalendarEvent[],
    position: { x: 0, y: 0 }
  });
  const [selectedSlot, setSelectedSlot] = useState<SlotInfo | null>(null);
  const [isEventModalOpen, setIsEventModalOpen] = useState(false);
  const [currentEvent, setCurrentEvent] = useState<CalendarEvent | null>(null);
  const [isSaving, setIsSaving] = useState(false);
  const [saveToOutlook, setSaveToOutlook] = useState(showOutlookEvents && !showSpEventsOnly);
  const [selectedPeople, setSelectedPeople] = useState<any[]>([]);
  const [canAddEditEvents, setCanAddEditEvents] = useState<boolean>(true);

  // Enhanced: Supports user, SharePoint group, and AAD group permissions
  const isCurrentUserInAudience = async (context: WebPartContext, audienceIds: string[]): Promise<boolean> => {
    if (!audienceIds || audienceIds.length === 0) {
      return true;
    }

    try {
      const currentUserLoginName = context.pageContext.user.loginName?.toLowerCase();
      const currentUserEmail = context.pageContext.user.email?.toLowerCase();

      // Separate AAD group IDs and SharePoint group IDs
      const aadGroupIds = audienceIds
        .filter(id => id && id.toString().startsWith("c:0o.c|"))
        .map(id => id.toString().split("|").pop()!);

      const spGroupIds = audienceIds
        .filter(id => id && !isNaN(Number(id)))
        .map(id => id.toString().trim());

      const userIds = audienceIds
        .filter(id => id && isNaN(Number(id)) && !id.toString().startsWith("c:0o.c|"))
        .map(id => id.toLowerCase().trim());

      // Direct user match (loginName or email)
      const isDirectUserMatch = userIds.some(id =>
        id.includes(currentUserLoginName) || id.includes(currentUserEmail)
      );
      if (isDirectUserMatch) return true;

      // SharePoint group check
      if (spGroupIds.length > 0) {
        const response = await context.spHttpClient.get(
          `${context.pageContext.web.absoluteUrl}/_api/web/currentuser/groups`,
          SPHttpClient.configurations.v1
        );
        if (!response.ok) throw new Error(`Failed to get user groups: ${response.statusText}`);
        const data = await response.json();
        const userGroups = data.value || [];
        const hasSpGroupMatch = spGroupIds.some(audienceId =>
          userGroups.some((group: any) => group.Id.toString() === audienceId)
        );
        if (hasSpGroupMatch) return true;
      }

      // AAD group check
      if (aadGroupIds.length > 0) {
        const cacheKey = `aadGroups_${aadGroupIds.join("_")}_${currentUserEmail}`;
        const cached = sessionStorage.getItem(cacheKey);
        if (cached !== null) return cached === "true";
        try {
          const graphClient = await context.msGraphClientFactory.getClient("3");
          const me = await graphClient.api("/me").get();
          for (const groupId of aadGroupIds) {
            const result = await graphClient
              .api(`/groups/${groupId}/members/$/microsoft.graph.user`)
              .filter(`id eq '${me.id}'`)
              .get();
            if (result?.value?.length > 0) {
              sessionStorage.setItem(cacheKey, "true");
              return true;
            }
          }
          sessionStorage.setItem(cacheKey, "false");
        } catch (error) {
          console.error("AAD group check error:", error);
          sessionStorage.setItem(cacheKey, "false");
        }
      }

      return false;
    } catch (error) {
      console.error("Error checking audience targeting:", error);
      return false;
    }
  };

  useEffect(() => {
    const checkUserPermission = async () => {
      try {
        console.log("Checking permissions for user:", context.pageContext.user.loginName);

        if (!allowedGroups || allowedGroups.length === 0) {
          console.log("No restrictions - allowing edit");
          setCanAddEditEvents(true);
          return;
        }

        console.log("Allowed Groups:", allowedGroups);
        const audienceIds = allowedGroups.map(group =>
          group.id?.toString() ||
          group.loginName ||
          group.email
        ).filter(Boolean);

        if (audienceIds.length === 0) {
          console.log("No valid audience IDs found");
          setCanAddEditEvents(true);
          return;
        }

        const isInAudience = await isCurrentUserInAudience(context, audienceIds);
        console.log("User in audience:", isInAudience);
        setCanAddEditEvents(isInAudience);
      } catch (error) {
        console.error("Permission check failed:", error);
        setCanAddEditEvents(false); // Fail secure
      }
    };

    checkUserPermission();
  }, [allowedGroups, context]);

  useEffect(() => {
    fetchEvents();
  }, [listName, showOutlookEvents, showSpEventsOnly]);

  const fetchOutlookEvents = async (): Promise<IEventItem[]> => {
    try {
      const client = await context.msGraphClientFactory.getClient('3');
      const response = await client.api("/me/calendar/events").top(100).get();

      return response.value.map((item: any, index: number): IEventItem => {
        const startTime = moment.tz(item.start.dateTime, item.start.timeZone).toDate();
        const endTime = moment.tz(item.end.dateTime, item.end.timeZone).toDate();

        return {
          Id: index,
          Title: item.subject,
          EventDate: startTime,
          EndDate: endTime,
          Location: item.location?.displayName || '',
          Description: item.bodyPreview || '',
          TeamMembers: item.attendees?.map((a: any) => a.emailAddress?.name) || [],
          Source: 'Outlook'
        };
      });
    } catch (error) {
      console.error("Error fetching Outlook events:", error);
      return [];
    }
  };

  const fetchEvents = async () => {
    let events: IEventItem[] = [];

    if (showSpEventsOnly) {
      console.log("Fetching SharePoint events (showSpEventsOnly is true)");
      const spEvents = await getEventsFromList(listName);
      events = [...events, ...spEvents];
    }

    if (showOutlookEvents && (!showSpEventsOnly || (showSpEventsOnly && showOutlookEvents))) {
      console.log("Fetching Outlook events (showOutlookEvents is true)");
      const outlookEvents = await fetchOutlookEvents();
      events = [...events, ...outlookEvents];
    }

    const mappedEvents: CalendarEvent[] = events.map((ev) => ({
      id: ev.Id,
      title: ev.Title,
      start: new Date(ev.EventDate),
      end: new Date(ev.EndDate),
      resource: ev
    }));

    setCalendarEvents(mappedEvents);
  };

  const handleMouseEnter = (event: CalendarEvent, e: React.MouseEvent<HTMLDivElement>) => {
    const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
    const x = rect.left + rect.width / 2;
    const y = rect.top - 10;

    setHoveredEvent(event.resource || null);
    setTooltipPosition({ x, y });
    setTooltipVisible(true);
  };

  const handleMouseLeave = () => {
    setTooltipVisible(false);
  };

  const EventWrapper: React.FC<EventWrapperProps<CalendarEvent>> = (props) => {
    return (
      <div 
        style={props.style}
        className={props.className}
        onClick={props.onClick}
        onDoubleClick={props.onDoubleClick}
      >
        {props.children}
      </div>
    );
  };

  const handleEventClick = (events: CalendarEvent[], e: React.MouseEvent<HTMLDivElement>) => {
    const rect = e.currentTarget.getBoundingClientRect();
    setEventPopup({
      show: true,
      events,
      position: {
        x: rect.left + rect.width / 2,
        y: rect.top + rect.height
      }
    });
    e.stopPropagation();
  };

  const getEventsAtSlot = (event: CalendarEvent) => {
    return calendarEvents.filter(e => 
      e.start.getTime() === event.start.getTime() && 
      e.end.getTime() === event.end.getTime()
    );
  };

  const customEventPropGetter = (event: CalendarEvent) => {
    const source = event.resource?.Source;
    let backgroundColor = '#666';

    if (source === 'Outlook') {
      backgroundColor = outlookEventColor || '#0078D4';
    } else if (source === 'SharePoint') {
      backgroundColor = spEventColor || '#107C10';
    }
    const eventsAtSlot = getEventsAtSlot(event);
    const isFirstEvent = eventsAtSlot.length > 0 ? eventsAtSlot[0]?.title === event.title : true;

    const baseStyle = {
      backgroundColor,
      color: 'white',
      borderRadius: '4px',
      border: 'none',
      padding: '2px 4px',
      width: '100%',
      height: '100%', // Full height
      minHeight: '20px', // Ensure visibility
      
    };

    if ((currentView === Views.DAY || currentView === Views.WEEK) && !isFirstEvent) {
      return {
        style: {
          display: 'none'
        }
      };
    }

    return {
      style: baseStyle
    };
  };

  const handleSelectSlot = (slotInfo: SlotInfo) => {
    if (!canAddEditEvents) {
      alert('You do not have permission to add new events.');
      return;
    }
    setSelectedSlot(slotInfo);
    setCurrentEvent(null);
    setSelectedPeople([]);
    setIsEventModalOpen(true);
  };

  const handleSelectEvent = (event: CalendarEvent) => {
    if (!canAddEditEvents) {
      alert('You do not have permission to edit events.');
      return;
    }
    setCurrentEvent(event);
    setSelectedPeople(event.resource?.TeamMembers?.map((name: string) => ({ text: name })) || []);
    setIsEventModalOpen(true);
  };

  const handleSaveEvent = async (eventData: {
    title: string;
    start: Date;
    end: Date;
    description?: string;
    location?: string;
    teamMembers?: string[];
  }) => {
    setIsSaving(true);
    try {
      if (currentEvent?.id) {
        if (currentEvent.resource?.Source === 'SharePoint') {
          console.log("Updating SharePoint event with ID:", currentEvent.id);
          await updateEventInList(listName, {
            Id: currentEvent.id,
            Title: eventData.title,
            EventDate: eventData.start,
            EndDate: eventData.end,
            Location: eventData.location || '',
            Description: eventData.description || '',
            TeamMembers: eventData.teamMembers || []
          }, context);
        } else if (currentEvent.resource?.Source === 'Outlook') {
          console.log("Updating Outlook event");
          await saveEventToOutlook(context, {
            Title: eventData.title,
            EventDate: eventData.start,
            EndDate: eventData.end,
            Location: eventData.location || '',
            Description: eventData.description || '',
            TeamMembers: eventData.teamMembers || []
          });
        }
      } else {
        if (showSpEventsOnly && !showOutlookEvents) {
          console.log("Saving new event to SharePoint (only showSpEventsOnly is true)");
          await saveEventToList(listName, {
            Title: eventData.title,
            EventDate: eventData.start,
            EndDate: eventData.end,
            Location: eventData.location || '',
            Description: eventData.description || '',
            TeamMembers: eventData.teamMembers || []
          }, context);
        } else if (showOutlookEvents && !showSpEventsOnly) {
          console.log("Saving new event to Outlook (only showOutlookEvents is true)");
          await saveEventToOutlook(context, {
            Title: eventData.title,
            EventDate: eventData.start,
            EndDate: eventData.end,
            Location: eventData.location || '',
            Description: eventData.description || '',
            TeamMembers: eventData.teamMembers || []
          });
        } else if (showSpEventsOnly && showOutlookEvents) {
          if (saveToOutlook) {
            console.log("Saving new event to Outlook (toggle is set to Outlook)");
            await saveEventToOutlook(context, {
              Title: eventData.title,
              EventDate: eventData.start,
              EndDate: eventData.end,
              Location: eventData.location || '',
              Description: eventData.description || '',
              TeamMembers: eventData.teamMembers || []
            });
          } else {
            console.log("Saving new event to SharePoint (toggle is set to SharePoint)");
            await saveEventToList(listName, {
              Title: eventData.title,
              EventDate: eventData.start,
              EndDate: eventData.end,
              Location: eventData.location || '',
              Description: eventData.description || '',
              TeamMembers: eventData.teamMembers || []
            }, context);
          }
        } else {
          console.log("Neither source enabled, defaulting to SharePoint");
          await saveEventToList(listName, {
            Title: eventData.title,
            EventDate: eventData.start,
            EndDate: eventData.end,
            Location: eventData.location || '',
            Description: eventData.description || '',
            TeamMembers: eventData.teamMembers || []
          }, context);
        }
      }
      
      await fetchEvents();
      setIsEventModalOpen(false);
      setCurrentEvent(null);
      setSelectedSlot(null);
      setSelectedPeople([]);
    } catch (error) {
      console.error("Error saving event:", error);
      alert(`Failed to save the event: ${error.message || 'Unknown error'}. Please try again.`);
    } finally {
      setIsSaving(false);
    }
  };

  const renderEventContent = (event: CalendarEvent, view: View) => {
    const eventsAtSlot = getEventsAtSlot(event);
    
    if (view === Views.MONTH) {
      return (
        <div
          onMouseEnter={(e) => handleMouseEnter(event, e)}
          onMouseLeave={handleMouseLeave}
          className={styles.monthEvent}
        >
          {event.title}
        </div>
      );
    }

    const isFirstEvent = eventsAtSlot[0]?.title === event.title;

    
    if (!isFirstEvent) return null;

    return (
      <div
        onMouseEnter={(e) => handleMouseEnter(event, e)}
        onMouseLeave={handleMouseLeave}
        className={view === Views.DAY ? styles.dayEvent : styles.weekEvent}
        onClick={(e) => handleEventClick(eventsAtSlot, e)}
        style={{
          width: '100%',
          height: '100%',
          display: 'flex',
          alignItems: 'center',
          padding: '2px 4px',
          boxSizing: 'border-box',
          overflow: 'hidden'
        }}
      >
        <span style={{
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          whiteSpace: 'nowrap',
        }}>
          {view === Views.DAY ? (
            <>
              {event.title} {moment(event.start).format('h:mm A')} - {moment(event.end).format('h:mm A')}
            </>
          ) : (
            <>
              {event.title}
            </>
          )}
          {eventsAtSlot.length > 1 && (
            <span className={styles.moreEventsBadge}>+{eventsAtSlot.length - 1}</span>
          )}
        </span>
      </div>
    );
  };

  const CustomHeader: React.FC<HeaderProps> = ({ date }) => {
    return (
      <div>
        {moment(date).format('ddd')}
      </div>
    );
  };

  const calendarStyles = calendarBackgroundColor
    ? `
      .${styles.calendarContainer} {
        background: ${calendarBackgroundColor} !important;
      }
    `
    : '';

  const peoplePickerContext: IPeoplePickerContext = {
    ...context,
    absoluteUrl: context.pageContext.web.absoluteUrl,
    spHttpClient: context.spHttpClient as any,
    msGraphClientFactory: context.msGraphClientFactory as any
  };

  return (
    <div 
      className={`${styles.calendarContainer} ${styles.customDaySlot} ${styles.customEventLabel} ${styles.customSelectedEvent} ${styles.customTimeSlotGroup} 
      ${styles.smCalendar} 
    }`}
    >
      {calendarBackgroundColor && (
        <style>{calendarStyles}</style>
      )}
      {currentView === "year" ? (
        <CustomYearView
          events={calendarEvents}
          onViewChange={setCurrentView}
          onMouseEnter={handleMouseEnter}
          onMouseLeave={handleMouseLeave}
          calendarBackgroundColor={calendarBackgroundColor}
        />
      ) : (
        <BigCalendar
          localizer={localizer}
          events={calendarEvents}
          startAccessor="start"
          endAccessor="end"
          view={currentView}
          views={{ month: true, week: true, day: true }}
          onView={(view: View) => setCurrentView(view as CustomView)}
          eventPropGetter={customEventPropGetter}
          selectable={(currentView === Views.DAY || currentView === Views.WEEK) && canAddEditEvents}
          onSelectSlot={handleSelectSlot}
          onSelectEvent={handleSelectEvent}
          components={{
            toolbar: (props:any) => (
              <CustomToolbar
                {...(props as ToolbarProps<CalendarEvent, object>)}
                onView={setCurrentView}
              />
            ),
            eventWrapper: EventWrapper,
            event: ({ event }: { event: any }) => (
              <div
                onMouseEnter={(e) => handleMouseEnter(event, e)}
                onMouseLeave={handleMouseLeave}
                className={styles.monthEvent}
              >
                {event.title}
              </div>
            ),
            week: {
              header: CustomHeader,
              event: (props:any) => renderEventContent(props.event, Views.WEEK)
            },
            month: {
              header: CustomHeader,
              event: (props:any) => renderEventContent(props.event, Views.MONTH)
            }
          }}
          style={{ height: '500px' }}
          popup={true}
          step={60}
          timeslots={1}
        />
      )}

      {tooltipVisible && hoveredEvent && (
        <div
          className={styles.eventTooltip}
          style={{
            position: 'fixed',
            top: tooltipPosition.y,
            left: tooltipPosition.x,
            transform: 'translate(-50%, -100%)'
          }}
        >
           <button
    className={styles.tooltipCloseButton}
    onClick={() => setTooltipVisible(false)}
    aria-label="Close"
    type="button"
  >
    ×
  </button>
          <div className={styles.tooltipHeader}>
            <strong>{hoveredEvent.Title}</strong>
            
          </div>
          <div className={styles.tooltipContent}>
            {hoveredEvent.Location && (
              <div className={styles.tooltipRow}>
                <span className={styles.tooltipLabel}>Location:</span>
                <span>{hoveredEvent.Location}</span>
              </div>
            )}
            <div className={styles.tooltipRow}>
              <span className={styles.tooltipLabel}>Time:</span>
              <span>
                {moment(hoveredEvent.EventDate).format('h:mm A')} - {moment(hoveredEvent.EndDate).format('h:mm A')}
              </span>
            </div>
            {hoveredEvent.Description && (
              <div className={styles.tooltipRow}>
                <span className={styles.tooltipLabel}>Description:</span>
                <span>{hoveredEvent.Description}</span>
              </div>
            )}
            {hoveredEvent.TeamMembers && hoveredEvent.TeamMembers.length > 0 && (
              <div className={styles.tooltipRow}>
                <span className={styles.tooltipLabel}>Attendees:</span>
                <span>{hoveredEvent.TeamMembers.join(', ')}</span>
              </div>
            )}
          </div>
        </div>
      )}

      {eventPopup.show && (
        <div 
          className={styles.eventsPopup}
          style={{
            left: eventPopup.position.x,
            top: eventPopup.position.y,
            transform: 'translateX(-50%)'
          }}
          onClick={(e) => e.stopPropagation()}
        >
          <div className={styles.popupHeader}>
            <h4>{eventPopup.events.length} Events</h4>
            <button 
              className={styles.closeButton}
              onClick={() => setEventPopup({...eventPopup, show: false})}
            >
              ×
            </button>
          </div>
          <div className={styles.eventsList}>
            {eventPopup.events.map((event, index) => (
              <div 
                key={index}
                className={styles.popupEventItem}
                onMouseEnter={(e) => handleMouseEnter(event, e)}
                onMouseLeave={handleMouseLeave}
              >
                <div className={styles.eventTime}>
                  {moment(event.start).format('h:mm A')} - {moment(event.end).format('h:mm A')}
                </div>
                <div className={styles.eventTitle}>{event.title}</div>
                {event.resource?.Location && (
                  <div className={styles.eventLocation}>{event.resource.Location}</div>
                )}
              </div>
            ))}
          </div>
        </div>
      )}

      {isEventModalOpen && (
        <div className={styles.eventModalOverlay} onClick={() => setIsEventModalOpen(false)}>
          <div className={styles.eventModal} onClick={(e) => e.stopPropagation()}>
            <h3>{currentEvent ? 'Edit Event' : 'New Event'}</h3>
            <form onSubmit={(e) => {
              e.preventDefault();
              const formData = new FormData(e.currentTarget);
              const teamMembers = selectedPeople.map(person => person.text);
              handleSaveEvent({
                title: formData.get('title') as string,
                start: selectedSlot?.start || currentEvent?.start || new Date(),
                end: selectedSlot?.end || currentEvent?.end || new Date(),
                description: formData.get('description') as string,
                location: formData.get('location') as string,
                teamMembers: teamMembers
              });
            }}>
              <div className={styles.formGroup}>
                <label>Title:</label>
                <input 
                  name="title" 
                  defaultValue={currentEvent?.title || ''} 
                  required 
                  aria-label='Event Title'
                />
              </div>
              <div className={styles.formGroup}>
                <label>Start:</label>
                <input
                  type="datetime-local"
                  aria-label='Event Start'
                  name="start"
                  defaultValue={moment(selectedSlot?.start || currentEvent?.start).format('YYYY-MM-DDTHH:mm')}
                  onChange={(e) => {
                    if (selectedSlot) {
                      setSelectedSlot({
                        ...selectedSlot,
                        start: new Date(e.target.value)
                      });
                    }
                  }}
                  required
                />
              </div>
              <div className={styles.formGroup}>
                <label>End:</label>
                <input
                  type="datetime-local"
                  aria-label='Event End'
                  name="end"
                  defaultValue={moment(selectedSlot?.end || currentEvent?.end).format('YYYY-MM-DDTHH:mm')}
                  onChange={(e) => {
                    if (selectedSlot) {
                      setSelectedSlot({
                        ...selectedSlot,
                        end: new Date(e.target.value)
                      });
                    }
                  }}
                  required
                />
              </div>
              <div className={styles.formGroup}>
                <label>Description:</label>
                <textarea 
                  name="description" 
                  aria-label='Event Description'
                  defaultValue={currentEvent?.resource?.Description || ''} 
                  rows={4}
                />
              </div>
              <div className={styles.formGroup}>
                <label>Location:</label>
                <input 
                  name="location" 
                  aria-label='Event Location'
                  defaultValue={currentEvent?.resource?.Location || ''} 
                />
              </div>
              <div className={styles.formGroup}>
                <label>Team Members:</label>
                <PeoplePicker
                  context={peoplePickerContext}
                  titleText=""
                  personSelectionLimit={10}
                  groupName=""
                  showtooltip={true}
                  disabled={false}
                  onChange={(items: any[]) => setSelectedPeople(items)}
                  defaultSelectedUsers={currentEvent?.resource?.TeamMembers || []}
                  principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup]}
                  resolveDelay={1000}
                />
              </div>
              
              {showSpEventsOnly && showOutlookEvents && (
                <div className={styles.formGroup}>
                  <label>Save to:</label>
                  <div className={styles.saveToggle}>
                    <button 
                      type="button"
                      className={!saveToOutlook ? styles.activeToggle : ''}
                      onClick={() => setSaveToOutlook(false)}
                    >
                      SharePoint
                    </button>
                    <button 
                      type="button"
                      className={saveToOutlook ? styles.activeToggle : ''}
                      onClick={() => setSaveToOutlook(true)}
                    >
                      Outlook
                    </button>
                  </div>
                </div>
              )}
              
              <div className={styles.modalButtons}>
                <button className={styles.btnOutline} type="button" onClick={() => setIsEventModalOpen(false)}>
                  Cancel
                </button>
                <button className={styles.btnPrimary} type="submit" disabled={isSaving}>
                  {isSaving ? 'Saving...' : 'Save'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};

export default CollaborationCalendar;