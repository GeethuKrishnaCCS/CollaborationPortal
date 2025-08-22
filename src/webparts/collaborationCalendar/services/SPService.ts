import { getSP } from '../pnpjsConfig';
import { IEventItem } from '../models/IEventItem';
import { WebPartContext } from '@microsoft/sp-webpart-base';

const ensureListExists = async (listName: string): Promise<void> => {
  const sp = getSP();
  try {
    await sp.web.lists.getByTitle(listName).select('Title')();
  } catch (err) {
    if (err?.status === 404) {
      throw new Error(`List '${listName}' not found`);
    }
    throw err;
  }
};


const checkColumnExists = async (listName: string, columnName: string): Promise<boolean> => {
  const sp = getSP();
  try {
    await sp.web.lists.getByTitle(listName).fields.getByInternalNameOrTitle(columnName).select('InternalName')();
    return true;
  } catch (err) {
    if (err?.status === 404) {
      console.warn(`Column '${columnName}' does not exist in list '${listName}'.`);
      return false;
    }
    throw err;
  }
};


export const getEventsFromList = async (listName: string): Promise<IEventItem[]> => {
  try {
    await ensureListExists(listName);
    const sp = getSP();
    
    
    const hasTeamMembersColumn = await checkColumnExists(listName, 'TeamMembers');
    
    const selectFields = ["Id", "Title", "EventDate", "EndDate", "Location", "Description"];
    if (hasTeamMembersColumn) {
      selectFields.push("TeamMembers");
    }

    const items = await sp.web.lists.getByTitle(listName).items
      .select(...selectFields)
      .top(100)
      .orderBy('EventDate', true)();

    return items.map(item => ({
      Id: item.Id,
      Title: item.Title || 'No Title',
      EventDate: new Date(item.EventDate),
      EndDate: new Date(item.EndDate),
      Location: item.Location || '',
      Description: item.Description || '',
      TeamMembers: hasTeamMembersColumn && item.TeamMembers ? item.TeamMembers.split(',').map((name: string) => name.trim()) : [],
      Source: 'SharePoint',
    }));
  } catch (err) {
    console.error("Error getting SharePoint events:", err);
    return [];
  }
};


export const getOutlookEvents = async (context: WebPartContext): Promise<IEventItem[]> => {
  try {
    if (!context.msGraphClientFactory) {
      throw new Error("MS Graph Client Factory not available");
    }

    const client = await context.msGraphClientFactory.getClient('3');
    const response = await client
      .api("/me/calendar/events")
      .select("subject,start,end,location,bodyPreview,attendees")
      .top(100)
      .orderby("start/dateTime")
      .get();

    return response.value.map((item: any, index: number): IEventItem => {
      const startTime = item.start?.dateTime ? new Date(item.start.dateTime) : new Date();
      const endTime = item.end?.dateTime ? new Date(item.end.dateTime) : new Date();
      
      return {
        Id: index, 
        Title: item.subject || 'No Title',
        EventDate: startTime,
        EndDate: endTime,
        Location: item.location?.displayName || '',
        Description: item.bodyPreview || '',
        Source: 'Outlook',
        TeamMembers: item.attendees?.map((a: any) => a.emailAddress?.name) || [],
      };
    });
  } catch (error) {
    console.error("Error getting Outlook events:", error);
    return [];
  }
};


export const resolveUserEmails = async (context: WebPartContext, displayNames: string[]): Promise<string[]> => {
  try {
    const client = await context.msGraphClientFactory.getClient('3');
    const emails: string[] = [];

    for (const displayName of displayNames) {
      try {
        const response = await client
          .api('/users')
          .filter(`displayName eq '${displayName}'`)
          .select('mail')
          .top(1)
          .get();

        if (response.value && response.value.length > 0 && response.value[0].mail) {
          emails.push(response.value[0].mail);
        } else {
          console.warn(`Could not resolve email for display name: ${displayName}`);
        }
      } catch (error) {
        console.error(`Error resolving email for ${displayName}:`, error);
      }
    }

    return emails;
  } catch (error) {
    console.error("Error resolving user emails:", error);
    return [];
  }
};

const sendEmailNotifications = async (
  context: WebPartContext,
  recipients: string[],
  subject: string,
  body: string
): Promise<void> => {
  try {
    const client = await context.msGraphClientFactory.getClient('3');
    const emailMessage = {
      message: {
        subject: subject,
        body: {
          contentType: 'Text',
          content: body,
        },
        toRecipients: recipients.map(email => ({
          emailAddress: {
            address: email,
          },
        })),
      },
      saveToSentItems: true,
    };

    await client.api('/me/sendMail').post(emailMessage);
    console.log(`Email sent successfully to ${recipients.join(', ')}`);
  } catch (error) {
    console.error('Error sending email notification:', error);
    throw error;
  }
};

export const saveEventToList = async (
  listName: string,
  eventData: {
    Title: string;
    EventDate: Date;
    EndDate: Date;
    Location?: string;
    Description?: string;
    TeamMembers?: string[];
  },
  context: WebPartContext
): Promise<number> => {
  try {
    await ensureListExists(listName);
    const sp = getSP();


    const hasTeamMembersColumn = await checkColumnExists(listName, 'TeamMembers');

    const formattedEvent: { [key: string]: string } = {
      Title: eventData.Title,
      EventDate: eventData.EventDate.toISOString(),
      EndDate: eventData.EndDate.toISOString(),
      Location: eventData.Location || '',
      Description: eventData.Description || '',
    };

    if (hasTeamMembersColumn && eventData.TeamMembers) {
      formattedEvent.TeamMembers = eventData.TeamMembers.join(', ');
    }

    console.log("Attempting to save event:", formattedEvent);

   
    const result = await sp.web.lists
      .getByTitle(listName)
      .items.add(formattedEvent);

   
    if (eventData.TeamMembers && eventData.TeamMembers.length > 0) {
      const recipientEmails = await resolveUserEmails(context, eventData.TeamMembers);
      if (recipientEmails.length > 0) {
        const emailSubject = `New Event: ${eventData.Title}`;
        const emailBody = `You have been invited to an event.\n\n` +
          `Title: ${eventData.Title}\n` +
          `Start: ${eventData.EventDate.toLocaleString()}\n` +
          `End: ${eventData.EndDate.toLocaleString()}\n` +
          `Location: ${eventData.Location || 'Not specified'}\n` +
          `Description: ${eventData.Description || 'Not specified'}\n` +
          `Team Members: ${eventData.TeamMembers.join(', ')}`;

        await sendEmailNotifications(context, recipientEmails, emailSubject, emailBody);
      }
    }

    if (result?.data?.Id) {
      console.log("Successfully saved event with ID (from result):", result.data.Id);
      return result.data.Id;
    }

   
    if (result?.item) {
      const item = await result.item.select("Id")();
      if (item?.Id) {
        console.log("Successfully saved event with ID (from item refetch):", item.Id);
        return item.Id;
      }
    }

 
    console.warn("ID not found in result or item fetch, attempting to query list...");
    const recentItems = await sp.web.lists
      .getByTitle(listName)
      .items.select("Id", "Title", "EventDate")
      .filter(`Title eq '${formattedEvent.Title}' and EventDate eq '${formattedEvent.EventDate}'`)
      .top(1)();

    if (recentItems?.length > 0 && recentItems[0]?.Id) {
      console.log("Successfully retrieved event ID from list query:", recentItems[0].Id);
      return recentItems[0].Id;
    }

    throw new Error("Failed to retrieve the ID of the created item.");

  } catch (err) {
    console.error("Error saving event to SharePoint:", err);
    throw err;
  }
};

export const updateEventInList = async (
  listName: string,
  eventData: {
    Id: number;
    Title: string;
    EventDate: Date;
    EndDate: Date;
    Location?: string;
    Description?: string;
    TeamMembers?: string[];
  },
  context: WebPartContext
): Promise<void> => {
  try {
    await ensureListExists(listName);
    const sp = getSP();
    
    
    const hasTeamMembersColumn = await checkColumnExists(listName, 'TeamMembers');

    const updateData: { [key: string]: string } = {
      Title: eventData.Title,
      EventDate: eventData.EventDate.toISOString(),
      EndDate: eventData.EndDate.toISOString(),
      Location: eventData.Location || '',
      Description: eventData.Description || '',
    };

    if (hasTeamMembersColumn && eventData.TeamMembers) {
      updateData.TeamMembers = eventData.TeamMembers.join(', ');
    }

    console.log("Attempting to update event:", updateData);

    await sp.web.lists.getByTitle(listName).items.getById(eventData.Id).update(updateData);

    
    if (eventData.TeamMembers && eventData.TeamMembers.length > 0) {
      const recipientEmails = await resolveUserEmails(context, eventData.TeamMembers);
      if (recipientEmails.length > 0) {
        const emailSubject = `Updated Event: ${eventData.Title}`;
        const emailBody = `An event you are invited to has been updated.\n\n` +
          `Title: ${eventData.Title}\n` +
          `Start: ${eventData.EventDate.toLocaleString()}\n` +
          `End: ${eventData.EndDate.toLocaleString()}\n` +
          `Location: ${eventData.Location || 'Not specified'}\n` +
          `Description: ${eventData.Description || 'Not specified'}\n` +
          `Team Members: ${eventData.TeamMembers.join(', ')}`;

        await sendEmailNotifications(context, recipientEmails, emailSubject, emailBody);
      }
    }

    console.log("Successfully updated event with ID:", eventData.Id); 
  } catch (err) {
    console.error("Error updating event in SharePoint:", err);
    throw err;
  }
};

export const deleteEventFromList = async (listName: string, eventId: number): Promise<void> => {
  try {
    await ensureListExists(listName);
    const sp = getSP();
    
    console.log("Attempting to delete event with ID:", eventId);
    
    await sp.web.lists.getByTitle(listName).items.getById(eventId).delete();
    
    console.log("Successfully deleted event with ID:", eventId);
  } catch (err) {
    console.error("Error deleting event from SharePoint:", err);
    throw err;
  }
};



export const saveEventToOutlook = async (
  context: WebPartContext,
  eventData: {
    Title: string;
    EventDate: Date;
    EndDate: Date;
    Location?: string;
    Description?: string;
    TeamMembers?: string[];
  }
): Promise<void> => {
  try {
    const client = await context.msGraphClientFactory.getClient("3");

    
    const attendeesEmails = eventData.TeamMembers && eventData.TeamMembers.length > 0
      ? await resolveUserEmails(context, eventData.TeamMembers)
      : [];

    if (eventData.TeamMembers && eventData.TeamMembers.length > 0 && attendeesEmails.length === 0) {
      throw new Error("Failed to resolve email addresses for attendees.");
    }

    const newEvent = {
      subject: eventData.Title,
      start: {
        dateTime: eventData.EventDate.toISOString(),
        timeZone: "Asia/Kolkata" // Enforce IST time zone
      },
      end: {
        dateTime: eventData.EndDate.toISOString(),
        timeZone: "Asia/Kolkata"
      },
      location: {
        displayName: eventData.Location || ""
      },
      body: {
        contentType: "Text",
        content: eventData.Description || ""
      },
      attendees: attendeesEmails.map(email => ({
        emailAddress: { address: email, name: email },
        type: 'required'
      }))
    };

    await client.api("/me/events").post(newEvent);
    console.log("Event successfully saved to Outlook calendar.");
  } catch (err) {
    console.error("Error saving event to Outlook:", err);
    throw err; 
  }
};

export const getAllEvents = async (
  listName: string,
  context: WebPartContext
): Promise<IEventItem[]> => {
  try {
    const [spEvents, outlookEvents] = await Promise.all([
      getEventsFromList(listName),
      getOutlookEvents(context)
    ]);

    return [...spEvents, ...outlookEvents].sort(
      (a, b) => a.EventDate.getTime() - b.EventDate.getTime()
    );
  } catch (err) {
    console.error("Error fetching all events:", err);
    return [];
  }
};