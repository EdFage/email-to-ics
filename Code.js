// Configuration and globals
const TARGET_EMAIL = 'makecalendarevent@gmail.com';
const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');

/**
 * Main function triggered when a new email arrives.
 * Searches for unread emails addressed to the target email and processes them.
 *
 * @async
 * @param {object} e - The event object provided by the Gmail trigger.
 * @returns {Promise<void>} - Resolves after all unread emails are processed.
 * @throws {Error} Throws an error if there is an issue processing emails.
 */
async function onNewEmail(e) {
  try {
    console.log('Starting email check...');
    const threads = GmailApp.search('is:unread to:' + TARGET_EMAIL);
    console.log(`Found ${threads.length} unread threads`);
    
    for (const thread of threads) {
      const messages = thread.getMessages();
      console.log(`Processing thread with ${messages.length} messages`);
      
      for (const message of messages) {
        // Log email details
        console.log('Subject:', message.getSubject());
        console.log('From:', message.getFrom());
        console.log('Body:', message.getPlainBody().substring(0, 100) + '...'); // First 100 chars
        
        await processEmail(message);
        message.markRead();
        console.log('Email marked as read');
      }
    }
  } catch (error) {
    console.error('Error in onNewEmail:', error);
  }
}

/**
 * Processes an individual email to extract event details
 * and send a calendar invite response.
 *
 * @async
 * @param {object} message - The Gmail message object.
 * @returns {Promise<void>} - Resolves when the email has been processed and responded to.
 * @throws {Error} Throws an error if email processing or response fails.
 */
async function processEmail(message) {
    try {
      // Concatenate the email content into a plain-text input
      const text = `Subject: ${message.getSubject()}\nBody: ${message.getPlainBody()}`;
      console.log('Processing email text:', text);
  
      // Call OpenAI API to get event details
      const eventData = await openAiApiCall(text);
  
      // Create ICS content using the extracted event details
      const icsContent = createICSFile(eventData);
  
      // Send the response email with the ICS file attached
      sendResponse(message, icsContent, eventData);
  
      console.log('Email processed successfully.');
    } catch (error) {
      console.error('Error in processEmail:', error);
      throw error;
    }
  }


/**
 * Sends text to the OpenAI API and parses the response into event details.
 *
 * @param {string} text - The plain text to be sent to the API.
 * @returns {Promise<object>} - Resolves with an object containing event details (event_title, datetime_start, datetime_end, location).
 * @throws {Error} Throws an error if the API call fails or returns invalid data.
 */
async function openAiApiCall(text) {
    const url = 'https://api.openai.com/v1/chat/completions';
    
    const requestBody = {
      model: "gpt-3.5-turbo",
      messages: [
        {
          role: "system",
          content: `You are an assistant that extracts event details from an email and returns them as valid JSON for a calendar invite. 
          The JSON must include the following fields:
          - "event_title" (string, required)
          - "datetime_start" (string, required, in the format YYYYMMDDTHHMMSS)
          - "datetime_end" (string, required, in the format YYYYMMDDTHHMMSS)
          - "location" (string, optional)
          
          Your response should always be strictly valid JSON. Do not include any extra explanations, markdown, or formatting. If a required field cannot be determined, return an empty string for that field.`
              },
        {
          role: "user",
          content: text
        }
      ],
      temperature: 0.1
    };
  
    const options = {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${OPENAI_API_KEY}`,
        'Content-Type': 'application/json',
      },
      payload: JSON.stringify(requestBody)
    };
  
    try {
      console.log('Sending prompt to OpenAI:', requestBody.messages);
      const response = UrlFetchApp.fetch(url, options);
      const json = JSON.parse(response.getContentText());
      const eventJson = json.choices[0].message.content;
  
      // Parse the response string into JSON
      const eventData = JSON.parse(eventJson);
      console.log('Parsed event data:', eventData);
  
      // Validate required fields
      const requiredFields = ["event_title", "datetime_start", "datetime_end"];
      for (const field of requiredFields) {
        if (!eventData[field]) {
          throw new Error(`Missing required field: ${field}`);
        }
      }
  
      return eventData;
    } catch (error) {
      console.error('Error in openAiApiCall:', error);
      throw error;
    }
  }


/**
 * Creates ICS file content for a calendar event from provided event details.
 *
 * @param {object} eventData - An object containing event details.
 * @param {string} eventData.event_title - The event's title.
 * @param {string} eventData.datetime_start - The event's start time in ICS format.
 * @param {string} eventData.datetime_end - The event's end time in ICS format.
 * @param {string} [eventData.location] - The optional event location.
 * @returns {string} - The ICS file content as a string.
 * @throws {Error} Throws an error if required event details are missing.
 */
function createICSFile(eventData) {
  console.log('Creating ICS file with data:', JSON.stringify(eventData, null, 2));
  
  // Simple escape function for text fields
  const escapeText = (text) => {
    if (!text) return '';
    return text.replace(/[\\;,]/g, '\\$&').replace(/\n/g, '\\n');
  };

  // Validate required fields
  if (!eventData.datetime_start || !eventData.datetime_end || !eventData.event_title) {
    console.error('Missing required event data:', eventData);
    throw new Error('Missing required event data');
  }

  const icsContent = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'BEGIN:VEVENT',
    `SUMMARY:${escapeText(eventData.event_title)}`,
    `DTSTART:${eventData.datetime_start}`,
    `DTEND:${eventData.datetime_end}`,
    eventData.location ? `LOCATION:${escapeText(eventData.location)}` : '',
    'END:VEVENT',
    'END:VCALENDAR'
  ].filter(Boolean).join('\r\n');

  console.log('Generated ICS content:', icsContent.replace(/\r\n/g, '\\r\\n'));
  return icsContent;
}


/**
 * Sends a response email with a calendar invite attached.
 *
 * @param {object} originalMessage - The original Gmail message object.
 * @param {string} icsContent - The generated ICS file content.
 * @param {object} eventDetails - The event details object.
 * @param {string} eventDetails.event_title - The event's title.
 * @param {string} eventDetails.datetime_start - The event's start time.
 * @param {string} eventDetails.datetime_end - The event's end time.
 * @param {string} [eventDetails.location] - The event's location, if provided.
 * @returns {void}
 * @throws {Error} Throws an error if the response email fails to send.
 */
function sendResponse(originalMessage, icsContent, eventDetails) {
    try {
      // Create the email body
      const emailBody = `Hello!
    
  I've created a calendar event based on the email you forwarded:
  Event: ${eventDetails.event_title}
  Date: ${eventDetails.datetime_start} to ${eventDetails.datetime_end}
  ${eventDetails.location ? `Location: ${eventDetails.location}` : ''}
    
  I've attached the calendar invite (.ics file) to this email. You can open it to add this event to your calendar.
    
  Best regards,
  Your Calendar Assistant`;
  
      // Ensure ICS content ends with proper line endings
      const formattedICS = icsContent.trim() + '\r\n';
      
      // Create the ICS file as a blob
      const icsBlob = Utilities.newBlob('')
        .setDataFromString(formattedICS, 'UTF-8')
        .setContentType('text/calendar; charset=UTF-8; method=REQUEST')
        .setName('invite.ics');
      
      // Reply to the thread of the original message
      const thread = originalMessage.getThread();
      thread.reply(emailBody, {
        attachments: [icsBlob],
        from: TARGET_EMAIL,
        name: 'Calendar Assistant'
      });
  
      console.log('Reply sent successfully to the thread.');
    } catch (error) {
      console.error('Error sending response:', error);
      throw error;
    }
  }

/**
 * Sets up a time-driven trigger to execute the `onNewEmail` function every minute.
 * Removes existing triggers to avoid duplicates before creating a new one.
 *
 * @returns {void}
 * @throws {Error} Throws an error if trigger creation fails.
 */
function createTrigger() {
  // Delete any existing triggers first to avoid duplicates
  const existingTriggers = ScriptApp.getProjectTriggers();
  existingTriggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Create a new time-driven trigger that runs every minute
  ScriptApp.newTrigger('onNewEmail')
    .timeBased()
    .everyMinutes(1)
    .create();
    
  console.log('Trigger created successfully');
}