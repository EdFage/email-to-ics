// Configuration and globals
const TARGET_EMAIL = 'makecalendarevent@gmail.com';
const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');

// Main function that triggers when email arrives
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
        console.log('Processing email:');
        console.log('From:', message.getFrom());
        console.log('Subject:', message.getSubject());
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

async function processEmail(message) {
  const emailContent = {
    subject: message.getSubject(),
    body: message.getPlainBody(),
    from: message.getFrom()
  };
  
  // Add 'await' here to get the actual data
  const eventDetails = await extractEventDetails(emailContent);
  const icsContent = createICSFile(eventDetails);
  
  // Pass all three required arguments: recipient, icsContent, and eventDetails
  sendResponse(emailContent.from, icsContent, eventDetails);
}

async function extractEventDetails(emailContent) {
  const url = 'https://api.openai.com/v1/chat/completions';
  
  // Get current date in format "24th Dec 2024"
  const now = new Date().toLocaleDateString('en-US', {
    day: 'numeric',
    month: 'short',
    year: 'numeric'
  }).replace(',', '');
  
  const requestBody = {
    model: "gpt-3.5-turbo",
    messages: [
      {
        role: "system",
        content: "You are an assistant that takes an email and creates calendar event JSON. Return dates in ICS format (YYYYMMDDTHHMMSSZ). Only return valid JSON with no markdown formatting or backticks."
      },
      {
        role: "user", 
        content: `Parse this email into a JSON object with event_title, datetime_start, datetime_end, and location. The current time is ${now}

Email Subject: ${emailContent.subject}
Email Body: ${emailContent.body}`
      }
    ],
    response_format: { type: "json_object" },
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
    console.log('Raw OpenAI response:', eventJson);
    
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
    console.error('Error in extractEventDetails:', error);
    throw error;
  }
}

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

function sendResponse(recipient, icsContent, eventDetails) {
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

    // Make sure ICS content ends with proper line endings
    const formattedICS = icsContent.trim() + '\r\n';
    
    // Create blob with proper MIME type and encoding
    const icsBlob = Utilities.newBlob('')
      .setDataFromString(formattedICS, 'UTF-8')
      .setContentType('text/calendar; charset=UTF-8; method=REQUEST')
      .setName('invite.ics');
    
    // Log the attachment details for debugging
    console.log('ICS Blob details:', {
      contentType: icsBlob.getContentType(),
      size: icsBlob.getBytes().length,
      content: formattedICS
    });
    
    // Create and send the email with attachment
    GmailApp.sendEmail(recipient, 
      'Your Calendar Invite', 
      emailBody,
      {
        attachments: [icsBlob],
        from: TARGET_EMAIL,
        name: 'Calendar Assistant'
      }
    );

    console.log('Response email sent successfully to:', recipient);
    
  } catch (error) {
    console.error('Error sending response:', error);
    throw error;
  }
}

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