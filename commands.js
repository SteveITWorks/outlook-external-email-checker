// commands.js - Event handler implementation
Office.onReady(() => {
  console.log('Office Add-in initialized');
});

// Configuration - Set your organization's domain(s)
const INTERNAL_DOMAINS = [
  'itworks.co.nz',
  'subsidiary.com'
];

/**
 * Checks if an email address belongs to an external user
 * @param {string} email - Email address to check
 * @returns {boolean} - True if external, false if internal
 */
function isExternalEmail(email) {
  if (!email) return false;
  
  const emailLower = email.toLowerCase();
  const domain = emailLower.split('@')[1];
  
  if (!domain) return false;
  
  return !INTERNAL_DOMAINS.some(internalDomain => 
    domain === internalDomain.toLowerCase()
  );
}

/**
 * Extracts email addresses from recipient objects
 * @param {Array} recipients - Array of recipient objects
 * @returns {Array} - Array of email addresses
 */
function extractEmails(recipients) {
  if (!recipients || !Array.isArray(recipients)) return [];
  return recipients.map(recipient => recipient.emailAddress).filter(email => email);
}

/**
 * Main event handler for OnMessageSend
 * @param {object} event - Event object from Outlook
 */
function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;
  
  // Get all recipients
  item.to.getAsync((toResult) => {
    if (toResult.status === Office.AsyncResultStatus.Failed) {
      console.error('Failed to get TO recipients:', toResult.error);
      event.completed({ allowEvent: true });
      return;
    }
    
    item.cc.getAsync((ccResult) => {
      if (ccResult.status === Office.AsyncResultStatus.Failed) {
        console.error('Failed to get CC recipients:', ccResult.error);
        event.completed({ allowEvent: true });
        return;
      }
      
      item.bcc.getAsync((bccResult) => {
        if (bccResult.status === Office.AsyncResultStatus.Failed) {
          console.error('Failed to get BCC recipients:', bccResult.error);
          event.completed({ allowEvent: true });
          return;
        }
        
        // Combine all recipients
        const allRecipients = [
          ...extractEmails(toResult.value),
          ...extractEmails(ccResult.value),
          ...extractEmails(bccResult.value)
        ];
        
        // Check for external recipients
        const externalRecipients = allRecipients.filter(isExternalEmail);
        
        if (externalRecipients.length > 0) {
          // Block send and show warning
          const externalList = externalRecipients.join(', ');
          const message = `Warning: You are sending this email to ${externalRecipients.length} external recipient(s):\n\n${externalList}\n\nAre you sure you want to continue?`;
          
          event.completed({
            allowEvent: false,
            errorMessage: message,
            errorMessageMarkdown: `**Warning: External Recipients Detected**\n\nYou are sending this email to ${externalRecipients.length} external recipient(s):\n\n${externalRecipients.map(e => `- ${e}`).join('\n')}\n\nPlease review and click **Send Anyway** if you want to proceed.`
          });
        } else {
          // All recipients are internal, allow send
          event.completed({ allowEvent: true });
        }
      });
    });
  });
}

// Register the function
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);