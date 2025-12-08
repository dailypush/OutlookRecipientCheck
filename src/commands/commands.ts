import { analyzeRecipients, formatSummary, getDomain, RecipientDetails } from './recipientAnalysis';

const notificationKey = 'recipient-check';

Office.onReady(() => {
  Office.actions.associate('checkRecipients', checkRecipients);
});

function getRecipients(field: 'to' | 'cc' | 'bcc'): Promise<RecipientDetails[]> {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item as Office.MessageCompose | undefined;
    if (!item) {
      reject(new Error('Mailbox item unavailable.'));
      return;
    }

    const callback = (asyncResult: Office.AsyncResult<Office.EmailAddressDetails[]>) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        const recipients = (asyncResult.value || []).map((recipient) => ({
          displayName: recipient.displayName || recipient.emailAddress,
          emailAddress: recipient.emailAddress,
          recipientType: field
        }));
        resolve(recipients);
      } else {
        reject(asyncResult.error);
      }
    };

    switch (field) {
      case 'to':
        item.to.getAsync(callback);
        break;
      case 'cc':
        item.cc.getAsync(callback);
        break;
      case 'bcc':
        item.bcc.getAsync(callback);
        break;
      default:
        resolve([]);
    }
  });
}

async function collectRecipients(): Promise<RecipientDetails[]> {
  const [toRecipients, ccRecipients, bccRecipients] = await Promise.all([
    getRecipients('to'),
    getRecipients('cc'),
    getRecipients('bcc')
  ]);
  return [...toRecipients, ...ccRecipients, ...bccRecipients];
}

function showNotification(message: string): Promise<void> {
  return new Promise((resolve) => {
    const item = Office.context.mailbox.item;
    if (!item || !item.notificationMessages) {
      resolve();
      return;
    }

    const clippedMessage = message.length > 150 ? `${message.slice(0, 147)}...` : message;
    item.notificationMessages.replaceAsync(
      notificationKey,
      {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: clippedMessage,
        icon: 'Icon.16',
        persistent: false
      },
      () => resolve()
    );
  });
}

async function checkRecipients(event: Office.AddinCommands.Event) {
  try {
    const recipients = await collectRecipients();
    const analysis = analyzeRecipients(recipients);
    const summary = formatSummary(recipients, analysis);
    await showNotification(summary);
  } catch (error: unknown) {
    const fallback = error instanceof Error ? error.message : 'Unknown error during recipient check.';
    await showNotification(`Recipient check failed: ${fallback}`);
  } finally {
    event.completed();
  }
}

export { checkRecipients };
