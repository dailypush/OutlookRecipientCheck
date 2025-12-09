import {
  analyzeRecipients,
  AnalysisOptions,
  formatSummary,
  getDomain,
  RecipientDetails
} from './recipientAnalysis';
import { createLogger, isDebugEnabled, setDebugEnabled } from './logger';

const notificationKey = 'recipient-check';
const trustedDomains: string[] = [];
const logger = createLogger('commands');

Office.onReady(() => {
  Office.actions.associate('checkRecipients', checkRecipients);
});

function getUserDomain(): string {
  const email = Office.context.mailbox?.userProfile?.emailAddress;
  return getDomain(email);
}

function refreshDebugFlag() {
  try {
    const stored = typeof localStorage !== 'undefined' ? localStorage.getItem('recipientCheckDebug') : null;
    setDebugEnabled(stored === 'true');
  } catch (error) {
    console.warn('[RecipientCheck] Unable to read debug flag', error);
  }
}

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
  logger.debug('Collected recipients by field', { to: toRecipients.length, cc: ccRecipients.length, bcc: bccRecipients.length });
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
  refreshDebugFlag();
  logger.info('Starting recipient check', { debugEnabled: isDebugEnabled() });
  try {
    const recipients = await collectRecipients();
    const analysisOptions: AnalysisOptions = {
      trustedDomains,
      defaultPrimaryDomain: getUserDomain(),
      logger
    };
    const analysis = analyzeRecipients(recipients, analysisOptions);
    logger.debug('Analysis complete', analysis);
    const summary = formatSummary(recipients, analysis);
    logger.info('Summary prepared for notification');
    await showNotification(summary);
  } catch (error: unknown) {
    const fallback = error instanceof Error ? error.message : 'Unknown error during recipient check.';
    logger.error('Recipient check failed', error);
    await showNotification(`Recipient check failed: ${fallback}`);
  } finally {
    logger.debug('Recipient check completed, signaling Office host.');
    event.completed();
  }
}

export { checkRecipients };
