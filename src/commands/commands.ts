type RecipientDetails = {
  displayName: string;
  emailAddress: string;
  recipientType: 'to' | 'cc' | 'bcc';
};

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

function getDomain(email: string | undefined): string {
  if (!email) {
    return '';
  }
  const domain = email.split('@')[1];
  return domain ? domain.toLowerCase() : '';
}

function analyzeRecipients(recipients: RecipientDetails[]) {
  const domainCounts = new Map<string, number>();
  recipients.forEach((recipient) => {
    const domain = getDomain(recipient.emailAddress);
    if (domain) {
      domainCounts.set(domain, (domainCounts.get(domain) || 0) + 1);
    }
  });

  const sortedDomains = Array.from(domainCounts.entries()).sort((a, b) => b[1] - a[1]);
  const primaryDomain = sortedDomains[0]?.[0] ?? '';

  const outliers = primaryDomain
    ? recipients.filter((recipient) => {
        const domain = getDomain(recipient.emailAddress);
        return domain && domain !== primaryDomain;
      })
    : [];

  const namesToDomains = new Map<string, Set<string>>();
  recipients.forEach((recipient) => {
    const name = (recipient.displayName || recipient.emailAddress || '').trim().toLowerCase();
    const domain = getDomain(recipient.emailAddress);
    if (!name || !domain) {
      return;
    }
    if (!namesToDomains.has(name)) {
      namesToDomains.set(name, new Set());
    }
    namesToDomains.get(name)!.add(domain);
  });

  const similarNameDifferentDomain = Array.from(namesToDomains.entries())
    .filter(([, domains]) => domains.size > 1)
    .map(([name, domains]) => ({ name, domains: Array.from(domains).sort() }));

  return {
    domainCounts,
    primaryDomain,
    outliers,
    similarNameDifferentDomain
  };
}

function formatSummary(
  recipients: RecipientDetails[],
  analysis: ReturnType<typeof analyzeRecipients>
): string {
  const { primaryDomain, outliers, similarNameDifferentDomain, domainCounts } = analysis;
  const sections: string[] = [];

  sections.push(`Total recipients: ${recipients.length}`);

  if (domainCounts.size > 0) {
    const dominant = primaryDomain ? `${primaryDomain} (${domainCounts.get(primaryDomain)})` : 'varied';
    sections.push(`Primary domain: ${dominant}`);
  }

  if (outliers.length > 0) {
    const details = outliers
      .map((recipient) => `${recipient.displayName} (${getDomain(recipient.emailAddress)})`)
      .join('; ');
    sections.push(`Domain outliers (${outliers.length}): ${details}`);
  } else {
    sections.push('No domain outliers detected.');
  }

  if (similarNameDifferentDomain.length > 0) {
    const details = similarNameDifferentDomain
      .map(({ name, domains }) => `${name} => ${domains.join(', ')}`)
      .join('; ');
    sections.push(`Same name on multiple domains: ${details}`);
  } else {
    sections.push('No same-name cross-domain recipients detected.');
  }

  return sections.join(' | ');
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
