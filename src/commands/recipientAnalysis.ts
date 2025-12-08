export type RecipientDetails = {
  displayName: string;
  emailAddress: string;
  recipientType: 'to' | 'cc' | 'bcc';
};

export function getDomain(email: string | undefined): string {
  if (!email) {
    return '';
  }
  const domain = email.split('@')[1];
  return domain ? domain.toLowerCase() : '';
}

export function analyzeRecipients(recipients: RecipientDetails[]) {
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

export function formatSummary(
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

  const collisionSection = similarNameDifferentDomain.length > 0
    ? `Same name on multiple domains: ${similarNameDifferentDomain
        .map(({ name, domains }) => `${name} => ${domains.join(', ')}`)
        .join('; ')}`
    : 'Same name on multiple domains: none detected.';

  sections.push(collisionSection);

  return sections.join(' | ');
}
