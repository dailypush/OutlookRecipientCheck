export type RecipientDetails = {
  displayName: string;
  emailAddress: string;
  recipientType: 'to' | 'cc' | 'bcc';
};

export type AnalysisOptions = {
  /** Domains that should never be treated as outliers. */
  trustedDomains?: string[];
  /** Default domain to treat as the "safe" domain, e.g., the sender's domain. */
  defaultPrimaryDomain?: string;
  /** Optional logger for debugging and diagnostics. */
  logger?: import('./logger').Logger;
};

export function getDomain(email: string | undefined): string {
  if (!email) {
    return '';
  }
  const domain = email.split('@')[1];
  return domain ? domain.toLowerCase() : '';
}

function normalizeDomains(domains: string[] | undefined): Set<string> {
  return new Set((domains ?? []).map((domain) => domain.trim().toLowerCase()).filter(Boolean));
}

export function analyzeRecipients(recipients: RecipientDetails[], options: AnalysisOptions = {}) {
  const trustedDomainSet = normalizeDomains(options.trustedDomains);
  const defaultPrimaryDomain = getDomain(options.defaultPrimaryDomain);
  const logger = options.logger;

  const domainCounts = new Map<string, number>();
  recipients.forEach((recipient) => {
    const domain = getDomain(recipient.emailAddress);
    if (domain) {
      domainCounts.set(domain, (domainCounts.get(domain) || 0) + 1);
    }
  });

  logger?.debug('Domain counts computed', domainCounts);

  const sortedDomains = Array.from(domainCounts.entries()).sort((a, b) => b[1] - a[1]);
  const primaryDomain = defaultPrimaryDomain || sortedDomains[0]?.[0] || '';

  logger?.debug('Primary domain derived', { primaryDomain, sortedDomains, defaultPrimaryDomain });

  const outliers = primaryDomain
    ? recipients.filter((recipient) => {
        const domain = getDomain(recipient.emailAddress);
        return domain && domain !== primaryDomain && !trustedDomainSet.has(domain);
      })
    : [];

  logger?.debug('Outlier recipients detected', outliers);

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

  if (similarNameDifferentDomain.length > 0) {
    logger?.debug('Same name detected on multiple domains', similarNameDifferentDomain);
  }

  return {
    domainCounts,
    primaryDomain,
    outliers,
    trustedDomains: Array.from(trustedDomainSet).sort(),
    defaultPrimaryDomain,
    similarNameDifferentDomain
  };
}

export function formatSummary(
  recipients: RecipientDetails[],
  analysis: ReturnType<typeof analyzeRecipients>
): string {
  const {
    primaryDomain,
    outliers,
    similarNameDifferentDomain,
    domainCounts,
    trustedDomains,
    defaultPrimaryDomain
  } = analysis;
  const sections: string[] = [];

  sections.push(`Total recipients: ${recipients.length}`);

  if (domainCounts.size > 0 || primaryDomain) {
    const primaryLabel = primaryDomain
      ? `${primaryDomain} (${domainCounts.get(primaryDomain) ?? 0})`
      : 'varied';
    sections.push(`Primary domain: ${primaryLabel}`);
  }

  sections.push(
    trustedDomains.length > 0
      ? `Trusted domains ignored: ${trustedDomains.join(', ')}`
      : 'Trusted domains ignored: none'
  );

  if (defaultPrimaryDomain) {
    sections.push(`Safe domain baseline: ${defaultPrimaryDomain}`);
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
