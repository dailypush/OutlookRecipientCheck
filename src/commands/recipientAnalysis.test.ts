import { describe, expect, it } from 'vitest';
import { analyzeRecipients, formatSummary, getDomain, RecipientDetails } from './recipientAnalysis';

describe('getDomain', () => {
  it('returns normalized domain when present', () => {
    expect(getDomain('User@Example.COM')).toBe('example.com');
  });

  it('returns empty string for malformed input', () => {
    expect(getDomain('invalid-email')).toBe('');
    expect(getDomain(undefined)).toBe('');
  });
});

describe('analyzeRecipients', () => {
  const recipients: RecipientDetails[] = [
    { displayName: 'Alice', emailAddress: 'alice@alpha.com', recipientType: 'to' },
    { displayName: 'Bob', emailAddress: 'bob@alpha.com', recipientType: 'cc' },
    { displayName: 'Carol', emailAddress: 'carol@beta.com', recipientType: 'bcc' },
    { displayName: 'Alex Doe', emailAddress: 'alex@alpha.com', recipientType: 'to' },
    { displayName: 'Alex Doe', emailAddress: 'alex@other.org', recipientType: 'cc' }
  ];

  it('identifies primary domain and outliers', () => {
    const analysis = analyzeRecipients(recipients);
    expect(analysis.primaryDomain).toBe('alpha.com');
    expect(analysis.outliers.map((r) => r.emailAddress)).toContain('carol@beta.com');
  });

  it('detects same name on multiple domains', () => {
    const analysis = analyzeRecipients(recipients);
    expect(analysis.similarNameDifferentDomain).toEqual([
      { name: 'alex doe', domains: ['alpha.com', 'other.org'] }
    ]);
  });
});

describe('formatSummary', () => {
  it('summarizes analysis into a readable string', () => {
    const recipients: RecipientDetails[] = [
      { displayName: 'Alice', emailAddress: 'alice@alpha.com', recipientType: 'to' },
      { displayName: 'Bob', emailAddress: 'bob@alpha.com', recipientType: 'to' },
      { displayName: 'Carol', emailAddress: 'carol@beta.com', recipientType: 'cc' }
    ];
    const analysis = analyzeRecipients(recipients);
    const summary = formatSummary(recipients, analysis);

    expect(summary).toContain('Total recipients: 3');
    expect(summary).toContain('Primary domain: alpha.com (2)');
    expect(summary).toContain('Domain outliers (1): Carol (beta.com)');
    expect(summary).toContain('Same name on multiple domains');
  });
});
