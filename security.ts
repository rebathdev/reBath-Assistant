// security.ts
// Lightweight, dependency-free redaction + access guards.
//
// Why this exists: a help desk bot is the single most likely place for a user to
// paste a password, an MFA code, an API key, or a customer's personal data. Any of
// that flowing into OpenAI, into ServiceNow, or into your console logs is a real
// data-handling problem. These helpers scrub free text before it leaves the bot.
//
// This is defense-in-depth, not a guarantee. Regexes catch the common shapes, not
// every possible secret. Keep your OpenAI data-retention settings and ServiceNow
// access controls configured correctly regardless.

export interface RedactionResult {
  text: string;
  redacted: boolean;
  categories: string[];
}

interface Rule {
  label: string;
  pattern: RegExp;
  replacement: string;
}

// Order matters: more specific patterns first so they win before generic ones.
const RULES: Rule[] = [
  // Lines where the user explicitly types a credential, e.g. "password: hunter2",
  // "pwd is abc123", "passcode = 9981". Capture the value after the separator.
  {
    label: "credential",
    pattern:
      /\b(pass(?:word|wd|code)?|pwd|secret|api[\s_-]?key|token|client[\s_-]?secret)\b\s*(?:is|:|=|->)?\s*\S+/gi,
    replacement: "[REDACTED CREDENTIAL]",
  },
  // Standalone 6-8 digit codes that look like MFA/OTP (avoids matching years/ticket
  // numbers by requiring an explicit MFA/OTP/code/verification keyword nearby).
  {
    label: "mfa_code",
    pattern: /\b(?:mfa|otp|one[\s-]?time|verification|auth(?:enticator)?)\b[^\d]{0,12}\d{4,8}\b/gi,
    replacement: "[REDACTED MFA CODE]",
  },
  // Credit-card-shaped numbers (13-16 digits, optional spaces/dashes).
  {
    label: "credit_card",
    pattern: /\b(?:\d[ -]?){13,16}\b/g,
    replacement: "[REDACTED CARD NUMBER]",
  },
  // US SSN shape.
  {
    label: "ssn",
    pattern: /\b\d{3}-\d{2}-\d{4}\b/g,
    replacement: "[REDACTED SSN]",
  },
  // Bearer tokens / long base64-ish secrets (32+ chars of token alphabet).
  {
    label: "bearer_token",
    pattern: /\b(?:bearer\s+)?[A-Za-z0-9_\-]{32,}\.?[A-Za-z0-9_\-.]*\b/gi,
    replacement: "[REDACTED TOKEN]",
  },
  // Email addresses — usually fine to keep for support, but redact when scrubbing
  // for the LLM call so customer PII is not sent off to a third party unnecessarily.
  // This rule is opt-in via redactForLlm (see below).
];

const EMAIL_RULE: Rule = {
  label: "email",
  pattern: /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b/g,
  replacement: "[REDACTED EMAIL]",
};

const applyRules = (input: string, rules: Rule[]): RedactionResult => {
  let text = input;
  const categories = new Set<string>();

  for (const rule of rules) {
    if (rule.pattern.test(text)) {
      categories.add(rule.label);
    }
    // Reset lastIndex because we reuse global regexes.
    rule.pattern.lastIndex = 0;
    text = text.replace(rule.pattern, rule.replacement);
    rule.pattern.lastIndex = 0;
  }

  return {
    text,
    redacted: categories.size > 0,
    categories: Array.from(categories),
  };
};

// Use before storing in a ticket, sending to ServiceNow, or writing to logs.
// Keeps emails (support often needs them) but strips credentials/secrets/PII numbers.
export const redactForStorage = (input: string): RedactionResult => {
  return applyRules(input, RULES);
};

// Use before sending to OpenAI. Strips everything redactForStorage does, plus emails,
// so customer/employee email addresses are not transmitted to the LLM provider.
export const redactForLlm = (input: string): RedactionResult => {
  return applyRules(input, [...RULES, EMAIL_RULE]);
};

// Safe logging helper: redacts, then truncates. Never log raw user text.
export const safeLogValue = (input: string, maxLen = 200): string => {
  const { text } = redactForStorage(input);
  return text.length > maxLen ? `${text.slice(0, maxLen)}…` : text;
};

// --- Access guards -------------------------------------------------------------

// Restrict the bot to your tenant. Set ALLOWED_TENANT_ID in the environment
// (your Microsoft Entra tenant GUID). If unset, the guard is disabled and a
// warning should be logged at startup — fine for local/playground, not for prod.
export const isAllowedTenant = (activityTenantId?: string): boolean => {
  const allowed = process.env.ALLOWED_TENANT_ID?.trim();
  if (!allowed) {
    return true; // not configured -> allow (dev convenience)
  }
  return Boolean(activityTenantId) && activityTenantId === allowed;
};

export const isTenantGuardConfigured = (): boolean => {
  return Boolean(process.env.ALLOWED_TENANT_ID?.trim());
};
