import OpenAI from "openai";
import { redactForLlm } from "./security";
import {
  SupportCategory,
  SUPPORT_CATEGORIES,
  classifyOffline,
  isValidCategory,
  RESPONSE_LIBRARY,
} from "./responses";

// HARD-LOCK DESIGN: the AI returns classification + extraction ONLY. There is no
// user-facing "reply" field — the model has no channel through which to write
// troubleshooting advice. All user-facing text comes from responses.ts.

export interface SupportClassification {
  category: SupportCategory;
  affectedSystem: string; // display label; may be more specific than the category
  impact: "1" | "2" | "3"; // ServiceNow scale: 1 high .. 3 low
  urgency: "1" | "2" | "3";
  escalateImmediately: boolean; // skip self-help (security, account, severe impact)
  shortDescription: string; // one-line summary for the ticket
  startedAt?: string;
  errorMessage?: string;
  snCategory?: string;
  snSubcategory?: string;
}

// Construct the client lazily. Building it at module load throws when there is no
// API key, which would crash the bot in its supported offline (rules-only) mode.
let cachedClient: OpenAI | null = null;
const getClient = (): OpenAI => {
  if (!cachedClient) {
    cachedClient = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
  }
  return cachedClient;
};

// Build a classification purely from the deterministic offline router. Used when
// there is no API key, or as the fallback when the API call fails.
export const classifyFromRules = (
  originalMessage: string
): SupportClassification => {
  const category = classifyOffline(originalMessage.toLowerCase());
  const entry = RESPONSE_LIBRARY[category];
  return {
    category,
    affectedSystem: entry.label,
    impact: "3",
    urgency: "3",
    escalateImmediately: category === "security" || category === "account",
    shortDescription: originalMessage.slice(0, 120) || "IT support request",
    snCategory: entry.snCategory,
    snSubcategory: entry.snSubcategory,
  };
};

const asImpactUrgency = (value: unknown): "1" | "2" | "3" =>
  value === "1" || value === "2" || value === "3" ? value : "3";

const asOptionalString = (value: unknown): string | undefined =>
  typeof value === "string" && value.trim() ? value.trim() : undefined;

const normalize = (
  value: Record<string, unknown>,
  originalMessage: string
): SupportClassification => {
  const rulesFallback = classifyFromRules(originalMessage);
  const category: SupportCategory = isValidCategory(value.category)
    ? value.category
    : rulesFallback.category;
  const entry = RESPONSE_LIBRARY[category];

  // Security and account ALWAYS escalate, regardless of what the model returned.
  const escalateImmediately =
    category === "security" ||
    category === "account" ||
    value.escalateImmediately === true;

  return {
    category,
    affectedSystem:
      asOptionalString(value.affectedSystem) || entry.label,
    impact: asImpactUrgency(value.impact),
    urgency: asImpactUrgency(value.urgency),
    escalateImmediately,
    shortDescription:
      asOptionalString(value.shortDescription) ||
      rulesFallback.shortDescription,
    startedAt: asOptionalString(value.startedAt),
    errorMessage: asOptionalString(value.errorMessage),
    snCategory: entry.snCategory,
    snSubcategory: entry.snSubcategory,
  };
};

export const classifySupportMessage = async (
  userMessage: string
): Promise<SupportClassification> => {
  if (!process.env.OPENAI_API_KEY) {
    return classifyFromRules(userMessage);
  }

  // Scrub credentials/secrets/PII before the message ever leaves the bot.
  const { text: safeMessage } = redactForLlm(userMessage);

  try {
    const completion = await getClient().chat.completions.create({
      model: "gpt-4o-mini",
      response_format: { type: "json_object" },
      messages: [
        {
          role: "system",
          content: [
            "You are the classifier for ReBath IT Helper. Your ONLY job is to classify an IT support message and extract fields.",
            "You DO NOT write advice, troubleshooting steps, or any message to the user. Return JSON only.",
            `Classify the message into exactly one category from this list: ${SUPPORT_CATEGORIES.join(", ")}.`,
            "Use 'account' for password, MFA, authenticator, or sign-in problems. Use 'security' for phishing, malware, viruses, suspicious messages, or possible compromise. Use 'general' if nothing else fits.",
            "Set escalateImmediately true when the issue is security-related, account/sign-in related, affects multiple users or a whole store, involves possible data loss, or completely blocks the person from working.",
            "impact and urgency use ServiceNow's scale where 1 is high, 2 is medium, 3 is low.",
            "Extract affectedSystem (a short label like Outlook, Teams, VPN, Printer, Network), startedAt (when it began, if stated), and errorMessage (verbatim error text, if any). Use empty string when not stated.",
            "The message may contain [REDACTED ...] placeholders where sensitive data was removed. Treat them as opaque; never ask for the redacted value.",
            "Return only valid JSON with this exact shape:",
            JSON.stringify({
              category: "one of the listed categories",
              affectedSystem: "string",
              impact: "1 | 2 | 3",
              urgency: "1 | 2 | 3",
              escalateImmediately: false,
              shortDescription: "string",
              startedAt: "string",
              errorMessage: "string",
            }),
          ].join("\n"),
        },
        { role: "user", content: safeMessage },
      ],
      temperature: 0,
    });

    const content = completion.choices[0]?.message?.content;
    if (!content) return classifyFromRules(userMessage);

    const parsed = JSON.parse(content) as Record<string, unknown>;
    return normalize(parsed, userMessage);
  } catch (error) {
    console.error("AI classification failed:", error);
    return classifyFromRules(userMessage);
  }
};
