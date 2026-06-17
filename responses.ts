// responses.ts
// The single source of truth for everything a user can see during self-help.
//
// HARD-LOCK DESIGN: the AI never writes user-facing troubleshooting text. The AI's
// only job is to classify a message into one of the categories below and extract a
// few fields. The actual words shown to the user always come from this file. To
// change what users see, edit THIS file — not the AI prompt.
//
// Every step here must be safe, first-line, and non-destructive: close/reopen,
// sign out/in, reboot once, clear cache/cookies, try another browser, check the
// connection, capture an error. No scripts, no admin commands, no registry edits.

export type SupportCategory =
  | "outlook"
  | "teams"
  | "vpn"
  | "printer"
  | "network"
  | "onedrive"
  | "sharepoint"
  | "browser"
  | "computer"
  | "account"
  | "security"
  | "general";

export const SUPPORT_CATEGORIES: SupportCategory[] = [
  "outlook",
  "teams",
  "vpn",
  "printer",
  "network",
  "onedrive",
  "sharepoint",
  "browser",
  "computer",
  "account",
  "security",
  "general",
];

export interface ResponseEntry {
  label: string; // Friendly display name, e.g. "Microsoft Teams"
  selfHelp: boolean; // false => skip troubleshooting, go straight to a ticket
  steps: string[]; // Approved safe steps (only used when selfHelp is true)
  escalateMessage?: string; // Shown instead of steps when selfHelp is false
  // ServiceNow classification mapping for this category.
  snCategory?: string;
  snSubcategory?: string;
}

export const RESPONSE_LIBRARY: Record<SupportCategory, ResponseEntry> = {
  outlook: {
    label: "Outlook",
    selfHelp: true,
    steps: [
      "1. Fully close Outlook and reopen it.",
      "2. Check that you’re connected to the internet (try loading any website).",
      "3. Save your work, reboot the computer once, and sign back in.",
      "4. If you see an error message, copy it or take a screenshot.",
    ],
    snCategory: "software",
    snSubcategory: "email",
  },
  teams: {
    label: "Microsoft Teams",
    selfHelp: true,
    steps: [
      "1. Fully quit Teams (right-click the Teams icon in the system tray and choose Quit), then reopen it.",
      "2. Sign out of Teams and sign back in.",
      "3. Save your work, reboot the computer once, and sign back in.",
      "4. If you see an error message, copy it or take a screenshot.",
    ],
    snCategory: "software",
    snSubcategory: "collaboration",
  },
  vpn: {
    label: "VPN",
    selfHelp: true,
    steps: [
      "1. Disconnect from the VPN, wait a few seconds, and reconnect.",
      "2. Check that you’re connected to the internet first (try loading any website).",
      "3. Reboot the computer once and try connecting again.",
      "4. If you see an error message or code, copy it or take a screenshot.",
    ],
    snCategory: "network",
    snSubcategory: "vpn",
  },
  printer: {
    label: "Printer",
    selfHelp: true,
    steps: [
      "1. Check the printer is powered on and any cables or Wi-Fi connection are in place.",
      "2. Turn the printer off, wait 10 seconds, and turn it back on.",
      "3. Make sure the correct printer is selected when you print.",
      "4. Reboot the computer once and try a test print.",
    ],
    snCategory: "hardware",
    snSubcategory: "printer",
  },
  network: {
    label: "Network / Wi-Fi",
    selfHelp: true,
    steps: [
      "1. Check Wi-Fi is turned on and you’re connected to the right network.",
      "2. Turn Wi-Fi off, wait a few seconds, and turn it back on.",
      "3. If you’re at home, restart your router if you can.",
      "4. Reboot the computer once and test again.",
    ],
    snCategory: "network",
    snSubcategory: "connectivity",
  },
  onedrive: {
    label: "OneDrive",
    selfHelp: true,
    steps: [
      "1. Close OneDrive and reopen it.",
      "2. Check that you’re connected to the internet.",
      "3. Reboot the computer once and let OneDrive finish syncing.",
      "4. If a file shows a sync error, take a screenshot of it.",
    ],
    snCategory: "software",
    snSubcategory: "file_sync",
  },
  sharepoint: {
    label: "SharePoint",
    selfHelp: true,
    steps: [
      "1. Open the site in a private/InPrivate browser window.",
      "2. Clear your browser’s cache and cookies for the site.",
      "3. Try a different browser if one is available.",
      "4. Reboot the computer once and try again.",
    ],
    snCategory: "software",
    snSubcategory: "intranet",
  },
  browser: {
    label: "Browser / Website",
    selfHelp: true,
    steps: [
      "1. Close and reopen the browser.",
      "2. Open the site in a private/InPrivate window.",
      "3. Clear the browser’s cache and cookies for the site.",
      "4. Try a different browser if one is available, then reboot once if it still fails.",
    ],
    snCategory: "software",
    snSubcategory: "browser",
  },
  computer: {
    label: "Computer",
    selfHelp: true,
    steps: [
      "1. Save your work and close any apps you’re not using.",
      "2. Reboot the computer once and sign back in.",
      "3. Give it a couple of minutes to finish starting up before testing again.",
      "4. If you see an error message, copy it or take a screenshot.",
    ],
    snCategory: "hardware",
    snSubcategory: "endpoint",
  },
  account: {
    label: "Account / Sign-in",
    selfHelp: false,
    steps: [],
    escalateMessage:
      "For password, MFA, authenticator, or sign-in problems I won’t guess at steps — and please don’t type your password or any codes into this chat. Let’s get this to IT directly.",
    snCategory: "account",
    snSubcategory: "access",
  },
  security: {
    label: "Security",
    selfHelp: false,
    steps: [],
    escalateMessage:
      "This looks security-related, so I’m getting IT involved right away. Please don’t click anything further, don’t delete anything, and don’t forward suspicious content unless IT asks — they may need it as evidence.",
    snCategory: "security",
    snSubcategory: "incident",
  },
  general: {
    label: "IT issue",
    selfHelp: true,
    steps: [
      "1. Fully close the affected app and reopen it.",
      "2. Check that you’re connected to the internet.",
      "3. Save your work, reboot the computer once, and sign back in.",
      "4. If you see an error message, copy it or take a screenshot.",
    ],
    snCategory: "inquiry",
    snSubcategory: "general",
  },
};

// Templated acknowledgment + steps. Slots are filled with extracted values, but the
// sentence structure is always ours — the AI never composes this.
export const buildSelfHelpMessage = (
  category: SupportCategory,
  requesterName: string,
  errorMessage?: string
): string => {
  const entry = RESPONSE_LIBRARY[category];
  const lines: string[] = [`Thanks, ${requesterName}.`, ""];

  lines.push(`That sounds like a problem with **${entry.label}**.`);
  if (errorMessage) {
    lines.push("", `I’ve noted the error you mentioned: “${errorMessage}”.`);
  }
  lines.push("", "Please try these safe first steps:", "", ...entry.steps);
  lines.push(
    "",
    "After trying that, reply **fixed** if it’s working now, or **still not working** and I’ll open a ticket for IT."
  );

  return lines.join("\n");
};

// Message shown for categories that skip self-help and go straight to a ticket.
export const buildEscalationMessage = (
  category: SupportCategory,
  requesterName: string
): string => {
  const entry = RESPONSE_LIBRARY[category];
  return [
    `Thanks, ${requesterName}.`,
    "",
    entry.escalateMessage || "Let’s get this to IT.",
    "",
    "I’ll start a ticket now and ask only for what I still need.",
  ].join("\n");
};

// --- Deterministic offline classifier -----------------------------------------
// Used when OpenAI is unavailable or returns nothing usable. It only ROUTES to a
// category (picking which approved script to show) — it never generates text. This
// keeps the safety guarantee intact even with no AI: security and account issues
// still escalate, everything else still gets an approved script.
export const classifyOffline = (lowerText: string): SupportCategory => {
  const has = (...words: string[]) => words.some((w) => lowerText.includes(w));

  if (has("hacked", "compromised", "phishing", "suspicious email", "suspicious link", "bad link", "malware", "virus", "ransomware", "scam", "clicked a link", "weird email", "fraud"))
    return "security";
  if (has("password", "mfa", "authenticator", "2fa", "locked out", "can't sign in", "cant sign in", "can't log in", "cant log in"))
    return "account";
  if (has("outlook", "email", "mailbox")) return "outlook";
  if (has("teams")) return "teams";
  if (has("vpn")) return "vpn";
  if (has("printer", "print", "printing")) return "printer";
  if (has("wifi", "wi-fi", "internet", "network", "connection")) return "network";
  if (has("onedrive")) return "onedrive";
  if (has("sharepoint")) return "sharepoint";
  if (has("website", "browser", "chrome", "edge", "page", "site", "url", "cookies", "cache"))
    return "browser";
  if (has("computer", "laptop", "pc", "desktop", "slow", "frozen", "freezing", "crash"))
    return "computer";
  return "general";
};

export const isValidCategory = (value: unknown): value is SupportCategory => {
  return typeof value === "string" && (SUPPORT_CATEGORIES as string[]).includes(value);
};
