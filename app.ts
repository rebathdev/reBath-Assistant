import { stripMentionsText, TokenCredentials } from "@microsoft/teams.api";
import { App } from "@microsoft/teams.apps";
import { LocalStorage } from "@microsoft/teams.common";
import { ClientSecretCredential } from "@azure/identity";
import { classifySupportMessage, SupportClassification } from "./aiHelper";
import { createIncident, isServiceNowConfigured } from "./serviceNow";
import {
  SupportCategory,
  RESPONSE_LIBRARY,
  buildSelfHelpMessage,
  buildEscalationMessage,
} from "./responses";
import {
  redactForStorage,
  safeLogValue,
  isAllowedTenant,
  isTenantGuardConfigured,
} from "./security";

const storage = new LocalStorage();

const createTokenFactory = () => {
  return async (scope: string | string[], tenantId?: string): Promise<string> => {
    const clientId = process.env.CLIENT_ID || process.env.BOT_ID || "";
    const clientSecret = process.env.CLIENT_SECRET || "";
    const resolvedTenantId = tenantId || process.env.TENANT_ID || "";

    if (!clientId || !clientSecret || !resolvedTenantId) {
      throw new Error(
        "Missing required bot auth environment variables. Required: CLIENT_ID or BOT_ID, CLIENT_SECRET, and TENANT_ID."
      );
    }

    const credential = new ClientSecretCredential(
      resolvedTenantId,
      clientId,
      clientSecret
    );

    const scopes = Array.isArray(scope) ? scope : [scope];
    const tokenResponse = await credential.getToken(scopes);

    if (!tokenResponse?.token) {
      throw new Error("Failed to acquire token for Teams bot authentication.");
    }

    return tokenResponse.token;
  };
};

const tokenCredentials: TokenCredentials = {
  clientId: process.env.CLIENT_ID || process.env.BOT_ID || "",
  token: createTokenFactory(),
};

const app = new App({ ...tokenCredentials, storage });

// Startup posture warnings.
if (!isTenantGuardConfigured()) {
  console.warn(
    "[startup] ALLOWED_TENANT_ID is not set. The bot will respond to any tenant. Set it for production."
  );
}
if (!isServiceNowConfigured()) {
  console.warn(
    "[startup] ServiceNow is not configured. Tickets will be saved locally but NOT created in ServiceNow."
  );
}

interface TicketDetails {
  requesterName?: string;
  requesterTeamsId?: string;
  requesterAadObjectId?: string;
  requesterEmail?: string;
  affectedSystem?: string;
  issueSummary?: string;
  startedAt?: string;
  device?: string;
  errorMessage?: string;
  businessImpact?: string;
  othersAffected?: string;
  troubleshootingTried?: string;
  category?: string;
  subcategory?: string;
  impact?: "1" | "2" | "3";
  urgency?: "1" | "2" | "3";
}

interface ConversationState {
  count: number;
  mode?: "idle" | "ticket" | "awaitingTroubleshootingResult";
  ticketStep?: number;
  ticket?: TicketDetails;
  lastTicketNumber?: string;
}

const getConversationState = (conversationId: string): ConversationState => {
  let state = storage.get(conversationId) as ConversationState | undefined;
  if (!state) {
    state = { count: 0, mode: "idle" };
    storage.set(conversationId, state);
  }
  return state;
};

const saveConversationState = (id: string, state: ConversationState): void => {
  storage.set(id, state);
};

const clearConversationState = (id: string): void => {
  storage.delete(id);
};

const getTeamsUserInfo = (activity: any) => ({
  requesterName: activity.from?.name || "there",
  requesterTeamsId: activity.from?.id || "Unknown Teams ID",
  requesterAadObjectId: activity.from?.aadObjectId || "Unknown AAD object ID",
  requesterEmail:
    activity.from?.userPrincipalName || activity.from?.email || undefined,
});

const isGreeting = (t: string): boolean =>
  ["hi", "hello", "hey", "yo", "good morning", "good afternoon", "good evening"].includes(t);

const isGeneralHelpRequest = (t: string): boolean =>
  ["help", "need help", "i need help", "can you help", "can you help me", "support", "it help"].includes(t);

const isTicketStartRequest = (t: string): boolean =>
  ["/ticket", "ticket", "new ticket", "create ticket", "open ticket", "start ticket", "submit ticket", "support ticket"].includes(t);

const isCancelRequest = (t: string): boolean =>
  ["/cancel", "cancel", "stop", "nevermind", "never mind"].includes(t);

const isConfirmRequest = (t: string): boolean =>
  ["confirm", "submit", "yes", "y", "create", "create ticket"].includes(t);

const stillNotWorking = (t: string): boolean =>
  ["still not working", "still broken", "still frozen", "still slow", "did not work", "didn't work", "didnt work", "not fixed", "same issue", "still happening", "issue still exists", "yes still broken", "no it did not work", "no it didn't work", "no"].includes(t);

const issueResolved = (t: string): boolean =>
  ["fixed", "it worked", "working now", "resolved", "all good", "yes it worked", "yes"].includes(t);

const isRebootHowTo = (t: string): boolean =>
  ["how do i reboot", "how to reboot", "how do i restart", "how to restart", "how do i restart my computer"].some((p) => t.includes(p));

const rebootGuidance = (): string =>
  [
    "**Safe reboot guidance:**",
    "",
    "A reboot can clear common app, VPN, printer, and Windows session issues.",
    "",
    "1. Save your work.",
    "2. Close open apps.",
    "3. Restart from the Windows Start menu.",
    "4. Sign back in and test the issue again.",
    "",
    "If the issue continues, reply **still not working** and I’ll help open a ticket for IT.",
  ].join("\n");

const generalHelpMessage = (): string =>
  [
    "Hi — I can help with common IT issues first, then open a ticket for IT if it still doesn’t work.",
    "",
    "Just tell me what’s happening, for example:",
    "",
    "- **Outlook is frozen**",
    "- **Teams won’t open**",
    "- **VPN won’t connect**",
    "- **A website isn’t loading**",
    "- **My computer is slow**",
    "",
    "I’ll suggest safe first steps. If it still fails, I’ll collect the ticket details automatically.",
    "",
    "Please don’t paste passwords, MFA codes, or other secrets — I’ll automatically remove anything that looks sensitive.",
  ].join("\n");

const helpMessage = (): string =>
  [
    "**ReBath IT Helper**",
    "",
    "I can suggest safe first steps for common IT issues and collect details for an IT ticket.",
    "",
    "**Commands**",
    "",
    "**/ticket** - Start a new IT support ticket",
    "**/status** - Show current ticket progress",
    "**/reboot** - Show safe reboot guidance",
    "**/cancel** - Cancel the current ticket intake",
    "**/reset** - Clear the current conversation",
    "**/help** - Show this help message",
    "",
    "You can also just describe the problem, like **Outlook is frozen** or **my computer is slow**.",
  ].join("\n");

const ticketStatusMessage = (state: ConversationState): string => {
  if (state.mode !== "ticket" || !state.ticketStep) {
    return [
      "No ticket is currently in progress.",
      "",
      "Type **/ticket** to start a new IT support ticket.",
      state.lastTicketNumber ? `Last ticket: **${state.lastTicketNumber}**` : "",
    ]
      .filter(Boolean)
      .join("\n");
  }
  const t = state.ticket || {};
  return [
    "**Current ticket progress**",
    "",
    `**Requester:** ${t.requesterName || "Not provided yet"}`,
    `**Affected system:** ${t.affectedSystem || "Not provided yet"}`,
    `**Issue:** ${t.issueSummary || "Not provided yet"}`,
    `**Started:** ${t.startedAt || "Not provided yet"}`,
    `**Device:** ${t.device || "Not provided yet"}`,
    `**Error message:** ${t.errorMessage || "Not provided yet"}`,
    `**Business impact:** ${t.businessImpact || "Not provided yet"}`,
    "",
    "Type **/cancel** to cancel this ticket intake.",
  ].join("\n");
};

const buildTicketSummary = (state: ConversationState): string => {
  const t = state.ticket || {};
  return [
    "**IT Support Ticket Summary**",
    "",
    `**Requester:** ${t.requesterName || "Not provided"}`,
    `**Affected system:** ${t.affectedSystem || "Not provided"}`,
    `**Issue:** ${t.issueSummary || "Not provided"}`,
    `**Started:** ${t.startedAt || "Not provided"}`,
    `**Device:** ${t.device || "Not provided"}`,
    `**Error message:** ${t.errorMessage || "Not provided"}`,
    `**Business impact:** ${t.businessImpact || "Not provided"}`,
    `**Troubleshooting tried:** ${t.troubleshootingTried || "Not provided"}`,
    "",
    isServiceNowConfigured()
      ? "Type **confirm** to create this ticket in ServiceNow, or **/cancel** to cancel."
      : "ServiceNow isn’t configured, so this will be saved but not sent. Type **confirm** to finish, or **/cancel** to cancel.",
  ].join("\n");
};

const buildServiceNowDescription = (t: TicketDetails): string =>
  [
    `Requester: ${t.requesterName || "Unknown"}`,
    t.requesterEmail ? `Email: ${t.requesterEmail}` : "",
    `Affected system: ${t.affectedSystem || "Not provided"}`,
    `Issue: ${t.issueSummary || "Not provided"}`,
    `Started: ${t.startedAt || "Not provided"}`,
    `Device: ${t.device || "Not provided"}`,
    `Error message: ${t.errorMessage || "Not provided"}`,
    `Business impact: ${t.businessImpact || "Not provided"}`,
    `Others affected: ${t.othersAffected || "Not provided"}`,
    `Troubleshooting tried: ${t.troubleshootingTried || "Not provided"}`,
    "",
    "Submitted via ReBath IT Helper (Teams).",
  ]
    .filter(Boolean)
    .join("\n");

// Collapsed intake: ask only for fields still missing, in priority order.
const NEXT_PROMPTS: Array<{ field: keyof TicketDetails; ask: string }> = [
  { field: "affectedSystem", ask: "Which app or system is affected? Example: Outlook, Teams, VPN, OneDrive, printer, internet." },
  { field: "issueSummary", ask: 'What exactly is happening? Example: "Outlook will not open" or "VPN will not connect".' },
  { field: "startedAt", ask: "When did this start? Example: today, this morning, after a reboot, or after a recent update." },
  { field: "device", ask: "What device are you on? Example: Windows laptop, desktop, shared store PC, iPhone." },
  { field: "businessImpact", ask: "How is this affecting your work? Example: blocked completely, slowing me down, one task affected, or store operations impacted." },
];

const advanceTicketIntake = async (
  context: any,
  conversationId: string,
  state: ConversationState
): Promise<void> => {
  const ticket = state.ticket || {};
  for (let i = 0; i < NEXT_PROMPTS.length; i++) {
    const { field, ask } = NEXT_PROMPTS[i];
    const current = ticket[field];
    if (!current || (typeof current === "string" && !current.trim())) {
      state.ticketStep = 100 + i;
      saveConversationState(conversationId, state);
      await context.send(ask);
      return;
    }
  }
  state.ticketStep = 200;
  saveConversationState(conversationId, state);
  await context.send(buildTicketSummary(state));
};

const captureFieldAnswer = (state: ConversationState, scrubbed: string): void => {
  const idx = (state.ticketStep || 100) - 100;
  const prompt = NEXT_PROMPTS[idx];
  if (!prompt) return;
  state.ticket = state.ticket || {};
  (state.ticket as any)[prompt.field] = scrubbed;
};

// Seed ticket state from a classification so intake can skip known fields.
const seedTicketFromClassification = (
  state: ConversationState,
  activity: any,
  scrubbedIssue: string,
  cls: SupportClassification,
  troubleshootingTried?: string
): void => {
  state.mode = "ticket";
  state.ticket = {
    ...(state.ticket || {}),
    ...getTeamsUserInfo(activity),
    issueSummary: state.ticket?.issueSummary || scrubbedIssue,
    affectedSystem: cls.affectedSystem,
    startedAt: cls.startedAt || state.ticket?.startedAt,
    errorMessage: cls.errorMessage || state.ticket?.errorMessage,
    businessImpact: cls.escalateImmediately
      ? state.ticket?.businessImpact || "Flagged as high-impact during triage."
      : state.ticket?.businessImpact,
    category: cls.snCategory,
    subcategory: cls.snSubcategory,
    impact: cls.impact,
    urgency: cls.urgency,
    troubleshootingTried:
      troubleshootingTried || state.ticket?.troubleshootingTried,
  };
};

const submitTicket = async (
  context: any,
  conversationId: string,
  state: ConversationState
): Promise<void> => {
  const ticket = state.ticket || {};

  if (!isServiceNowConfigured()) {
    const localRef = `LOCAL-${Date.now().toString().slice(-6)}`;
    state.mode = "idle";
    state.ticketStep = undefined;
    state.lastTicketNumber = localRef;
    saveConversationState(conversationId, state);
    await context.send(
      [
        `Saved your request as **${localRef}** (local only).`,
        "",
        "ServiceNow isn’t connected yet, so this wasn’t sent to IT. Once an admin sets the ServiceNow credentials, confirmed tickets will create real incidents automatically.",
      ].join("\n")
    );
    return;
  }

  await context.send("Creating your ticket in ServiceNow…");

  const result = await createIncident({
    shortDescription:
      ticket.issueSummary || `${ticket.affectedSystem || "IT"} issue reported via Teams`,
    description: buildServiceNowDescription(ticket),
    impact: ticket.impact || "3",
    urgency: ticket.urgency || "3",
    category: ticket.category,
    subcategory: ticket.subcategory,
    callerEmail: ticket.requesterEmail,
    callerName: ticket.requesterName,
    contactType: "chat",
  });

  if (result.ok && result.incidentNumber) {
    state.mode = "idle";
    state.ticketStep = undefined;
    state.lastTicketNumber = result.incidentNumber;
    state.ticket = undefined;
    saveConversationState(conversationId, state);
    await context.send(
      [
        `Done — created **${result.incidentNumber}** in ServiceNow.`,
        "",
        "IT has the details. You can reference that number in any follow-up. If the problem changes or gets worse, message me again.",
      ].join("\n")
    );
    return;
  }

  console.error("ServiceNow incident creation failed:", result.error);
  saveConversationState(conversationId, state);
  await context.send(
    [
      "I couldn’t create the ticket in ServiceNow just now.",
      "",
      "Your details are saved. Please reply **confirm** to try again, or contact the IT service desk directly if it’s urgent.",
    ].join("\n")
  );
};

// Central handler for a freshly described issue. This is the ONLY place a new issue
// is interpreted, and every user-facing word it produces comes from responses.ts.
const handleNewIssue = async (
  context: any,
  conversationId: string,
  state: ConversationState,
  rawText: string
): Promise<void> => {
  const { text: scrubbed } = redactForStorage(rawText);
  const cls = await classifySupportMessage(scrubbed);
  const user = getTeamsUserInfo(context.activity);
  const entry = RESPONSE_LIBRARY[cls.category];

  // Escalate-immediately categories (security, account) or severe impact: skip
  // self-help and go straight into ticket intake, pre-seeded from the classification.
  if (cls.escalateImmediately || !entry.selfHelp) {
    seedTicketFromClassification(
      state,
      context.activity,
      scrubbed,
      cls,
      "Escalated at triage without self-help (security/account/high-impact)."
    );
    saveConversationState(conversationId, state);
    await context.send(buildEscalationMessage(cls.category, user.requesterName));
    await advanceTicketIntake(context, conversationId, state);
    return;
  }

  // Otherwise: offer the approved self-help script for this category, and remember
  // the classification so we can pre-fill the ticket if it doesn't work.
  state.mode = "awaitingTroubleshootingResult";
  state.ticketStep = undefined;
  state.ticket = {
    ...user,
    affectedSystem: cls.affectedSystem,
    issueSummary: scrubbed,
    startedAt: cls.startedAt,
    errorMessage: cls.errorMessage,
    category: cls.snCategory,
    subcategory: cls.snSubcategory,
    impact: cls.impact,
    urgency: cls.urgency,
    troubleshootingTried: `Bot recommended approved ${entry.label} self-help steps.`,
  };
  saveConversationState(conversationId, state);
  await context.send(
    buildSelfHelpMessage(cls.category, user.requesterName, cls.errorMessage)
  );
};

app.on("message", async (context) => {
  const activity = context.activity;
  const conversationId = activity.conversation.id;
  const text: string = stripMentionsText(activity).trim();
  const lowerText = text.toLowerCase();

  const activityTenantId =
    activity.conversation?.tenantId ||
    (activity as any).channelData?.tenant?.id ||
    undefined;
  if (!isAllowedTenant(activityTenantId)) {
    await context.send(
      "Sorry, this assistant is only available to users within the organization."
    );
    return;
  }

  console.log("Received Teams message:", {
    conversationId,
    conversationType: activity.conversation.conversationType,
    fromAadObjectId: activity.from?.aadObjectId,
    textPreview: safeLogValue(text),
  });

  if (!text) {
    await context.send(
      "ReBath IT Helper is online. Tell me what issue you’re having, or type **/help** for commands."
    );
    return;
  }

  if (lowerText === "/reset") {
    clearConversationState(conversationId);
    await context.send(
      "Ok, I cleared the current conversation. Tell me the issue, or type **/ticket** to start a new IT support ticket."
    );
    return;
  }

  const state = getConversationState(conversationId);

  if (isCancelRequest(lowerText)) {
    state.mode = "idle";
    state.ticketStep = undefined;
    state.ticket = undefined;
    saveConversationState(conversationId, state);
    await context.send(
      "Canceled the current ticket intake. Tell me the issue if you still need help."
    );
    return;
  }

  if (lowerText === "/help") {
    await context.send(helpMessage());
    return;
  }

  if (lowerText === "/reboot" || isRebootHowTo(lowerText)) {
    await context.send(rebootGuidance());
    return;
  }

  if (lowerText === "/status") {
    await context.send(ticketStatusMessage(state));
    return;
  }

  if (isGreeting(lowerText) || isGeneralHelpRequest(lowerText)) {
    await context.send(generalHelpMessage());
    return;
  }

  if (isTicketStartRequest(lowerText)) {
    state.mode = "ticket";
    state.ticketStep = 1;
    state.ticket = { ...getTeamsUserInfo(activity) };
    saveConversationState(conversationId, state);
    await context.send(
      'Ok, let’s open a ticket for IT. In one message, tell me what’s wrong — which app or system, what’s happening, and any error you see. Example: "Outlook won’t open on my work laptop, error 0x800."'
    );
    return;
  }

  // Did the self-help work?
  if (state.mode === "awaitingTroubleshootingResult") {
    if (issueResolved(lowerText)) {
      state.mode = "idle";
      state.ticketStep = undefined;
      state.ticket = undefined;
      saveConversationState(conversationId, state);
      await context.send(
        "Great — glad that fixed it. If the issue comes back, message me again and I can help open a ticket for IT."
      );
      return;
    }
    if (stillNotWorking(lowerText)) {
      state.mode = "ticket";
      state.ticket = {
        ...(state.ticket || {}),
        ...getTeamsUserInfo(activity),
        troubleshootingTried:
          state.ticket?.troubleshootingTried ||
          "Bot recommended approved self-help steps. User reported the issue still exists.",
      };
      saveConversationState(conversationId, state);
      await advanceTicketIntake(context, conversationId, state);
      return;
    }
    await context.send(
      "Did that fix the issue? Reply **fixed** if it’s working now, or **still not working** if it continues."
    );
    return;
  }

  // Ticket intake: free-text first message (step 1) -> classify + pre-fill.
  if (state.mode === "ticket" && state.ticketStep === 1) {
    const { text: scrubbed } = redactForStorage(text);
    const cls = await classifySupportMessage(scrubbed);
    seedTicketFromClassification(state, activity, scrubbed, cls);
    saveConversationState(conversationId, state);
    await advanceTicketIntake(context, conversationId, state);
    return;
  }

  // Ticket intake: answering an individual field (steps 100-199).
  if (
    state.mode === "ticket" &&
    typeof state.ticketStep === "number" &&
    state.ticketStep >= 100 &&
    state.ticketStep < 200
  ) {
    const { text: scrubbed } = redactForStorage(text);
    captureFieldAnswer(state, scrubbed);
    saveConversationState(conversationId, state);
    await advanceTicketIntake(context, conversationId, state);
    return;
  }

  // Ticket intake: awaiting confirm (step 200).
  if (state.mode === "ticket" && state.ticketStep === 200) {
    if (isConfirmRequest(lowerText)) {
      await submitTicket(context, conversationId, state);
      return;
    }
    await context.send(
      "Please type **confirm** to submit this ticket, or **/cancel** to cancel."
    );
    return;
  }

  // Anything else is treated as a new issue description. Classification picks the
  // approved response; the AI never writes the user-facing text.
  await handleNewIssue(context, conversationId, state, text);
});

export default app;
