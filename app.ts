import { stripMentionsText, TokenCredentials } from "@microsoft/teams.api";
import { App } from "@microsoft/teams.apps";
import { LocalStorage } from "@microsoft/teams.common";
import { ClientSecretCredential } from "@azure/identity";
import { analyzeSupportMessage } from "./aiHelper";

// Create storage for conversation history
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

const app = new App({
  ...tokenCredentials,
  storage,
});

interface TicketDetails {
  requesterName?: string;
  requesterTeamsId?: string;
  requesterAadObjectId?: string;
  affectedSystem?: string;
  issueSummary?: string;
  startedAt?: string;
  location?: string;
  device?: string;
  errorMessage?: string;
  businessImpact?: string;
  othersAffected?: string;
  troubleshootingTried?: string;
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
    state = {
      count: 0,
      mode: "idle",
    };

    storage.set(conversationId, state);
  }

  return state;
};

const saveConversationState = (
  conversationId: string,
  state: ConversationState
): void => {
  storage.set(conversationId, state);
};

const clearConversationState = (conversationId: string): void => {
  storage.delete(conversationId);
};

const getTeamsUserInfo = (activity: any) => {
  return {
    requesterName: activity.from?.name || "Unknown user",
    requesterTeamsId: activity.from?.id || "Unknown Teams ID",
    requesterAadObjectId: activity.from?.aadObjectId || "Unknown AAD object ID",
  };
};

const isGreeting = (lowerText: string): boolean => {
  return [
    "hi",
    "hello",
    "hey",
    "yo",
    "good morning",
    "good afternoon",
    "good evening",
  ].includes(lowerText);
};

const isGeneralHelpRequest = (lowerText: string): boolean => {
  return [
    "help",
    "need help",
    "i need help",
    "can you help",
    "can you help me",
    "support",
    "it help",
  ].includes(lowerText);
};

const isTicketStartRequest = (lowerText: string): boolean => {
  return [
    "/ticket",
    "ticket",
    "new ticket",
    "create ticket",
    "open ticket",
    "start ticket",
    "submit ticket",
    "support ticket",
  ].includes(lowerText);
};

const isCancelRequest = (lowerText: string): boolean => {
  return ["/cancel", "cancel", "stop", "nevermind", "never mind"].includes(
    lowerText
  );
};

const isConfirmRequest = (lowerText: string): boolean => {
  return ["confirm", "submit", "yes", "y", "create", "create ticket"].includes(
    lowerText
  );
};

const stillNotWorking = (lowerText: string): boolean => {
  return [
    "still not working",
    "still broken",
    "still frozen",
    "still slow",
    "did not work",
    "didn't work",
    "didnt work",
    "not fixed",
    "same issue",
    "still happening",
    "issue still exists",
    "yes still broken",
    "no it did not work",
    "no it didn't work",
    "no",
  ].includes(lowerText);
};

const issueResolved = (lowerText: string): boolean => {
  return [
    "fixed",
    "it worked",
    "working now",
    "resolved",
    "all good",
    "yes it worked",
    "yes",
  ].includes(lowerText);
};

const detectAffectedSystem = (lowerText: string): string => {
  if (lowerText.includes("outlook") || lowerText.includes("email")) {
    return "Outlook";
  }

  if (lowerText.includes("teams")) {
    return "Microsoft Teams";
  }

  if (lowerText.includes("vpn")) {
    return "VPN";
  }

  if (lowerText.includes("printer") || lowerText.includes("print")) {
    return "Printer";
  }

  if (
    lowerText.includes("internet") ||
    lowerText.includes("wifi") ||
    lowerText.includes("wi-fi") ||
    lowerText.includes("network")
  ) {
    return "Network";
  }

  if (lowerText.includes("onedrive")) {
    return "OneDrive";
  }

  if (lowerText.includes("sharepoint")) {
    return "SharePoint";
  }

  if (
    lowerText.includes("website") ||
    lowerText.includes("browser") ||
    lowerText.includes("chrome") ||
    lowerText.includes("edge") ||
    lowerText.includes("page") ||
    lowerText.includes("site")
  ) {
    return "Browser / Website";
  }

  if (
    lowerText.includes("computer") ||
    lowerText.includes("laptop") ||
    lowerText.includes("pc") ||
    lowerText.includes("desktop")
  ) {
    return "Computer";
  }

  return "General IT";
};

const looksLikeIssue = (lowerText: string): boolean => {
  const issueWords = [
    "not working",
    "broken",
    "frozen",
    "freezing",
    "slow",
    "crashing",
    "crashed",
    "will not open",
    "won't open",
    "wont open",
    "cant open",
    "can't open",
    "cannot open",
    "error",
    "stuck",
    "down",
    "offline",
    "cannot connect",
    "can't connect",
    "cant connect",
    "wont connect",
    "won't connect",
    "not connecting",
    "keeps closing",
    "keeps freezing",
    "keeps crashing",
    "not loading",
    "page not loading",
    "website not loading",
    "site not loading",
    "blank page",
    "cache",
    "cookies",
  ];

  return issueWords.some((word) => lowerText.includes(word));
};

const rebootGuidance = (): string => {
  return [
    "**Safe reboot guidance:**",
    "",
    "A reboot can clear common app, VPN, printer, and Windows session issues.",
    "",
    "Before rebooting:",
    "1. Save your work.",
    "2. Close open apps.",
    "3. Restart from the Windows Start menu.",
    "4. Sign back in and test the issue again.",
    "",
    "If the issue continues, reply **still not working** and I’ll help open a ticket for IT.",
  ].join("\n");
};

const generalHelpMessage = (): string => {
  return [
    "Hi — I can help with basic IT issues first, then open a ticket for IT if it still does not work.",
    "",
    "Tell me what is happening, for example:",
    "",
    "- **Outlook is frozen**",
    "- **Teams will not open**",
    "- **VPN will not connect**",
    "- **A website is not loading**",
    "- **My computer is slow**",
    "",
    "I’ll suggest safe first steps like restarting the app, rebooting, clearing cache/cookies, or trying an incognito/InPrivate browser. If it still fails, I’ll collect the ticket details automatically.",
  ].join("\n");
};

const getSafeGuidance = (lowerText: string): string | undefined => {
  if (
    lowerText.includes("reboot") ||
    lowerText.includes("restart computer") ||
    lowerText.includes("restart my computer")
  ) {
    return rebootGuidance();
  }

  if (
    lowerText.includes("password") ||
    lowerText.includes("mfa") ||
    lowerText.includes("authenticator")
  ) {
    return [
      "**Account access issue:**",
      "",
      "For password, MFA, authenticator, or sign-in issues, I should collect details for IT instead of guessing.",
      "",
      "Tell me what system or account is affected, and I’ll help collect the details for IT.",
    ].join("\n");
  }

  if (
    lowerText.includes("hacked") ||
    lowerText.includes("compromised") ||
    lowerText.includes("phishing") ||
    lowerText.includes("suspicious email") ||
    lowerText.includes("malware") ||
    lowerText.includes("virus")
  ) {
    return [
      "**Possible security issue detected.**",
      "",
      "Do not click anything else, do not delete evidence, and do not forward suspicious content unless IT asks.",
      "",
      "Ok, let’s open a ticket for IT. Please type **/ticket** and include what happened, when it happened, and any suspicious sender, link, file, or message.",
    ].join("\n");
  }

  return undefined;
};

const helpMessage = (): string => {
  return [
    "**ReBath IT Helper**",
    "",
    "I can answer basic IT support questions, recommend safe first steps, and collect details for an IT ticket.",
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
    "You can also just tell me the issue, like **Outlook is frozen** or **my computer is slow**.",
  ].join("\n");
};

const ticketStatusMessage = (state: ConversationState): string => {
  if (state.mode !== "ticket" || !state.ticketStep) {
    return [
      "No ticket is currently in progress.",
      "",
      "Type **/ticket** to start a new IT support ticket.",
      state.lastTicketNumber ? `Last demo ticket: **${state.lastTicketNumber}**` : "",
    ]
      .filter(Boolean)
      .join("\n");
  }

  const ticket = state.ticket || {};

  return [
    "**Current ticket progress**",
    "",
    `**Step:** ${state.ticketStep}`,
    `**Requester:** ${ticket.requesterName || "Not provided yet"}`,
    `**Affected system:** ${ticket.affectedSystem || "Not provided yet"}`,
    `**Issue:** ${ticket.issueSummary || "Not provided yet"}`,
    `**Started:** ${ticket.startedAt || "Not provided yet"}`,
    `**Location:** ${ticket.location || "Not provided yet"}`,
    `**Device:** ${ticket.device || "Not provided yet"}`,
    `**Error message:** ${ticket.errorMessage || "Not provided yet"}`,
    `**Business impact:** ${ticket.businessImpact || "Not provided yet"}`,
    `**Others affected:** ${ticket.othersAffected || "Not provided yet"}`,
    `**Troubleshooting tried:** ${ticket.troubleshootingTried || "Not provided yet"}`,
    "",
    "Type **/cancel** to cancel this ticket intake.",
  ].join("\n");
};

const buildTicketSummary = (state: ConversationState): string => {
  const ticket = state.ticket || {};

  return [
    "**IT Support Ticket Summary**",
    "",
    `**Requester:** ${ticket.requesterName || "Not provided"}`,
    `**Teams user ID:** ${ticket.requesterTeamsId || "Not provided"}`,
    `**AAD object ID:** ${ticket.requesterAadObjectId || "Not provided"}`,
    `**Affected system:** ${ticket.affectedSystem || "Not provided"}`,
    `**Issue:** ${ticket.issueSummary || "Not provided"}`,
    `**Started:** ${ticket.startedAt || "Not provided"}`,
    `**Location:** ${ticket.location || "Not provided"}`,
    `**Device:** ${ticket.device || "Not provided"}`,
    `**Error message:** ${ticket.errorMessage || "Not provided"}`,
    `**Business impact:** ${ticket.businessImpact || "Not provided"}`,
    `**Others affected:** ${ticket.othersAffected || "Not provided"}`,
    `**Troubleshooting tried:** ${ticket.troubleshootingTried || "Not provided"}`,
    "",
    "This is ready to become a ServiceNow ticket.",
    "",
    "Type **confirm** to submit this ticket summary, or type **/cancel** to cancel.",
  ].join("\n");
};

const startTicketFlow = async (
  context: any,
  conversationId: string,
  state: ConversationState
): Promise<void> => {
  state.mode = "ticket";
  state.ticketStep = 1;
  state.ticket = {
    ...getTeamsUserInfo(context.activity),
  };

  saveConversationState(conversationId, state);

  await context.send(
    "Ok, let’s open a ticket for IT. What app, system, or service is affected? Example: Outlook, Teams, VPN, OneDrive, printer, internet."
  );
};

const startTicketFromExistingIssue = async (
  context: any,
  conversationId: string,
  state: ConversationState
): Promise<void> => {
  state.mode = "ticket";
  state.ticketStep = 3;

  state.ticket = {
    ...(state.ticket || {}),
    ...getTeamsUserInfo(context.activity),
    troubleshootingTried:
      state.ticket?.troubleshootingTried ||
      "Bot recommended basic troubleshooting. User reported the issue still exists.",
  };

  saveConversationState(conversationId, state);

  await context.send(
    [
      "Ok, let’s open a ticket for IT.",
      "",
      `I already have this recorded for **${state.ticket.requesterName || "you"}**:`,
      `**Affected system:** ${state.ticket.affectedSystem || "Not provided"}`,
      `**Issue:** ${state.ticket.issueSummary || "Not provided"}`,
      "",
      "When did this issue start? Example: today, yesterday, this morning, after a reboot, or after a recent update.",
    ].join("\n")
  );
};

const buildTroubleshootingSteps = (
  affectedSystem: string,
  lowerText: string
): string[] => {
  const isBrowserIssue =
    affectedSystem === "Browser / Website" ||
    lowerText.includes("website") ||
    lowerText.includes("browser") ||
    lowerText.includes("chrome") ||
    lowerText.includes("edge") ||
    lowerText.includes("page") ||
    lowerText.includes("site");

  if (isBrowserIssue) {
    return [
      "1. Close and reopen the browser.",
      "2. Try the site in an incognito/InPrivate browser window.",
      "3. Clear cache and cookies for the affected site.",
      "4. Try another browser if available.",
      "5. If it still fails, reboot once and test again.",
    ];
  }

  return [
    "1. Close and reopen the affected app.",
    "2. Save your work and reboot the computer once.",
    "3. After signing back in, test the issue again.",
    "4. If there is an error message, copy it or take a screenshot.",
  ];
};

app.on("message", async (context) => {
  const activity = context.activity;
  const conversationId = activity.conversation.id;
  const text: string = stripMentionsText(activity).trim();
  const lowerText = text.toLowerCase();

  console.log("Received Teams message:", {
    conversationId,
    conversationType: activity.conversation.conversationType,
    fromName: activity.from?.name,
    fromId: activity.from?.id,
    fromAadObjectId: activity.from?.aadObjectId,
    text,
  });

  if (!text) {
    await context.send(
      "ReBath IT Helper is online. Tell me what issue you are having, or type **/help** for commands."
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

  if (lowerText === "/reboot") {
    await context.send(rebootGuidance());
    return;
  }

  if (lowerText === "/status") {
    await context.send(ticketStatusMessage(state));
    return;
  }

  if (lowerText === "/count") {
    await context.send(`The count is ${state.count}`);
    return;
  }

  if (lowerText === "/state") {
    await context.send(JSON.stringify(state, null, 2));
    return;
  }

  if (lowerText === "/runtime") {
    const runtime = {
      nodeversion: process.version,
      sdkversion: "2.0.0",
      conversationType: activity.conversation.conversationType,
      fromName: activity.from?.name,
      fromId: activity.from?.id,
      fromAadObjectId: activity.from?.aadObjectId,
      hasBotId: Boolean(process.env.BOT_ID),
      hasClientId: Boolean(process.env.CLIENT_ID),
      hasClientSecret: Boolean(process.env.CLIENT_SECRET),
      hasTenantId: Boolean(process.env.TENANT_ID),
      hasOpenAiKey: Boolean(process.env.OPENAI_API_KEY),
    };

    await context.send(JSON.stringify(runtime, null, 2));
    return;
  }

  if (isGreeting(lowerText) || isGeneralHelpRequest(lowerText)) {
    await context.send(generalHelpMessage());
    return;
  }

  if (isTicketStartRequest(lowerText)) {
    await startTicketFlow(context, conversationId, state);
    return;
  }

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
      await startTicketFromExistingIssue(context, conversationId, state);
      return;
    }

    await context.send(
      "Did the troubleshooting fix the issue? Reply **fixed** if it is working now, or **still not working** if the issue continues."
    );

    return;
  }

  if (state.mode === "ticket" && state.ticketStep === 1) {
    state.ticket = {
      ...(state.ticket || {}),
      ...getTeamsUserInfo(activity),
    };

    state.ticket.affectedSystem = text;
    state.ticketStep = 2;

    saveConversationState(conversationId, state);

    await context.send(
      `Got it — affected system: **${text}**.\n\nWhat issue are you having? Example: "Outlook will not open", "VPN will not connect", or "printer is offline".`
    );

    return;
  }

  if (state.mode === "ticket" && state.ticketStep === 2) {
    state.ticket = {
      ...(state.ticket || {}),
      ...getTeamsUserInfo(activity),
    };

    state.ticket.issueSummary = text;
    state.ticketStep = 3;

    saveConversationState(conversationId, state);

    await context.send(
      "Thanks. When did this issue start? Example: today, yesterday, this morning, after a reboot, or after a recent update."
    );

    return;
  }

  if (state.mode === "ticket" && state.ticketStep === 3) {
    state.ticket = state.ticket || {};
    state.ticket.startedAt = text;
    state.ticketStep = 4;

    saveConversationState(conversationId, state);

    await context.send(
      "Got it. Where are you located? Example: remote, corporate office, or store location."
    );

    return;
  }

  if (state.mode === "ticket" && state.ticketStep === 4) {
    state.ticket = state.ticket || {};
    state.ticket.location = text;
    state.ticketStep = 5;

    saveConversationState(conversationId, state);

    await context.send(
      "Thanks. What device are you using? Example: Windows laptop, desktop, iPhone, iPad, shared store PC, or personal device."
    );

    return;
  }

  if (state.mode === "ticket" && state.ticketStep === 5) {
    state.ticket = state.ticket || {};
    state.ticket.device = text;
    state.ticketStep = 6;

    saveConversationState(conversationId, state);

    await context.send(
      "Do you see an error message? If yes, paste it here. If not, type **no error**."
    );

    return;
  }

  if (state.mode === "ticket" && state.ticketStep === 6) {
    state.ticket = state.ticket || {};
    state.ticket.errorMessage = text;
    state.ticketStep = 7;

    saveConversationState(conversationId, state);

    await context.send(
      "How is this impacting your work? Example: blocked from working, slowing me down, one task affected, question only, or store operations impacted."
    );

    return;
  }

  if (state.mode === "ticket" && state.ticketStep === 7) {
    state.ticket = state.ticket || {};
    state.ticket.businessImpact = text;
    state.ticketStep = 8;

    saveConversationState(conversationId, state);

    await context.send(
      "Is anyone else affected, or is it only you? Example: only me, multiple users, whole store, unknown."
    );

    return;
  }

  if (state.mode === "ticket" && state.ticketStep === 8) {
    state.ticket = state.ticket || {};
    state.ticket.othersAffected = text;
    state.ticketStep = 9;

    saveConversationState(conversationId, state);

    await context.send(
      "What troubleshooting have you already tried? Example: restarted app, rebooted computer, checked internet, reconnected VPN, tried another browser, or nothing yet."
    );

    return;
  }

  if (state.mode === "ticket" && state.ticketStep === 9) {
    state.ticket = state.ticket || {};
    state.ticket.troubleshootingTried = text;
    state.ticketStep = 10;

    saveConversationState(conversationId, state);

    await context.send(buildTicketSummary(state));

    return;
  }

  if (state.mode === "ticket" && state.ticketStep === 10) {
    if (isConfirmRequest(lowerText)) {
      const fakeTicketNumber = `DEMO-${Date.now().toString().slice(-6)}`;

      state.mode = "idle";
      state.ticketStep = undefined;
      state.lastTicketNumber = fakeTicketNumber;

      saveConversationState(conversationId, state);

      await context.send(
        [
          `Ticket submitted for demo purposes: **${fakeTicketNumber}**`,
          "",
          "Next build step: connect this confirmation step to the ServiceNow Incident API so it creates a real incident.",
        ].join("\n")
      );

      return;
    }

    await context.send(
      "Please type **confirm** to submit this ticket summary, or type **/cancel** to cancel."
    );

    return;
  }

  const safeGuidance = getSafeGuidance(lowerText);

  if (safeGuidance) {
    await context.send(safeGuidance);
    return;
  }

  if (looksLikeIssue(lowerText)) {
    const teamsUser = getTeamsUserInfo(activity);
    const affectedSystem = detectAffectedSystem(lowerText);
    const troubleshootingSteps = buildTroubleshootingSteps(affectedSystem, lowerText);

    state.mode = "awaitingTroubleshootingResult";
    state.ticketStep = undefined;
    state.ticket = {
      ...teamsUser,
      affectedSystem,
      issueSummary: text,
      troubleshootingTried: `Bot recommended basic troubleshooting: ${troubleshootingSteps.join(
        " "
      )}`,
    };

    saveConversationState(conversationId, state);

    await context.send(
      [
        `Thanks, ${teamsUser.requesterName}.`,
        "",
        `It sounds like an issue with **${affectedSystem}**.`,
        "",
        "Please try these safe first steps:",
        "",
        ...troubleshootingSteps,
        "",
        "After trying that, reply **fixed** if it is working, or **still not working** and I’ll open a ticket for IT.",
      ].join("\n")
    );

    return;
  }

  try {
    const aiResponse = await analyzeSupportMessage(text);

    await context.send(
      [
        aiResponse.reply,
        "",
        aiResponse.shouldCreateTicket
          ? "If the issue continues, reply with what is still happening and I can help open a ticket for IT."
          : "If you want IT to review this, type **/ticket**.",
      ].join("\n")
    );

    return;
  } catch (error) {
    console.error("AI response failed:", error);

    state.count++;
    state.mode = "idle";

    saveConversationState(conversationId, state);

    await context.send(
      [
        "ReBath IT Helper is online.",
        "",
        "Tell me what issue you are having, like **Outlook is frozen**, **VPN will not connect**, or **website is not loading**.",
        "",
        "I’ll suggest safe troubleshooting first. If it still does not work, I’ll help open a ticket for IT.",
      ].join("\n")
    );

    return;
  }
});

export default app;