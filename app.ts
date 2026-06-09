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
  mode?: "idle" | "ticket";
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
    "i need help",
    "need help",
    "support",
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
    "If the issue continues, type **/ticket** and I’ll collect the details for IT.",
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

  if (lowerText.includes("outlook")) {
    return [
      "**Basic Outlook check:**",
      "",
      "1. Close Outlook completely.",
      "2. Reopen Outlook.",
      "3. If it still fails, save your work and reboot once.",
      "4. If you see an error message, copy it or take a screenshot.",
      "",
      "If it still does not work, type **/ticket** and I’ll collect the details for IT.",
    ].join("\n");
  }

  if (lowerText.includes("teams")) {
    return [
      "**Basic Teams check:**",
      "",
      "1. Quit Teams completely.",
      "2. Reopen Teams.",
      "3. Check your internet connection.",
      "4. If audio/video is the issue, test another headset or speaker if available.",
      "5. If the issue continues, reboot once.",
      "",
      "If it still does not work, type **/ticket** and I’ll collect the details for IT.",
    ].join("\n");
  }

  if (lowerText.includes("vpn")) {
    return [
      "**Basic VPN check:**",
      "",
      "1. Confirm your internet connection is working.",
      "2. Disconnect and reconnect VPN.",
      "3. If VPN still fails, reboot once.",
      "4. Copy any VPN error message if one appears.",
      "",
      "If it still does not work, type **/ticket** and I’ll collect the details for IT.",
    ].join("\n");
  }

  if (lowerText.includes("printer") || lowerText.includes("print")) {
    return [
      "**Basic printer check:**",
      "",
      "1. Confirm the printer is powered on.",
      "2. Confirm you selected the correct printer.",
      "3. Try printing again.",
      "4. If you are at a store, check if others are affected too.",
      "5. If printing still fails, reboot once if it is your workstation.",
      "",
      "If it still does not work, type **/ticket** and I’ll collect the details for IT.",
    ].join("\n");
  }

  if (
    lowerText.includes("internet") ||
    lowerText.includes("network") ||
    lowerText.includes("wifi") ||
    lowerText.includes("wi-fi")
  ) {
    return [
      "**Basic network check:**",
      "",
      "1. Check if other websites or apps are working.",
      "2. Confirm whether you are on Wi-Fi, Ethernet, or VPN.",
      "3. If only one app is affected, restart that app.",
      "4. If everything is affected, reboot once and check if others nearby are impacted.",
      "",
      "If the issue continues, type **/ticket** and I’ll collect the details for IT.",
    ].join("\n");
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
      "Type **/ticket** and I’ll help create a support ticket.",
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
      "Type **/ticket** now and include what happened, when it happened, and any suspicious sender, link, file, or message.",
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
    "For anything advanced, urgent, security-related, or business-impacting, I will collect details for IT instead of guessing.",
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
    "**Safe recommendation:** If you have not already done so, save your work and reboot once if this is a workstation/app issue.",
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
  state.ticket = {};

  saveConversationState(conversationId, state);

  await context.send(
    "Let's create an IT support ticket. What app, system, or service is affected? Example: Outlook, Teams, VPN, OneDrive, printer, internet."
  );
};

app.on("message", async (context) => {
  const activity = context.activity;
  const conversationId = activity.conversation.id;
  const text: string = stripMentionsText(activity).trim();
  const lowerText = text.toLowerCase();

  console.log("Received Teams message:", {
    conversationId,
    conversationType: activity.conversation.conversationType,
    text,
  });

  if (!text) {
    await context.send(
      "ReBath IT Helper is online. Type **/ticket** to start a new IT support ticket, or type **/help** for commands."
    );
    return;
  }

  if (lowerText === "/reset") {
    clearConversationState(conversationId);

    await context.send(
      "Ok, I cleared the current conversation. Type **/ticket** to start a new IT support ticket."
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
      "Canceled the current ticket intake. Type **/ticket** to start a new one."
    );

    return;
  }

  if (lowerText === "/help" || lowerText === "help") {
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
      hasBotId: Boolean(process.env.BOT_ID),
      hasClientId: Boolean(process.env.CLIENT_ID),
      hasClientSecret: Boolean(process.env.CLIENT_SECRET),
      hasTenantId: Boolean(process.env.TENANT_ID),
      hasOpenAiKey: Boolean(process.env.OPENAI_API_KEY),
    };

    await context.send(JSON.stringify(runtime, null, 2));
    return;
  }

  if (isTicketStartRequest(lowerText)) {
    await startTicketFlow(context, conversationId, state);
    return;
  }

  if (state.mode === "ticket" && state.ticketStep === 1) {
    state.ticket = state.ticket || {};
    state.ticket.affectedSystem = text;
    state.ticketStep = 2;

    saveConversationState(conversationId, state);

    await context.send(
      `Got it — affected system: **${text}**.\n\nWhat issue are you having? Example: "Outlook will not open", "VPN will not connect", or "printer is offline".`
    );

    return;
  }

  if (state.mode === "ticket" && state.ticketStep === 2) {
    state.ticket = state.ticket || {};
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

  try {
    const aiResponse = await analyzeSupportMessage(text);

    await context.send(
      [
        aiResponse.reply,
        "",
        aiResponse.shouldCreateTicket
          ? "I can create a support ticket for this. Type **/ticket** to start."
          : "Type **/ticket** if you want IT to review this.",
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
        "I can help with basic IT questions or collect details for a support ticket.",
        "",
        "Type **/ticket** to start a new IT support ticket.",
        "Type **/help** to see available commands.",
      ].join("\n")
    );

    return;
  }
});

export default app;