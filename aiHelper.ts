import OpenAI from "openai";

export interface AiSupportResponse {
  reply: string;
  category: string;
  subcategory: string;
  impact: "1" | "2" | "3";
  urgency: "1" | "2" | "3";
  shouldCreateTicket: boolean;
  securitySensitive: boolean;
  shortDescription: string;
  serviceNowDescription: string;
}

const client = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

const fallbackResponse = (message: string): AiSupportResponse => {
  return {
    reply:
      "I can help with basic IT questions or collect details for a support ticket. For anything beyond simple checks, type /ticket and I’ll collect the details for IT.",
    category: "Inquiry",
    subcategory: "General",
    impact: "3",
    urgency: "3",
    shouldCreateTicket: message.trim().length > 0,
    securitySensitive: false,
    shortDescription: "General IT support request",
    serviceNowDescription: message,
  };
};

const normalizeAiResponse = (value: Partial<AiSupportResponse>, originalMessage: string): AiSupportResponse => {
  return {
    reply:
      typeof value.reply === "string" && value.reply.trim()
        ? value.reply
        : fallbackResponse(originalMessage).reply,
    category:
      typeof value.category === "string" && value.category.trim()
        ? value.category
        : "Inquiry",
    subcategory:
      typeof value.subcategory === "string" && value.subcategory.trim()
        ? value.subcategory
        : "General",
    impact: value.impact === "1" || value.impact === "2" || value.impact === "3" ? value.impact : "3",
    urgency: value.urgency === "1" || value.urgency === "2" || value.urgency === "3" ? value.urgency : "3",
    shouldCreateTicket:
      typeof value.shouldCreateTicket === "boolean"
        ? value.shouldCreateTicket
        : true,
    securitySensitive:
      typeof value.securitySensitive === "boolean"
        ? value.securitySensitive
        : false,
    shortDescription:
      typeof value.shortDescription === "string" && value.shortDescription.trim()
        ? value.shortDescription
        : "General IT support request",
    serviceNowDescription:
      typeof value.serviceNowDescription === "string" && value.serviceNowDescription.trim()
        ? value.serviceNowDescription
        : originalMessage,
  };
};

export const analyzeSupportMessage = async (
  userMessage: string
): Promise<AiSupportResponse> => {
  if (!process.env.OPENAI_API_KEY) {
    return fallbackResponse(userMessage);
  }

  try {
    const completion = await client.chat.completions.create({
      model: "gpt-4o-mini",
      response_format: { type: "json_object" },
      messages: [
        {
          role: "system",
          content: [
            "You are ReBath IT Helper, a safe IT support intake assistant.",
            "You answer basic IT support questions and recommend only safe low-risk steps.",
            "Allowed steps: restart the affected app, sign out and back in, check internet, reconnect VPN, try another browser, reboot once, gather screenshots or errors, create a ticket.",
            "Never provide admin commands, scripts, PowerShell, registry edits, security bypasses, access changes, password reset bypasses, deletion steps, or advanced troubleshooting.",
            "If the issue sounds security-related, urgent, multi-user, store-impacting, account-compromise-related, data-loss-related, or business-blocking, recommend creating a ticket.",
            "Keep reply concise, professional, and helpful.",
            "Return only valid JSON with this exact shape:",
            JSON.stringify({
              reply: "string",
              category: "string",
              subcategory: "string",
              impact: "1 | 2 | 3",
              urgency: "1 | 2 | 3",
              shouldCreateTicket: true,
              securitySensitive: false,
              shortDescription: "string",
              serviceNowDescription: "string",
            }),
          ].join("\n"),
        },
        {
          role: "user",
          content: userMessage,
        },
      ],
      temperature: 0.2,
    });

    const content = completion.choices[0]?.message?.content;

    if (!content) {
      return fallbackResponse(userMessage);
    }

    const parsed = JSON.parse(content) as Partial<AiSupportResponse>;
    return normalizeAiResponse(parsed, userMessage);
  } catch (error) {
    console.error("AI support analysis failed:", error);
    return fallbackResponse(userMessage);
  }
};