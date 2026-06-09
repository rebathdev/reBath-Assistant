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

export const analyzeSupportMessage = async (
  userMessage: string
): Promise<AiSupportResponse> => {
  if (!process.env.OPENAI_API_KEY) {
    return fallbackResponse(userMessage);
  }

  try {
    const response = await client.responses.create({
      model: "gpt-5.1-mini",
      input: [
        {
          role: "system",
          content: [
            "You are ReBath IT Helper, a safe IT support intake assistant.",
            "You answer basic IT support questions and recommend only safe low-risk steps.",
            "Allowed steps: restart the affected app, sign out and back in, check internet, reconnect VPN, try another browser, reboot once, gather screenshots or errors, create a ticket.",
            "Never provide admin commands, scripts, PowerShell, registry edits, security bypasses, access changes, password reset bypasses, deletion steps, or advanced troubleshooting.",
            "If the issue sounds security-related, urgent, multi-user, store-impacting, account-compromise-related, data-loss-related, or business-blocking, recommend creating a ticket.",
            "Keep reply concise, professional, and helpful.",
            "Return only valid JSON matching the schema.",
          ].join("\n"),
        },
        {
          role: "user",
          content: userMessage,
        },
      ],
      text: {
        format: {
          type: "json_schema",
          name: "support_response",
          strict: true,
          schema: {
            type: "object",
            additionalProperties: false,
            properties: {
              reply: { type: "string" },
              category: { type: "string" },
              subcategory: { type: "string" },
              impact: {
                type: "string",
                enum: ["1", "2", "3"],
              },
              urgency: {
                type: "string",
                enum: ["1", "2", "3"],
              },
              shouldCreateTicket: { type: "boolean" },
              securitySensitive: { type: "boolean" },
              shortDescription: { type: "string" },
              serviceNowDescription: { type: "string" },
            },
            required: [
              "reply",
              "category",
              "subcategory",
              "impact",
              "urgency",
              "shouldCreateTicket",
              "securitySensitive",
              "shortDescription",
              "serviceNowDescription",
            ],
          },
        },
      },
    });

    const outputText = response.output_text;

    if (!outputText) {
      return fallbackResponse(userMessage);
    }

    return JSON.parse(outputText) as AiSupportResponse;
  } catch (error) {
    console.error("AI support analysis failed:", error);
    return fallbackResponse(userMessage);
  }
};