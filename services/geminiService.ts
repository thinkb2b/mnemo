import { GoogleGenAI } from "@google/genai";
import { SnippetFormData } from "../types";

const getAiClient = () => {
  const apiKey = process.env.API_KEY;
  if (!apiKey) {
    throw new Error("API Key is missing. Please check your environment variables.");
  }
  return new GoogleGenAI({ apiKey });
};

export const generateSnippet = async (prompt: string): Promise<Partial<SnippetFormData>> => {
  const ai = getAiClient();
  
  const systemInstruction = `
    Du bist ein Assistent für professionelle E-Mail-Kommunikation.
    Erstelle basierend auf der Anfrage des Benutzers einen E-Mail-Textbaustein.
    
    WICHTIG:
    - Identifiziere Variablen im Text und markiere sie mit geschweiften Klammern, z.B. {Name}, {Datum}, {Rechnungsnummer}.
    - Gib das Ergebnis NUR als valides JSON zurück.
    - Das JSON muss folgende Felder haben: "title" (kurzer interner Name), "subject" (Betreffzeile), "body" (E-Mail Text).
    - Der "body" soll HTML-Zeilenumbrüche (<br/>) verwenden.
  `;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-2.5-flash",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        systemInstruction: systemInstruction,
      },
    });

    const text = response.text;
    if (!text) return {};

    const data = JSON.parse(text);
    return {
      title: data.title,
      subject: data.subject,
      body: data.body,
    };
  } catch (error) {
    console.error("Gemini generation error:", error);
    throw error;
  }
};

export const improveText = async (text: string): Promise<string> => {
  const ai = getAiClient();
  try {
    const response = await ai.models.generateContent({
      model: "gemini-2.5-flash",
      contents: `Verbessere den folgenden E-Mail-Text, mache ihn professioneller und freundlicher:\n\n${text}`,
    });
    return response.text || text;
  } catch (error) {
    console.error("Gemini improvement error:", error);
    return text;
  }
};