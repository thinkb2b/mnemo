import { GoogleGenAI } from "@google/genai";
// WICHTIG: Explizite Dateiendung für Browser-Importe
import { SnippetFormData } from "../types.ts";

const getAiClient = () => {
  // Zugriff auf globale Variable, die in index.html gesetzt wurde
  const apiKey = (window as any).process?.env?.API_KEY;
  
  if (!apiKey || apiKey.includes("HIER_IHREN")) {
    console.error("API Key missing or invalid.");
    // Wir werfen hier keinen harten Fehler beim Init, damit die App nicht crasht,
    // sondern erst beim Aufruf der Funktion.
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