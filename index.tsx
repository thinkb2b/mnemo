import React, { useState, useEffect } from 'react';
import ReactDOM from 'react-dom/client';
import { 
  Layout, Plus, Settings, Mail, Search, Folder, Edit3, Trash2, Send, 
  ArrowLeft, Sparkles, RefreshCw, X, Save, Link, Link2Off, Info, CheckCircle
} from 'lucide-react';
import { GoogleGenAI } from "@google/genai";

// --- Types ---
interface Snippet {
  id: string;
  title: string;
  subject: string;
  body: string;
  groupId: string;
  variables: string[];
}

interface Group {
  id: string;
  name: string;
  color: string;
}

type ViewState = 'LIST' | 'CREATE' | 'EDIT' | 'FILL_VARS' | 'INFO' | 'SETTINGS';

// --- AI Service ---
const generateAiSnippet = async (prompt: string) => {
  const apiKey = (window as any).process?.env?.API_KEY;
  if (!apiKey) throw new Error("API Key fehlt.");
  
  const ai = new GoogleGenAI({ apiKey });
  const systemInstruction = `
    Du bist ein Assistent für professionelle E-Mail-Kommunikation.
    Erstelle basierend auf der Anfrage einen E-Mail-Textbaustein.
    WICHTIG:
    - Markiere Variablen mit geschweiften Klammern, z.B. {Name}.
    - Gib das Ergebnis NUR als JSON zurück: {"title": "...", "subject": "...", "body": "..."}.
    - Nutze <br/> für Zeilenumbrüche im "body".
  `;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-3-flash-preview",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        systemInstruction: systemInstruction,
      },
    });
    return JSON.parse(response.text || '{}');
  } catch (error) {
    console.error("AI Error:", error);
    throw error;
  }
};

// --- Components ---
const Button = ({ children, onClick, variant = 'primary', className = '', icon: Icon, disabled = false }: any) => {
  const variants: any = {
    primary: "bg-blue-600 text-white hover:bg-blue-700 disabled:bg-blue-300",
    secondary: "bg-white text-gray-700 border border-gray-300 hover:bg-gray-50 disabled:bg-gray-100",
    ghost: "bg-transparent text-slate-500 hover:bg-slate-100",
  };
  return (
    <button 
      onClick={onClick} 
      disabled={disabled}
      className={`flex items-center justify-center px-4 py-2 rounded-md font-medium text-sm transition-all focus:outline-none ${variants[variant]} ${className}`}
    >
      {Icon && <Icon className="w-4 h-4 mr-2" />}
      {children}
    </button>
  );
};

const App = () => {
  const [view, setView] = useState<ViewState>('LIST');
  const [snippets, setSnippets] = useState<Snippet[]>([
    {
      id: 's1',
      groupId: 'g1',
      title: 'Beispiel: Termin',
      subject: 'Termin am {Datum}',
      body: 'Hallo {Name},<br/><br/>hiermit bestätige ich unseren Termin am {Datum}.',
      variables: ['Name', 'Datum']
    }
  ]);
  const [groups] = useState<Group[]>([
    { id: 'g1', name: 'Allgemein', color: 'bg-blue-500' },
    { id: 'g2', name: 'Vertrieb', color: 'bg-green-500' }
  ]);
  
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedGroup, setSelectedGroup] = useState<string | null>(null);
  const [currentSnippet, setCurrentSnippet] = useState<Snippet | null>(null);
  const [variableValues, setVariableValues] = useState<Record<string, string>>({});
  const [editorData, setEditorData] = useState({ title: '', subject: '', body: '', groupId: 'g1' });
  const [isGenerating, setIsGenerating] = useState(false);
  const [isOfficeReady, setIsOfficeReady] = useState(false);

  useEffect(() => {
    if ((window as any).Office) {
      (window as any).Office.onReady(() => setIsOfficeReady(true));
    }
  }, []);

  const extractVariables = (text: string): string[] => {
    const regex = /\{([^}]+)\}/g;
    const matches = new Set<string>();
    let match;
    while ((match = regex.exec(text)) !== null) matches.add(match[1]);
    return Array.from(matches);
  };

  const handleSaveSnippet = () => {
    const vars = extractVariables(editorData.subject + editorData.body);
    if (view === 'CREATE') {
      setSnippets([...snippets, { id: Date.now().toString(), ...editorData, variables: vars }]);
    } else {
      setSnippets(snippets.map(s => s.id === currentSnippet?.id ? { ...s, ...editorData, variables: vars } : s));
    }
    setView('LIST');
  };

  const handleAiAction = async (prompt: string) => {
    if (!prompt) return;
    setIsGenerating(true);
    try {
      const result = await generateAiSnippet(prompt);
      setEditorData(prev => ({ ...prev, ...result }));
    } catch (e) {
      alert("KI Fehler: " + e.message);
    } finally {
      setIsGenerating(false);
    }
  };

  const executeInsert = (snippet: Snippet, values: Record<string, string>) => {
    let finalSubject = snippet.subject;
    let finalBody = snippet.body;
    Object.entries(values).forEach(([k, v]) => {
      const reg = new RegExp(`\\{${k}\\}`, 'g');
      finalSubject = finalSubject.replace(reg, v);
      finalBody = finalBody.replace(reg, v);
    });

    if (isOfficeReady) {
      const office = (window as any).Office;
      office.context.mailbox.item.body.setSelectedDataAsync(finalBody, { coercionType: office.CoercionType.Html });
      office.context.mailbox.item.subject.setAsync(finalSubject);
      setView('LIST');
    } else {
      navigator.clipboard.writeText(`Betreff: ${finalSubject}\n\n${finalBody.replace(/<br\/>/g, '\n')}`);
      alert("In Zwischenablage kopiert (Browser)");
      setView('LIST');
    }
  };

  return (
    <div className="flex h-screen w-full bg-white text-slate-900 overflow-hidden select-none">
      {/* Sidebar */}
      <div className="w-14 bg-slate-50 border-r border-slate-200 flex flex-col items-center py-4 space-y-4">
        <button onClick={() => setView('LIST')} className={`p-2 rounded-lg ${view === 'LIST' ? 'bg-blue-100 text-blue-600' : 'text-slate-400'}`}><Layout size={20}/></button>
        <button onClick={() => { setEditorData({title:'', subject:'', body:'', groupId:'g1'}); setView('CREATE'); }} className={`p-2 rounded-lg ${view === 'CREATE' ? 'bg-blue-100 text-blue-600' : 'text-slate-400'}`}><Plus size={20}/></button>
        <div className="flex-grow" />
        <button onClick={() => setView('SETTINGS')} className={`p-2 rounded-lg ${view === 'SETTINGS' ? 'bg-blue-100 text-blue-600' : 'text-slate-400'}`}><Settings size={20}/></button>
      </div>

      {/* Content */}
      <div className="flex-1 flex flex-col min-w-0">
        {view === 'LIST' && (
          <>
            <div className="p-4 border-b border-slate-200 bg-white">
              <h1 className="text-lg font-bold mb-3">Bibliothek</h1>
              <div className="relative">
                <Search size={16} className="absolute left-3 top-2.5 text-slate-400" />
                <input 
                  type="text" placeholder="Suchen..." 
                  className="w-full pl-9 pr-4 py-2 bg-slate-100 border-none rounded-md text-sm focus:ring-2 focus:ring-blue-500"
                  value={searchQuery} onChange={e => setSearchQuery(e.target.value)}
                />
              </div>
            </div>
            <div className="flex-1 overflow-y-auto p-4 space-y-3">
              {snippets.filter(s => s.title.toLowerCase().includes(searchQuery.toLowerCase())).map(s => (
                <div 
                  key={s.id} onClick={() => { setCurrentSnippet(s); setVariableValues({}); s.variables.length ? setView('FILL_VARS') : executeInsert(s, {}); }}
                  className="p-3 bg-white border border-slate-200 rounded-lg hover:border-blue-400 cursor-pointer shadow-sm transition-all"
                >
                  <h3 className="font-semibold text-sm truncate">{s.title}</h3>
                  <p className="text-xs text-slate-500 line-clamp-1">{s.subject}</p>
                </div>
              ))}
            </div>
          </>
        )}

        {(view === 'CREATE' || view === 'EDIT') && (
          <div className="flex flex-col h-full">
            <div className="p-4 border-b flex items-center justify-between">
              <button onClick={() => setView('LIST')} className="text-slate-400"><ArrowLeft size={20}/></button>
              <h2 className="font-bold text-sm">Snippet</h2>
              <Button onClick={handleSaveSnippet}>Speichern</Button>
            </div>
            <div className="flex-1 overflow-y-auto p-4 space-y-4">
              <div className="bg-blue-50 p-3 rounded-lg border border-blue-100">
                <div className="flex items-center gap-2 mb-2 text-blue-700 font-bold text-[10px] uppercase"><Sparkles size={12}/> KI Generator</div>
                <div className="flex gap-2">
                  <input id="ai-prompt" className="flex-1 text-sm border-none rounded p-2" placeholder="z.B. Terminbestätigung..." />
                  <Button onClick={() => handleAiAction((document.getElementById('ai-prompt') as any).value)} disabled={isGenerating} className="px-2">
                    {isGenerating ? '...' : <Send size={14}/>}
                  </Button>
                </div>
              </div>
              <input className="w-full border p-2 rounded text-sm" placeholder="Titel" value={editorData.title} onChange={e => setEditorData({...editorData, title: e.target.value})} />
              <input className="w-full border p-2 rounded text-sm" placeholder="Betreff" value={editorData.subject} onChange={e => setEditorData({...editorData, subject: e.target.value})} />
              <textarea className="w-full flex-1 border p-2 rounded text-sm min-h-[150px] font-mono" placeholder="Inhalt..." value={editorData.body} onChange={e => setEditorData({...editorData, body: e.target.value})} />
            </div>
          </div>
        )}

        {view === 'FILL_VARS' && (
          <div className="flex flex-col h-full">
            <div className="p-4 border-b flex items-center"><button onClick={() => setView('LIST')} className="mr-3 text-slate-400"><ArrowLeft size={20}/></button><h2 className="font-bold text-sm">Variablen</h2></div>
            <div className="flex-1 p-4 space-y-4">
              {currentSnippet?.variables.map(v => (
                <div key={v}>
                  <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">{v}</label>
                  <input className="w-full border p-2 rounded text-sm" placeholder={v} value={variableValues[v] || ''} onChange={e => setVariableValues({...variableValues, [v]: e.target.value})} />
                </div>
              ))}
            </div>
            <div className="p-4 border-t flex gap-2">
              <Button onClick={() => currentSnippet && executeInsert(currentSnippet, variableValues)} className="w-full">Einfügen</Button>
            </div>
          </div>
        )}

        {view === 'SETTINGS' && (
          <div className="p-6">
            <div className="flex items-center mb-6"><button onClick={() => setView('LIST')} className="mr-3 text-slate-400"><ArrowLeft size={20}/></button><h2 className="font-bold">Settings</h2></div>
            <div className={`p-4 rounded-lg flex items-center gap-3 ${isOfficeReady ? 'bg-green-50 text-green-700' : 'bg-slate-50 text-slate-500'}`}>
              {isOfficeReady ? <CheckCircle size={20}/> : <Link2Off size={20}/>}
              <span className="text-sm font-medium">{isOfficeReady ? 'Outlook aktiv' : 'Browser Modus'}</span>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

const root = ReactDOM.createRoot(document.getElementById('root')!);
root.render(<App />);