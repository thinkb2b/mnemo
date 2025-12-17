import React, { useState, useEffect } from 'react';
import ReactDOM from 'react-dom/client';
import { 
  Layout, Plus, Settings, Mail, Search, Folder, Edit3, Trash2, Send, 
  ArrowLeft, Sparkles, RefreshCw, X, Save, Link, Link2Off, Info
} from 'lucide-react';
import { GoogleGenAI } from "@google/genai";

// --- 1. TYPES ---

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

interface SenderAccount {
  id: string;
  email: string;
  name: string;
  signature: string;
}

type ViewState = 'LIST' | 'CREATE' | 'EDIT' | 'FILL_VARS' | 'INFO' | 'SETTINGS';

interface SnippetFormData {
  title: string;
  subject: string;
  body: string;
  groupId: string;
}

// --- 2. SERVICES ---

const getAiClient = () => {
  const apiKey = (window as any).process?.env?.API_KEY;
  if (!apiKey || apiKey.includes("HIER_IHREN")) {
    console.warn("API Key might be missing.");
  }
  return new GoogleGenAI({ apiKey });
};

const generateSnippet = async (prompt: string): Promise<Partial<SnippetFormData>> => {
  const ai = getAiClient();
  const systemInstruction = `
    Du bist ein Assistent für professionelle E-Mail-Kommunikation.
    Erstelle basierend auf der Anfrage des Benutzers einen E-Mail-Textbaustein.
    WICHTIG:
    - Identifiziere Variablen im Text und markiere sie mit geschweiften Klammern, z.B. {Name}.
    - Gib das Ergebnis NUR als valides JSON zurück. Feldern: "title", "subject", "body".
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
    return JSON.parse(text);
  } catch (error) {
    console.error("Gemini error:", error);
    throw error;
  }
};

// --- 3. DATA ---

const AVAILABLE_COLORS = [
  { name: 'Blau', class: 'bg-blue-500' },
  { name: 'Grün', class: 'bg-green-500' },
  { name: 'Orange', class: 'bg-orange-500' },
  { name: 'Lila', class: 'bg-purple-500' },
  { name: 'Rot', class: 'bg-red-500' },
  { name: 'Türkis', class: 'bg-teal-500' },
  { name: 'Grau', class: 'bg-gray-500' },
  { name: 'Pink', class: 'bg-pink-500' },
];

const INITIAL_GROUPS: Group[] = [
  { id: 'g1', name: 'Allgemein', color: 'bg-blue-500' },
  { id: 'g2', name: 'Vertrieb', color: 'bg-green-500' },
  { id: 'g3', name: 'Support', color: 'bg-orange-500' },
  { id: 'g4', name: 'HR', color: 'bg-purple-500' },
];

const INITIAL_ACCOUNTS: SenderAccount[] = [
  { id: 'acc1', name: 'Max Mustermann', email: 'max@firma.de', signature: '\n\nMit freundlichen Grüßen,\nMax Mustermann' },
];

const INITIAL_SNIPPETS: Snippet[] = [
  {
    id: 's1',
    groupId: 'g1',
    title: 'Terminbestätigung',
    subject: 'Bestätigung unseres Termins am {Datum}',
    body: 'Hallo {Name},\n\nhiermit bestätige ich unseren Termin am {Datum} um {Uhrzeit}.\n\nIch freue mich auf das Gespräch.',
    variables: ['Name', 'Datum', 'Uhrzeit']
  }
];

// --- 4. APP COMPONENT ---

const Button = ({ children, onClick, variant = 'primary', className = '', icon: Icon, disabled = false }: any) => {
  const baseStyle = "flex items-center justify-center px-4 py-2 rounded-md font-medium transition-colors text-sm focus:outline-none focus:ring-2 focus:ring-offset-1";
  const variants: any = {
    primary: "bg-blue-600 text-white hover:bg-blue-700 focus:ring-blue-500 disabled:bg-blue-300",
    secondary: "bg-white text-gray-700 border border-gray-300 hover:bg-gray-50 focus:ring-gray-200 disabled:bg-gray-100",
    ghost: "bg-transparent text-gray-600 hover:bg-gray-100 focus:ring-gray-200",
    danger: "bg-red-50 text-red-600 hover:bg-red-100 focus:ring-red-200"
  };
  return (
    <button onClick={onClick} className={`${baseStyle} ${variants[variant]} ${className}`} disabled={disabled}>
      {Icon && <Icon className="w-4 h-4 mr-2" />}
      {children}
    </button>
  );
};

function App() {
  const [view, setView] = useState<ViewState>('LIST');
  const [snippets, setSnippets] = useState<Snippet[]>(INITIAL_SNIPPETS);
  const [groups, setGroups] = useState<Group[]>(INITIAL_GROUPS);
  const [accounts, setAccounts] = useState<SenderAccount[]>(INITIAL_ACCOUNTS);
  
  const [selectedGroup, setSelectedGroup] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState('');
  const [currentSnippet, setCurrentSnippet] = useState<Snippet | null>(null);
  const [variableValues, setVariableValues] = useState<Record<string, string>>({});
  const [selectedAccount, setSelectedAccount] = useState<string>(INITIAL_ACCOUNTS[0]?.id || '');
  
  const [newGroupName, setNewGroupName] = useState('');
  const [newGroupColor, setNewGroupColor] = useState(AVAILABLE_COLORS[0].class);
  const [editingGroupId, setEditingGroupId] = useState<string | null>(null);
  const [isSyncing, setIsSyncing] = useState(false);
  
  const [isOfficeInitialized, setIsOfficeInitialized] = useState(false);
  const [editorData, setEditorData] = useState<SnippetFormData>({ title: '', subject: '', body: '', groupId: 'g1' });
  const [isGenerating, setIsGenerating] = useState(false);

  useEffect(() => {
    const initOffice = async () => {
        if (typeof window !== 'undefined' && (window as any).Office) {
            try {
                await (window as any).Office.onReady((info: any) => {
                    if (info.host) setIsOfficeInitialized(true);
                });
            } catch (e) { console.error(e); }
        }
    };
    initOffice();
  }, []);

  const extractVariables = (text: string): string[] => {
    const regex = /\{([^}]+)\}/g;
    const matches = new Set<string>();
    let match;
    while ((match = regex.exec(text)) !== null) matches.add(match[1]);
    return Array.from(matches);
  };

  const handleCreate = () => {
    setEditorData({ title: '', subject: '', body: '', groupId: groups[0]?.id || 'g1' });
    setCurrentSnippet(null);
    setView('CREATE');
  };

  const handleEdit = (snippet: Snippet) => {
    setEditorData({ ...snippet });
    setCurrentSnippet(snippet);
    setView('EDIT');
  };

  const handleDelete = (id: string) => {
    if (confirm('Snippet löschen?')) setSnippets(prev => prev.filter(s => s.id !== id));
  };

  const handleSave = () => {
    const variables = extractVariables(editorData.subject + ' ' + editorData.body);
    if (view === 'CREATE') {
      setSnippets([...snippets, { id: Date.now().toString(), ...editorData, variables }]);
    } else if (view === 'EDIT' && currentSnippet) {
      setSnippets(snippets.map(s => s.id === currentSnippet.id ? { ...s, ...editorData, variables } : s));
    }
    setView('LIST');
  };

  const handleAiGenerate = async (prompt: string) => {
    setIsGenerating(true);
    try {
      const result = await generateSnippet(prompt);
      setEditorData(prev => ({ ...prev, ...result }));
    } catch (e) { alert("KI Fehler: Prüfen Sie den API Key."); }
    finally { setIsGenerating(false); }
  };

  const handleSaveGroup = () => {
    if (!newGroupName.trim()) return;
    if (editingGroupId) {
        setGroups(prev => prev.map(g => g.id === editingGroupId ? { ...g, name: newGroupName, color: newGroupColor } : g));
        setEditingGroupId(null);
    } else {
        setGroups([...groups, { id: `g${Date.now()}`, name: newGroupName, color: newGroupColor }]);
    }
    setNewGroupName('');
  };

  const handleSyncAccounts = async () => {
    setIsSyncing(true);
    setTimeout(() => {
      setAccounts([...INITIAL_ACCOUNTS, { id: 'acc_synced', name: 'Team', email: 'team@firma.de', signature: '\n\nTeam' }]);
      setIsSyncing(false);
    }, 800);
  };

  const executeInsert = async (snippet: Snippet, values: Record<string, string>) => {
    let finalSubject = snippet.subject;
    let finalBody = snippet.body;
    Object.entries(values).forEach(([key, val]) => {
      const regex = new RegExp(`\\{${key}\\}`, 'g');
      finalSubject = finalSubject.replace(regex, val);
      finalBody = finalBody.replace(regex, val);
    });

    const account = accounts.find(a => a.id === selectedAccount);
    if (account) finalBody += account.signature;

    if (isOfficeInitialized) {
        const office = (window as any).Office;
        office.context.mailbox.item.body.setSelectedDataAsync(
            finalBody, { coercionType: office.CoercionType.Html }, 
            (res: any) => {
                if (res.status === 'succeeded') {
                    office.context.mailbox.item.subject.setAsync(finalSubject, () => {});
                    setView('LIST');
                }
            }
        );
    } else {
        const fullText = `[Von: ${account?.email}]\nBetreff: ${finalSubject}\n\n${finalBody}`;
        navigator.clipboard.writeText(fullText);
        alert(`Simuliert (in Clipboard):\n${fullText}`);
        setView('LIST');
    }
  };

  // --- VIEW RENDERING ---

  const renderSidebar = () => (
    <div className="w-16 bg-white border-r border-gray-200 flex flex-col items-center py-4 space-y-4">
      <button onClick={() => setView('LIST')} className={`p-2 rounded-xl ${view === 'LIST' ? 'bg-blue-100 text-blue-600' : 'text-gray-400'}`}><Layout className="w-6 h-6" /></button>
      <button onClick={handleCreate} className={`p-2 rounded-xl ${view === 'CREATE' ? 'bg-blue-100 text-blue-600' : 'text-gray-400'}`}><Plus className="w-6 h-6" /></button>
      <div className="flex-grow" />
      <button onClick={() => setView('SETTINGS')} className={`p-2 rounded-xl ${view === 'SETTINGS' ? 'bg-blue-100 text-blue-600' : 'text-gray-400'}`}><Settings className="w-6 h-6" /></button>
      <button onClick={() => setView('INFO')} className={`p-2 rounded-xl ${view === 'INFO' ? 'bg-blue-100 text-blue-600' : 'text-gray-400'}`}><Info className="w-6 h-6" /></button>
    </div>
  );

  const renderSnippetList = () => {
    const filtered = snippets.filter(s => {
      const matchG = selectedGroup ? s.groupId === selectedGroup : true;
      const matchS = s.title.toLowerCase().includes(searchQuery.toLowerCase());
      return matchG && matchS;
    });

    return (
      <div className="flex flex-col h-full bg-gray-50">
        <div className="bg-white p-4 border-b border-gray-200 space-y-3 sticky top-0 z-10">
          <div className="flex justify-between items-center">
             <h1 className="text-xl font-bold">Bibliothek</h1>
             <select className="border rounded p-1 text-sm max-w-[120px]" value={selectedAccount} onChange={e => setSelectedAccount(e.target.value)}>
                {accounts.map(a => <option key={a.id} value={a.id}>{a.name}</option>)}
             </select>
          </div>
          <div className="relative">
            <Search className="absolute left-3 top-2.5 w-4 h-4 text-gray-400" />
            <input type="text" placeholder="Suchen..." className="w-full pl-9 pr-4 py-2 bg-gray-100 rounded-md text-sm" value={searchQuery} onChange={e => setSearchQuery(e.target.value)} />
          </div>
          <div className="flex gap-2 overflow-x-auto pb-1">
             <button onClick={() => setSelectedGroup(null)} className={`px-3 py-1 rounded-full text-xs font-bold whitespace-nowrap ${!selectedGroup ? 'bg-gray-800 text-white' : 'bg-gray-200'}`}>Alle</button>
             {groups.map(g => (
               <button key={g.id} onClick={() => setSelectedGroup(g.id)} className={`px-3 py-1 rounded-full text-xs whitespace-nowrap flex items-center gap-1 ${selectedGroup === g.id ? 'bg-blue-600 text-white' : 'bg-white border'}`}>
                 <div className={`w-2 h-2 rounded-full ${g.color}`} /> {g.name}
               </button>
             ))}
          </div>
        </div>
        <div className="flex-1 overflow-y-auto p-4 space-y-3">
          {filtered.map(s => {
              const grp = groups.find(g => g.id === s.groupId);
              return (
                <div key={s.id} onClick={() => { setCurrentSnippet(s); setVariableValues({}); s.variables.length ? setView('FILL_VARS') : executeInsert(s, {}); }} 
                     className="bg-white p-4 rounded-lg shadow-sm border hover:shadow-md cursor-pointer group relative">
                  <div className="flex justify-between mb-1">
                    <span className={`text-[10px] px-2 py-0.5 rounded text-white ${grp?.color || 'bg-gray-400'}`}>{grp?.name}</span>
                    <div className="hidden group-hover:flex gap-1">
                        <button onClick={(e) => { e.stopPropagation(); handleEdit(s); }} className="p-1 text-gray-400 hover:text-blue-500"><Edit3 className="w-3 h-3" /></button>
                        <button onClick={(e) => { e.stopPropagation(); handleDelete(s.id); }} className="p-1 text-gray-400 hover:text-red-500"><Trash2 className="w-3 h-3" /></button>
                    </div>
                  </div>
                  <h3 className="font-bold truncate">{s.title}</h3>
                  <p className="text-sm text-gray-500 line-clamp-2">{s.body}</p>
                </div>
              );
          })}
        </div>
      </div>
    );
  };

  const renderEditor = () => (
    <div className="flex flex-col h-full bg-white">
      <div className="p-4 border-b flex justify-between items-center">
        <button onClick={() => setView('LIST')}><ArrowLeft className="w-5 h-5 text-gray-500" /></button>
        <h2 className="font-bold">{view === 'CREATE' ? 'Neu' : 'Bearbeiten'}</h2>
        <Button onClick={handleSave} disabled={!editorData.title}>Speichern</Button>
      </div>
      <div className="p-4 space-y-4 flex-1 overflow-y-auto">
        <div className="bg-blue-50 p-3 rounded border border-blue-100">
           <div className="flex gap-2 mb-2">
             <input id="ai-prompt" className="flex-1 text-sm border rounded px-2" placeholder="KI-Prompt: z.B. Höfliche Absage..." onKeyDown={e => e.key === 'Enter' && handleAiGenerate(e.currentTarget.value)} />
             <Button variant="secondary" onClick={() => handleAiGenerate((document.getElementById('ai-prompt') as HTMLInputElement).value)} disabled={isGenerating}><Sparkles className="w-4 h-4" /></Button>
           </div>
        </div>
        <input className="w-full border p-2 rounded text-sm" placeholder="Titel" value={editorData.title} onChange={e => setEditorData({...editorData, title: e.target.value})} />
        <select className="w-full border p-2 rounded text-sm" value={editorData.groupId} onChange={e => setEditorData({...editorData, groupId: e.target.value})}>
            {groups.map(g => <option key={g.id} value={g.id}>{g.name}</option>)}
        </select>
        <input className="w-full border p-2 rounded text-sm" placeholder="Betreff" value={editorData.subject} onChange={e => setEditorData({...editorData, subject: e.target.value})} />
        <textarea className="w-full border p-2 rounded text-sm h-40 font-mono" placeholder="Inhalt..." value={editorData.body} onChange={e => setEditorData({...editorData, body: e.target.value})} />
      </div>
    </div>
  );

  const renderFillVars = () => (
    <div className="flex flex-col h-full bg-white">
      <div className="p-4 border-b flex items-center"><button onClick={() => setView('LIST')}><ArrowLeft className="w-5 h-5 mr-2" /></button><h2 className="font-bold">Ausfüllen</h2></div>
      <div className="p-4 space-y-3 flex-1">
        {currentSnippet?.variables.map(v => (
            <div key={v}>
                <label className="text-xs font-bold uppercase text-gray-500">{v}</label>
                <input className="w-full border p-2 rounded" placeholder={v} value={variableValues[v] || ''} onChange={e => setVariableValues({...variableValues, [v]: e.target.value})} />
            </div>
        ))}
      </div>
      <div className="p-4 border-t flex justify-end gap-2">
         <Button variant="secondary" onClick={() => setView('LIST')}>Abbrechen</Button>
         <Button onClick={() => currentSnippet && executeInsert(currentSnippet, variableValues)} icon={Send}>Einfügen</Button>
      </div>
    </div>
  );

  const renderSettings = () => (
    <div className="h-full bg-white flex flex-col">
       <div className="p-4 border-b flex items-center"><button onClick={() => setView('LIST')}><ArrowLeft className="w-5 h-5 mr-2" /></button><h2 className="font-bold">Einstellungen</h2></div>
       <div className="p-4 space-y-6 overflow-y-auto">
          <div className={`p-3 rounded flex gap-3 ${isOfficeInitialized ? 'bg-green-50 text-green-800' : 'bg-yellow-50 text-yellow-800'}`}>
              {isOfficeInitialized ? <Link className="w-5 h-5" /> : <Link2Off className="w-5 h-5" />}
              <div>
                  <div className="font-bold text-sm">{isOfficeInitialized ? 'Outlook Verbunden' : 'Browser Modus'}</div>
              </div>
          </div>
          <div>
              <div className="flex justify-between mb-2">
                  <h3 className="font-bold text-sm">Konten</h3>
                  <button onClick={handleSyncAccounts} disabled={isSyncing} className="text-blue-600 text-xs flex items-center"><RefreshCw className={`w-3 h-3 mr-1 ${isSyncing ? 'animate-spin': ''}`} /> Sync</button>
              </div>
              <div className="border rounded bg-gray-50">
                  {accounts.map(a => <div key={a.id} className="p-2 border-b last:border-0 text-sm">{a.name} <span className="text-gray-500 text-xs">({a.email})</span></div>)}
              </div>
          </div>
          <div>
              <h3 className="font-bold text-sm mb-2">Kategorien</h3>
              <div className="flex gap-2 mb-2">
                  <input className="border rounded p-1 text-sm flex-1" placeholder="Name" value={newGroupName} onChange={e => setNewGroupName(e.target.value)} />
                  <div className="flex gap-1 items-center">
                    {AVAILABLE_COLORS.slice(0,4).map(c => <button key={c.class} onClick={() => setNewGroupColor(c.class)} className={`w-4 h-4 rounded-full ${c.class} ${newGroupColor === c.class ? 'ring-2 ring-black':''}`} />)}
                  </div>
                  <Button onClick={handleSaveGroup} className="px-2 py-1"><Plus className="w-4 h-4" /></Button>
              </div>
              <div className="space-y-1">
                  {groups.map(g => (
                      <div key={g.id} className="flex justify-between items-center p-2 border rounded bg-white text-sm">
                          <div className="flex items-center gap-2"><div className={`w-3 h-3 rounded-full ${g.color}`} /> {g.name}</div>
                          <div className="flex gap-1">
                              <button onClick={() => {setEditingGroupId(g.id); setNewGroupName(g.name); setNewGroupColor(g.color);}}><Edit3 className="w-3 h-3 text-gray-400 hover:text-blue-500" /></button>
                              <button onClick={() => setGroups(groups.filter(x => x.id !== g.id))}><Trash2 className="w-3 h-3 text-gray-400 hover:text-red-500" /></button>
                          </div>
                      </div>
                  ))}
              </div>
          </div>
       </div>
    </div>
  );

  const renderInfo = () => (
      <div className="h-full bg-white flex flex-col p-4">
          <div className="flex items-center mb-4"><button onClick={() => setView('LIST')}><ArrowLeft className="w-5 h-5 mr-2" /></button><h2 className="font-bold">Info</h2></div>
          <div className="text-sm space-y-4">
              <p>Mnemo Snippets v1.0 (Dev)</p>
              <p>Status: {isOfficeInitialized ? 'In Outlook' : 'Browser Preview'}</p>
          </div>
      </div>
  );

  return (
    <div className="flex w-full h-screen bg-white text-gray-900 overflow-hidden">
      {renderSidebar()}
      <div className="flex-1 flex flex-col relative">
        {view === 'LIST' && renderSnippetList()}
        {(view === 'CREATE' || view === 'EDIT') && renderEditor()}
        {view === 'FILL_VARS' && renderFillVars()}
        {view === 'SETTINGS' && renderSettings()}
        {view === 'INFO' && renderInfo()}
      </div>
    </div>
  );
}

const rootElement = document.getElementById('root');
if (rootElement) {
  const root = ReactDOM.createRoot(rootElement);
  root.render(<App />);
}