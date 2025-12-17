import React, { useState, useEffect, useMemo } from 'react';
import { 
  Layout, 
  Plus, 
  Settings, 
  Mail, 
  Search, 
  Folder, 
  Edit3, 
  Trash2, 
  Send, 
  ArrowLeft,
  Sparkles,
  Copy,
  CheckCircle,
  Info,
  RefreshCw,
  X,
  Save,
  Link,
  Link2Off
} from 'lucide-react';
import { Snippet, Group, ViewState, SnippetFormData, SenderAccount } from './types';
import { generateSnippet, improveText } from './services/geminiService';

// --- Constants & Mock Data ---

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
  },
  {
    id: 's2',
    groupId: 'g2',
    title: 'Angebot Nachfassen',
    subject: 'Rückfrage zu Angebot #{Angebotsnummer}',
    body: 'Sehr geehrte(r) {Name},\n\nhatten Sie bereits Gelegenheit, unser Angebot #{Angebotsnummer} vom {Datum} zu prüfen?\n\nBei Fragen stehe ich gerne zur Verfügung.',
    variables: ['Name', 'Angebotsnummer', 'Datum']
  }
];

// --- Helper Functions ---

const extractVariables = (text: string): string[] => {
  const regex = /\{([^}]+)\}/g;
  const matches = new Set<string>();
  let match;
  while ((match = regex.exec(text)) !== null) {
    matches.add(match[1]);
  }
  return Array.from(matches);
};

// --- Components ---

const Button = ({ 
  children, 
  onClick, 
  variant = 'primary', 
  className = '', 
  icon: Icon,
  disabled = false
}: { 
  children?: React.ReactNode, 
  onClick?: () => void, 
  variant?: 'primary' | 'secondary' | 'ghost' | 'danger',
  className?: string,
  icon?: React.ElementType,
  disabled?: boolean
}) => {
  const baseStyle = "flex items-center justify-center px-4 py-2 rounded-md font-medium transition-colors text-sm focus:outline-none focus:ring-2 focus:ring-offset-1";
  const variants = {
    primary: "bg-blue-600 text-white hover:bg-blue-700 focus:ring-blue-500 disabled:bg-blue-300",
    secondary: "bg-white text-gray-700 border border-gray-300 hover:bg-gray-50 focus:ring-gray-200 disabled:bg-gray-100",
    ghost: "bg-transparent text-gray-600 hover:bg-gray-100 focus:ring-gray-200",
    danger: "bg-red-50 text-red-600 hover:bg-red-100 focus:ring-red-200"
  };

  return (
    <button 
      onClick={onClick} 
      className={`${baseStyle} ${variants[variant]} ${className}`}
      disabled={disabled}
    >
      {Icon && <Icon className="w-4 h-4 mr-2" />}
      {children}
    </button>
  );
};

export default function App() {
  const [view, setView] = useState<ViewState>('LIST');
  const [snippets, setSnippets] = useState<Snippet[]>(INITIAL_SNIPPETS);
  const [groups, setGroups] = useState<Group[]>(INITIAL_GROUPS);
  const [accounts, setAccounts] = useState<SenderAccount[]>(INITIAL_ACCOUNTS);
  
  const [selectedGroup, setSelectedGroup] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState('');
  const [currentSnippet, setCurrentSnippet] = useState<Snippet | null>(null);
  const [variableValues, setVariableValues] = useState<Record<string, string>>({});
  const [selectedAccount, setSelectedAccount] = useState<string>(INITIAL_ACCOUNTS[0]?.id || '');
  
  // Settings State
  const [newGroupName, setNewGroupName] = useState('');
  const [newGroupColor, setNewGroupColor] = useState(AVAILABLE_COLORS[0].class);
  const [editingGroupId, setEditingGroupId] = useState<string | null>(null);
  const [isSyncing, setIsSyncing] = useState(false);
  
  // Office Integration State
  const [isOfficeInitialized, setIsOfficeInitialized] = useState(false);

  // Editor State
  const [editorData, setEditorData] = useState<SnippetFormData>({
    title: '',
    subject: '',
    body: '',
    groupId: 'g1'
  });
  const [isGenerating, setIsGenerating] = useState(false);

  // Initial Data Load & Office Init
  useEffect(() => {
    // 1. Initialize Office.js
    const initOffice = async () => {
        if (typeof window !== 'undefined' && (window as any).Office) {
            try {
                await (window as any).Office.onReady((info: any) => {
                    console.log("Office.js is ready", info);
                    if (info.host) {
                        setIsOfficeInitialized(true);
                    }
                });
            } catch (e) {
                console.error("Office.onReady failed", e);
            }
        }
    };
    initOffice();

    // 2. Auto-Sync Accounts (Mock)
    handleSyncAccounts();
  }, []);

  // Ensure selectedAccount is valid
  useEffect(() => {
    if (accounts.length > 0 && !accounts.find(a => a.id === selectedAccount)) {
      setSelectedAccount(accounts[0].id);
    }
  }, [accounts, selectedAccount]);

  // --- Actions ---

  const handleCreate = () => {
    setEditorData({ title: '', subject: '', body: '', groupId: groups[0]?.id || 'g1' });
    setCurrentSnippet(null);
    setView('CREATE');
  };

  const handleEdit = (snippet: Snippet) => {
    setEditorData({
      title: snippet.title,
      subject: snippet.subject,
      body: snippet.body,
      groupId: snippet.groupId
    });
    setCurrentSnippet(snippet);
    setView('EDIT');
  };

  const handleDelete = (id: string) => {
    if (confirm('Möchten Sie diesen Baustein wirklich löschen?')) {
      setSnippets(prev => prev.filter(s => s.id !== id));
      if (view === 'FILL_VARS') setView('LIST');
    }
  };

  const handleSave = () => {
    const variables = extractVariables(editorData.subject + ' ' + editorData.body);
    
    if (view === 'CREATE') {
      const newSnippet: Snippet = {
        id: Date.now().toString(),
        ...editorData,
        variables
      };
      setSnippets([...snippets, newSnippet]);
    } else if (view === 'EDIT' && currentSnippet) {
      setSnippets(snippets.map(s => s.id === currentSnippet.id ? { ...s, ...editorData, variables } : s));
    }
    setView('LIST');
  };

  const handleAiGenerate = async (prompt: string) => {
    setIsGenerating(true);
    try {
      const result = await generateSnippet(prompt);
      setEditorData(prev => ({
        ...prev,
        title: result.title || prev.title,
        subject: result.subject || prev.subject,
        body: result.body || prev.body
      }));
    } catch (e) {
      alert("Fehler bei der KI-Generierung");
    } finally {
      setIsGenerating(false);
    }
  };

  // Group Management Actions
  const handleSaveGroup = () => {
    if (!newGroupName.trim()) return;

    if (editingGroupId) {
        // Edit existing
        setGroups(prev => prev.map(g => g.id === editingGroupId ? { ...g, name: newGroupName, color: newGroupColor } : g));
        setEditingGroupId(null);
    } else {
        // Add new
        const newGroup: Group = {
            id: `g${Date.now()}`,
            name: newGroupName,
            color: newGroupColor
        };
        setGroups([...groups, newGroup]);
    }
    setNewGroupName('');
    setNewGroupColor(AVAILABLE_COLORS[0].class);
  };

  const handleEditGroup = (group: Group) => {
    setEditingGroupId(group.id);
    setNewGroupName(group.name);
    setNewGroupColor(group.color);
  };

  const handleCancelEditGroup = () => {
    setEditingGroupId(null);
    setNewGroupName('');
    setNewGroupColor(AVAILABLE_COLORS[0].class);
  };

  const handleDeleteGroup = (id: string) => {
    if (confirm('Möchten Sie diese Kategorie löschen? Bausteine in dieser Kategorie bleiben erhalten, verlieren aber ihre Zuordnung.')) {
      setGroups(groups.filter(g => g.id !== id));
      if (selectedGroup === id) setSelectedGroup(null);
      // Reset edit state if we deleted the group being edited
      if (editingGroupId === id) handleCancelEditGroup();
    }
  };

  // Outlook Sync Simulation
  const handleSyncAccounts = async () => {
    setIsSyncing(true);
    // Simulate Office.context.mailbox.userProfile calls
    setTimeout(() => {
      const newAccounts: SenderAccount[] = [
        ...INITIAL_ACCOUNTS,
        { 
          id: 'acc_synced_1', 
          name: 'Team Shared Mailbox', 
          email: 'team@firma.de', 
          signature: '\n\nViele Grüße\nIhr Team' 
        },
        { 
          id: 'acc_synced_2', 
          name: 'Max Privat', 
          email: 'max.p@firma.de', 
          signature: '\n\nSent from Outlook Mobile' 
        }
      ];
      setAccounts(newAccounts);
      setIsSyncing(false);
      // Only alert if triggered manually to not annoy on auto-load
      if (view === 'SETTINGS') {
         // alert('Konten synchronisiert!'); // Optional feedback
      }
    }, 1000);
  };

  const handlePrepareInsert = (snippet: Snippet) => {
    setCurrentSnippet(snippet);
    setVariableValues({});
    if (snippet.variables.length > 0) {
      setView('FILL_VARS');
    } else {
      executeInsert(snippet, {});
    }
  };

  const executeInsert = async (snippet: Snippet, values: Record<string, string>) => {
    // 1. Replace variables
    let finalSubject = snippet.subject;
    let finalBody = snippet.body;

    Object.entries(values).forEach(([key, val]) => {
      const regex = new RegExp(`\\{${key}\\}`, 'g');
      finalSubject = finalSubject.replace(regex, val);
      finalBody = finalBody.replace(regex, val);
    });

    // 2. Append Signature
    const account = accounts.find(a => a.id === selectedAccount);
    if (account) {
      finalBody += account.signature;
    }

    // 3. Real Outlook Insert vs Simulation
    if (isOfficeInitialized) {
        // REAL OUTLOOK INSERT
        try {
            const office = (window as any).Office;
            office.context.mailbox.item.body.setSelectedDataAsync(
                finalBody, 
                { coercionType: office.CoercionType.Html }, 
                (asyncResult: any) => {
                    if (asyncResult.status === office.AsyncResultStatus.Failed) {
                         alert('Fehler beim Einfügen: ' + asyncResult.error.message);
                    } else {
                        // Also try to set Subject if possible (Compose mode)
                         office.context.mailbox.item.subject.setAsync(finalSubject, (res: any) => {
                             // Ignore errors here if not in Compose mode
                         });
                         setView('LIST');
                    }
                }
            );
        } catch (e) {
            console.error(e);
            alert("Fehler bei der Kommunikation mit Outlook.");
        }
    } else {
        // SIMULATION (Browser)
        console.log("INSERTING INTO OUTLOOK:", { subject: finalSubject, body: finalBody, account: account?.email });
        
        // Fallback for Web Demo: Copy to clipboard
        const fullText = `[Von: ${account?.email}]\nBetreff: ${finalSubject}\n\n${finalBody}`;
        navigator.clipboard.writeText(fullText);
        alert(`In Outlook eingefügt (simuliert):\n\n${fullText}\n\n(Text wurde in die Zwischenablage kopiert)`);
        
        setView('LIST');
    }
  };

  // --- Views ---

  const renderSidebar = () => (
    <div className="w-16 bg-white border-r border-gray-200 flex flex-col items-center py-4 space-y-4">
      <button 
        onClick={() => setView('LIST')}
        className={`p-2 rounded-xl transition-all ${view === 'LIST' ? 'bg-blue-100 text-blue-600' : 'text-gray-400 hover:bg-gray-100'}`}
        title="Bausteine"
      >
        <Layout className="w-6 h-6" />
      </button>
      <button 
        onClick={handleCreate}
        className={`p-2 rounded-xl transition-all ${view === 'CREATE' ? 'bg-blue-100 text-blue-600' : 'text-gray-400 hover:bg-gray-100'}`}
        title="Neu erstellen"
      >
        <Plus className="w-6 h-6" />
      </button>
      <div className="flex-grow" />
      <button 
         onClick={() => setView('SETTINGS')}
         className={`p-2 rounded-xl transition-all ${view === 'SETTINGS' ? 'bg-blue-100 text-blue-600' : 'text-gray-400 hover:bg-gray-100'}`}
         title="Einstellungen"
      >
        <Settings className="w-6 h-6" />
      </button>
      <button 
         onClick={() => setView('INFO')}
         className={`p-2 rounded-xl transition-all ${view === 'INFO' ? 'bg-blue-100 text-blue-600' : 'text-gray-400 hover:bg-gray-100'}`}
         title="Entwickler Infos"
      >
        <Info className="w-6 h-6" />
      </button>
    </div>
  );

  const renderSettings = () => (
    <div className="flex flex-col h-full bg-white overflow-y-auto">
      <div className="p-4 border-b border-gray-200 flex items-center bg-white sticky top-0 z-10">
        <button onClick={() => setView('LIST')} className="text-gray-500 hover:text-gray-700 mr-3">
          <ArrowLeft className="w-5 h-5" />
        </button>
        <h2 className="font-bold text-lg">Einstellungen</h2>
      </div>

      <div className="p-6 space-y-8">
        {/* Status Indicator */}
        <div className={`p-3 rounded-lg flex items-center gap-3 ${isOfficeInitialized ? 'bg-green-50 text-green-800 border border-green-200' : 'bg-yellow-50 text-yellow-800 border border-yellow-200'}`}>
            {isOfficeInitialized ? <Link className="w-5 h-5" /> : <Link2Off className="w-5 h-5" />}
            <div className="flex-1">
                <div className="font-bold text-sm">{isOfficeInitialized ? 'Mit Outlook verbunden' : 'Preview Modus (Browser)'}</div>
                <div className="text-xs opacity-80">{isOfficeInitialized ? 'Alle Funktionen aktiv' : 'Text-Einfügen wird nur simuliert'}</div>
            </div>
        </div>

        {/* Account Settings */}
        <section>
          <div className="flex items-center justify-between mb-4">
            <h3 className="font-bold text-gray-900 flex items-center gap-2">
              <Mail className="w-4 h-4" /> Outlook Konten
            </h3>
            <Button variant="secondary" onClick={handleSyncAccounts} disabled={isSyncing}>
              {isSyncing ? <RefreshCw className="w-4 h-4 animate-spin" /> : <RefreshCw className="w-4 h-4" />}
              <span className="ml-2">Synchronisieren</span>
            </Button>
          </div>
          <div className="bg-gray-50 rounded-lg border border-gray-200 overflow-hidden">
             {accounts.map((acc, idx) => (
               <div key={acc.id} className={`p-3 flex items-center justify-between ${idx !== accounts.length - 1 ? 'border-b border-gray-200' : ''}`}>
                 <div>
                   <div className="font-medium text-sm text-gray-900">{acc.name}</div>
                   <div className="text-xs text-gray-500">{acc.email}</div>
                 </div>
                 <div className="px-2 py-1 bg-green-100 text-green-700 text-[10px] rounded-full font-bold">VERBUNDEN</div>
               </div>
             ))}
             {accounts.length === 0 && <div className="p-4 text-center text-gray-500 text-sm">Keine Konten gefunden.</div>}
          </div>
        </section>

        {/* Group Management */}
        <section>
          <h3 className="font-bold text-gray-900 mb-4 flex items-center gap-2">
            <Folder className="w-4 h-4" /> Kategorien & Tags verwalten
          </h3>
          
          {/* Add/Edit Group Form */}
          <div className={`flex gap-2 mb-4 bg-gray-50 p-3 rounded-lg border items-end relative transition-colors ${editingGroupId ? 'border-blue-300 bg-blue-50' : 'border-gray-200'}`}>
            {editingGroupId && (
                <div className="absolute -top-2 -right-2">
                    <button onClick={handleCancelEditGroup} className="bg-white border border-gray-200 rounded-full p-1 text-gray-400 hover:text-gray-600 shadow-sm">
                        <X className="w-3 h-3" />
                    </button>
                </div>
            )}
            <div className="flex-1">
              <label className="block text-xs font-medium text-gray-500 mb-1">{editingGroupId ? 'Kategorie bearbeiten' : 'Neue Kategorie'}</label>
              <input 
                value={newGroupName}
                onChange={(e) => setNewGroupName(e.target.value)}
                className="w-full text-sm border-gray-300 rounded focus:ring-blue-500 focus:border-blue-500 px-2 py-1.5 border"
                placeholder="Bezeichnung..."
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-gray-500 mb-1">Farbe</label>
              <div className="flex gap-1">
                {AVAILABLE_COLORS.map(c => (
                  <button
                    key={c.class}
                    onClick={() => setNewGroupColor(c.class)}
                    className={`w-6 h-6 rounded-full ${c.class} transition-transform ${newGroupColor === c.class ? 'ring-2 ring-offset-1 ring-gray-400 scale-110' : 'hover:scale-105'}`}
                    title={c.name}
                  />
                ))}
              </div>
            </div>
            <Button onClick={handleSaveGroup} disabled={!newGroupName} className="h-9">
                {editingGroupId ? <Save className="w-4 h-4" /> : <Plus className="w-4 h-4" />}
            </Button>
          </div>

          {/* Group List */}
          <div className="grid grid-cols-1 gap-2">
            {groups.map(g => (
              <div key={g.id} className={`flex items-center justify-between bg-white border border-gray-200 p-2 rounded-md hover:shadow-sm ${editingGroupId === g.id ? 'ring-1 ring-blue-500 bg-blue-50' : ''}`}>
                <div className="flex items-center gap-3">
                  <div className={`w-3 h-3 rounded-full ${g.color}`} />
                  <span className="text-sm font-medium text-gray-700">{g.name}</span>
                </div>
                <div className="flex items-center space-x-1">
                  <button 
                    onClick={() => handleEditGroup(g)} 
                    className="text-gray-400 hover:text-blue-500 p-1 rounded hover:bg-blue-50"
                    title="Bearbeiten"
                  >
                    <Edit3 className="w-4 h-4" />
                  </button>
                  <button 
                    onClick={() => handleDeleteGroup(g.id)} 
                    className="text-gray-400 hover:text-red-500 p-1 rounded hover:bg-red-50"
                    title="Löschen"
                  >
                    <Trash2 className="w-4 h-4" />
                  </button>
                </div>
              </div>
            ))}
            {groups.length === 0 && <p className="text-sm text-gray-500 italic">Keine Kategorien vorhanden.</p>}
          </div>
        </section>
      </div>
    </div>
  );

  const renderSnippetList = () => {
    const filtered = snippets.filter(s => {
      const matchesGroup = selectedGroup ? s.groupId === selectedGroup : true;
      const matchesSearch = s.title.toLowerCase().includes(searchQuery.toLowerCase()) || 
                            s.subject.toLowerCase().includes(searchQuery.toLowerCase());
      return matchesGroup && matchesSearch;
    });

    return (
      <div className="flex flex-col h-full bg-gray-50">
        {/* Top Bar */}
        <div className="bg-white p-4 border-b border-gray-200 space-y-3 sticky top-0 z-10">
          <div className="flex items-center justify-between">
            <h1 className="text-xl font-bold text-gray-800">Bibliothek</h1>
             <div className="flex items-center space-x-2">
                <Button onClick={handleCreate} icon={Plus} className="h-8 text-xs">
                  Neu
                </Button>
                <select 
                  className="text-sm border-gray-300 border rounded p-1 h-8 text-gray-600 bg-gray-50 focus:ring-blue-500 focus:border-blue-500 max-w-[150px]"
                  value={selectedAccount}
                  onChange={(e) => setSelectedAccount(e.target.value)}
                >
                  {accounts.map(acc => (
                    <option key={acc.id} value={acc.id}>{acc.name}</option>
                  ))}
                </select>
            </div>
          </div>
          
          <div className="relative">
            <Search className="absolute left-3 top-2.5 w-4 h-4 text-gray-400" />
            <input 
              type="text" 
              placeholder="Suchen..." 
              className="w-full pl-9 pr-4 py-2 bg-gray-100 border-none rounded-md text-sm focus:ring-2 focus:ring-blue-500"
              value={searchQuery}
              onChange={e => setSearchQuery(e.target.value)}
            />
          </div>

          <div className="flex items-center space-x-2">
            <div className="flex-1 flex space-x-2 overflow-x-auto pb-1 no-scrollbar">
              <button 
                onClick={() => setSelectedGroup(null)}
                className={`px-3 py-1 rounded-full text-xs font-medium whitespace-nowrap transition-colors ${!selectedGroup ? 'bg-gray-800 text-white' : 'bg-gray-200 text-gray-600 hover:bg-gray-300'}`}
              >
                Alle
              </button>
              {groups.map(g => (
                <button 
                  key={g.id}
                  onClick={() => setSelectedGroup(g.id)}
                  className={`px-3 py-1 rounded-full text-xs font-medium whitespace-nowrap transition-colors flex items-center ${selectedGroup === g.id ? 'bg-blue-600 text-white' : 'bg-white border border-gray-200 text-gray-600 hover:bg-gray-50'}`}
                >
                  <div className={`w-2 h-2 rounded-full mr-2 ${g.color}`} />
                  {g.name}
                </button>
              ))}
            </div>
            <div className="w-px h-5 bg-gray-300 mx-1 flex-shrink-0" />
            <button 
              onClick={() => setView('SETTINGS')}
              className="p-1.5 text-gray-500 hover:text-gray-700 bg-gray-50 hover:bg-gray-100 border border-gray-200 rounded-md transition-colors flex-shrink-0"
              title="Tags verwalten"
            >
              <Settings className="w-4 h-4" />
            </button>
          </div>
        </div>

        {/* List Content */}
        <div className="flex-1 overflow-y-auto p-4 space-y-3">
          {filtered.length === 0 ? (
            <div className="text-center text-gray-500 mt-10">
              <p>Keine Bausteine gefunden.</p>
              <Button variant="ghost" className="mt-2" onClick={handleCreate}>Jetzt erstellen</Button>
            </div>
          ) : (
            filtered.map(snippet => {
              const group = groups.find(g => g.id === snippet.groupId);
              return (
                <div key={snippet.id} className="bg-white rounded-lg shadow-sm border border-gray-200 hover:shadow-md transition-shadow group relative">
                  <div className="p-4 cursor-pointer" onClick={() => handlePrepareInsert(snippet)}>
                    <div className="flex justify-between items-start mb-1">
                      <span className={`text-[10px] font-bold px-2 py-0.5 rounded uppercase tracking-wide text-white ${group?.color || 'bg-gray-400'}`}>
                        {group?.name || 'Allgemein'}
                      </span>
                      <div className="flex space-x-1 opacity-0 group-hover:opacity-100 transition-opacity">
                         <button 
                          onClick={(e) => { e.stopPropagation(); handleEdit(snippet); }}
                          className="p-1.5 text-gray-400 hover:text-blue-600 hover:bg-blue-50 rounded"
                        >
                          <Edit3 className="w-3.5 h-3.5" />
                        </button>
                         <button 
                          onClick={(e) => { e.stopPropagation(); handleDelete(snippet.id); }}
                          className="p-1.5 text-gray-400 hover:text-red-600 hover:bg-red-50 rounded"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </div>
                    <h3 className="font-semibold text-gray-900 mb-1 truncate">{snippet.title}</h3>
                    <p className="text-sm text-gray-500 line-clamp-2">{snippet.body}</p>
                    
                    {snippet.variables.length > 0 && (
                      <div className="mt-3 flex flex-wrap gap-1">
                        {snippet.variables.map(v => (
                          <span key={v} className="bg-yellow-50 text-yellow-700 border border-yellow-200 px-1.5 py-0.5 rounded text-[10px] font-mono">
                            {'{' + v + '}'}
                          </span>
                        ))}
                      </div>
                    )}
                  </div>
                </div>
              );
            })
          )}
        </div>
      </div>
    );
  };

  const renderEditor = () => (
    <div className="flex flex-col h-full bg-white">
      <div className="p-4 border-b border-gray-200 flex items-center justify-between bg-white sticky top-0 z-20">
        <button onClick={() => setView('LIST')} className="text-gray-500 hover:text-gray-700">
          <ArrowLeft className="w-5 h-5" />
        </button>
        <h2 className="font-bold text-lg">{view === 'CREATE' ? 'Neuer Baustein' : 'Baustein bearbeiten'}</h2>
        <Button onClick={handleSave} disabled={!editorData.title}>Speichern</Button>
      </div>

      <div className="flex-1 overflow-y-auto p-4 space-y-5">
        
        {/* AI Assistant Section */}
        <div className="bg-gradient-to-r from-blue-50 to-indigo-50 p-4 rounded-lg border border-blue-100 shadow-sm">
          <div className="flex items-center space-x-2 text-indigo-700 mb-2 font-semibold">
            <Sparkles className="w-4 h-4" />
            <span className="text-sm">KI-Assistent</span>
          </div>
          <p className="text-xs text-indigo-600 mb-3">Beschreibe den gewünschten Baustein und lass ihn generieren.</p>
          <div className="flex gap-2">
            <input 
              type="text" 
              placeholder="z.B. Höfliche Absage für eine Bewerbung..."
              className="flex-1 text-sm border-gray-300 rounded focus:ring-indigo-500 focus:border-indigo-500 px-3 py-2"
              onKeyDown={(e) => {
                if (e.key === 'Enter') handleAiGenerate(e.currentTarget.value);
              }}
              id="ai-prompt-input"
            />
            <Button 
              variant="secondary" 
              onClick={() => {
                const input = document.getElementById('ai-prompt-input') as HTMLInputElement;
                handleAiGenerate(input.value);
              }}
              disabled={isGenerating}
            >
              {isGenerating ? '...' : 'Generieren'}
            </Button>
          </div>
        </div>

        {/* Standard Form */}
        <div className="space-y-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Titel (Intern)</label>
            <input 
              className="w-full border-gray-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500 text-sm p-2 border"
              value={editorData.title}
              onChange={e => setEditorData({ ...editorData, title: e.target.value })}
              placeholder="z.B. Bewerbung Absage"
            />
          </div>

          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Kategorie</label>
            <div className="flex gap-2">
              <select 
                className="w-full border-gray-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500 text-sm p-2 border"
                value={editorData.groupId}
                onChange={e => setEditorData({ ...editorData, groupId: e.target.value })}
              >
                {groups.map(g => (
                  <option key={g.id} value={g.id}>{g.name}</option>
                ))}
              </select>
              <button 
                onClick={() => setView('SETTINGS')} 
                className="px-3 py-1 bg-gray-100 border border-gray-300 rounded text-gray-600 hover:bg-gray-200"
                title="Kategorien verwalten"
              >
                <Settings className="w-4 h-4" />
              </button>
            </div>
          </div>

          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Betreff</label>
            <input 
              className="w-full border-gray-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500 text-sm p-2 border"
              value={editorData.subject}
              onChange={e => setEditorData({ ...editorData, subject: e.target.value })}
              placeholder="Betreff der E-Mail"
            />
          </div>

          <div className="flex-1 flex flex-col">
            <label className="block text-sm font-medium text-gray-700 mb-1">Inhalt</label>
            <textarea 
              className="w-full flex-1 min-h-[200px] border-gray-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500 text-sm p-2 border font-mono"
              value={editorData.body}
              onChange={e => setEditorData({ ...editorData, body: e.target.value })}
              placeholder="Nutze {Variable} für Platzhalter..."
            />
            <p className="text-xs text-gray-400 mt-1">Variablen werden automatisch erkannt (z.B. {'{Name}'}).</p>
          </div>
        </div>
      </div>
    </div>
  );

  const renderFillVars = () => {
    if (!currentSnippet) return null;
    return (
      <div className="flex flex-col h-full bg-white">
         <div className="p-4 border-b border-gray-200 flex items-center bg-white sticky top-0">
          <button onClick={() => setView('LIST')} className="text-gray-500 hover:text-gray-700 mr-3">
            <ArrowLeft className="w-5 h-5" />
          </button>
          <h2 className="font-bold text-gray-800">Variablen ausfüllen</h2>
        </div>
        
        <div className="p-6 flex-1 overflow-y-auto">
          <div className="mb-6 bg-gray-50 p-4 rounded border border-gray-200">
            <h3 className="text-sm font-semibold text-gray-700 mb-1">{currentSnippet.title}</h3>
            <p className="text-xs text-gray-500">Dieser Baustein enthält Platzhalter, die vor dem Einfügen ausgefüllt werden müssen.</p>
          </div>

          <div className="space-y-4">
            {currentSnippet.variables.map(v => (
              <div key={v}>
                <label className="block text-sm font-medium text-gray-700 mb-1 capitalize">{v}</label>
                <input 
                  autoFocus={v === currentSnippet.variables[0]}
                  className="w-full border-gray-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500 p-2 border"
                  value={variableValues[v] || ''}
                  onChange={e => setVariableValues({ ...variableValues, [v]: e.target.value })}
                  placeholder={`Wert für ${v}...`}
                />
              </div>
            ))}
          </div>
        </div>

        <div className="p-4 border-t border-gray-200 bg-gray-50 flex justify-end space-x-3">
          <Button variant="secondary" onClick={() => setView('LIST')}>Abbrechen</Button>
          <Button 
            onClick={() => executeInsert(currentSnippet, variableValues)} 
            icon={Send}
          >
            In Outlook einfügen
          </Button>
        </div>
      </div>
    );
  };

  const renderInfo = () => (
    <div className="flex flex-col h-full bg-white overflow-y-auto">
       <div className="p-4 border-b border-gray-200 flex items-center">
        <button onClick={() => setView('LIST')} className="text-gray-500 hover:text-gray-700 mr-3">
          <ArrowLeft className="w-5 h-5" />
        </button>
        <h2 className="font-bold text-gray-800">Entwickler Infos</h2>
      </div>
      <div className="p-6 text-sm text-gray-700 space-y-6">
        <section>
          <h3 className="font-bold text-lg text-gray-900 mb-2">Anforderungen für den Start</h3>
          <p className="mb-2">Da Sie gefragt haben, was ich wissen muss, um das Plugin final zu bauen – hier ist die Checkliste für das "echte" Outlook Add-in (Version 1.2025+):</p>
          <ul className="list-disc pl-5 space-y-1 text-gray-600">
            <li><strong>Manifest XML:</strong> Wir benötigen eine `manifest.xml` Datei, die dem Exchange Server sagt, wo diese Web-App gehostet wird.</li>
            <li><strong>Office.js:</strong> Die Integration erfolgt über die JavaScript API für Office. Dieser Prototyp nutzt Mocks.</li>
            <li><strong>Hosting:</strong> Diese React App muss auf einem Webserver (z.B. Vercel, Azure Static Web Apps) laufen und per HTTPS erreichbar sein.</li>
            <li><strong>Permissions:</strong> Das Manifest muss `ReadWriteItem` Berechtigungen anfordern, um in den Body schreiben zu dürfen.</li>
          </ul>
        </section>

        <section>
          <h3 className="font-bold text-lg text-gray-900 mb-2">Features im Prototyp</h3>
          <ul className="list-disc pl-5 space-y-1 text-gray-600">
            <li><strong>Textbausteine verwalten:</strong> Erstellen, Bearbeiten, Löschen, Gruppieren.</li>
            <li><strong>Variablen-Engine:</strong> Automatische Erkennung von {'{Platzhaltern}'} und Prompting beim Einfügen.</li>
            <li><strong>KI-Integration:</strong> Gemini API Anbindung zur Generierung von Texten.</li>
            <li><strong>Konto-Simulation:</strong> Auswahl der Absender-Identität (Signatur-Wechsel).</li>
          </ul>
        </section>

        <div className="bg-blue-50 p-4 rounded border border-blue-100">
          <h4 className="font-semibold text-blue-800 mb-1">Nächster Schritt</h4>
          <p className="text-blue-700">Um dies in Ihr Outlook (Version 1.2025...) zu bringen, müsste dieser Code als Web-App deployed und per Sideloading (Manifest-Upload) in Outlook installiert werden.</p>
        </div>
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