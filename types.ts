export interface Snippet {
  id: string;
  title: string;
  subject: string;
  body: string; // Supports markdown or HTML-like structure
  groupId: string;
  variables: string[]; // e.g. ['name', 'date', 'invoice_id']
}

export interface Group {
  id: string;
  name: string;
  color: string;
}

export interface SenderAccount {
  id: string;
  email: string;
  name: string;
  signature: string;
}

export type ViewState = 'LIST' | 'CREATE' | 'EDIT' | 'FILL_VARS' | 'INFO' | 'SETTINGS';

export interface SnippetFormData {
  title: string;
  subject: string;
  body: string;
  groupId: string;
}