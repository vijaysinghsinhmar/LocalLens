export interface SearchResult {
  id: string;
  fileId: string;
  fileName: string;
  sheetName: string | null;
  rowNumber: number;
  searchString: string;
  score: number;
  _ri: number;
}

export interface FileEntry {
  id: string;
  name: string;
  path: string;
  type: string;
  blob: File;
  size: number;
  indexed: boolean;
  rowCount: number;
  error?: string;
}

export type SortKey = 'relevance' | 'fileName' | 'rowNumber';
export type SortDir = 'asc' | 'desc';
