export interface SearchResult {
  id: string;
  fileId: string;
  fileName: string;
  filePath: string;
  sheetName?: string;
  rowNumber: number;
  searchString: string;
  score?: number;
}

export interface FileMetadata {
  id: string;
  name: string;
  path: string;
  type: string;
  blob: File;
  size: number;
}

export type SortKey = 'relevance' | 'fileName' | 'rowNumber';
export type SortDir = 'asc' | 'desc';
