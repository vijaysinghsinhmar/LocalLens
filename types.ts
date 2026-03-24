
export interface SearchResult {
  id: string;
  fileId: string;
  fileName: string;
  filePath: string;
  sheetName?: string;
  rowNumber: number;
  searchString: string; // Limited preview for the list
}

export interface RecentSearch {
  query: string;
  timestamp: number;
  exactMatch: boolean;
  selectedFileTypes: string[];
  filterMode: 'include' | 'exclude';
}

export interface SortConfig {
  key: string;
  direction: 'asc' | 'desc';
}
