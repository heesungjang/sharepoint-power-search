export enum FilterComparisonOperator {
  Eq,
  Neq,
  Gt,
  Lt,
  Geq,
  Leq,
  Contains,
}

export interface IDataFilterValue {
  /**
   * The filter value display name
   */
  name: string;

  /**
   * Inner value to use when the value is selected
   */
  value: string;

  /**
   * The comparison operator to use with this value. If not provided, the 'Equals' operator will be used.
   */
  operator?: FilterComparisonOperator;
}

export interface IDataFilterResult {
  /**
   * The filter display name
   */
  filterName: string;

  /**
   * Values available in this filter
   */
  values: IDataFilterResultValue[];
}

export interface IDataFilterResultValue extends IDataFilterValue {
  /**
   * The number of results with this value
   */
  count: number;
}

export interface ISharePointSearchResult {
  [key: string]: string | undefined;

  // Known SharePoint managed properties
  Title?: string;
  Path?: string;
  FileType?: string;
  HitHighlightedSummary?: string;
  AuthorOWSUSER?: string;
  owstaxidmetadataalltagsinfo?: string;
  Created?: string;
  UniqueID?: string;
  NormSiteID?: string;
  NormWebID?: string;
  NormListID?: string;
  NormUniqueID?: string;
}

export interface ISharePointSearchResults {
  queryModification?: string;
  queryKeywords: string;
  relevantResults: ISharePointSearchResult[];
  secondaryResults: ISharePointSearchResultBlock[];
  refinementResults: IDataFilterResult[];
  promotedResults?: ISharePointSearchPromotedResult[];
  spellingSuggestion?: string;
  totalRows: number;
}

export interface ISharePointSearchPromotedResult {
  url: string;
  title: string;
  description: string;
}

export interface ISharePointSearchResultBlock {
  title: string;
  results: ISharePointSearchResult[];
}
