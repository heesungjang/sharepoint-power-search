import {
  IDataFilterResult,
  IDataFilterResultValue,
  ISharePointSearchResult,
  ISharePointSearchResults,
  FilterComparisonOperator,
  ISharePointSearchResultBlock,
  ISharePointSearchPromotedResult,
} from '../models/sharepoint-search-results';

import { getSP } from '../../common/pnpjs';
import { isEmpty, flatMap } from 'lodash';
import { useInfiniteQuery } from '@tanstack/react-query';
import { ISharePointSearchQuery } from '../models/sharepoint-search-query';
import { SEARCH_PROPERTIES } from '../models/sharepoint-search-query';

const DEFAULT_SEARCH_CONFIG = {
  queryTemplate:
    '{searchTerms} XRANK(cb=3.0) title:"{searchTerms}*" XRANK(cb=2.75) title:"{searchTerms}" XRANK(cb=2.25) filename:"{searchTerms}*" XRANK(cb=2.0) filename:"{searchTerms}" XRANK(cb=1.8) FileLeafRef:"{searchTerms}*" XRANK(cb=1.5) path:"{searchTerms}" OR title:"*{searchTerms}*" OR filename:"*{searchTerms}*" OR FileLeafRef:"*{searchTerms}*"',
  hitHighlightedProperties: [
    'Title',
    'Path',
    'Author',
    'Filename',
    'FileLeafRef',
    'Description',
  ],
  sourceId: '8413CD39-2156-4E00-B54D-11EFD9ABDB89',
  trimDuplicates: false,
  summaryLength: 300,
  rowLimit: 50,
  enableQueryRules: true,
  enableInterleaving: true,
  processBestBets: true,
  processPersonalFavorites: true,
};

export const useSearch = (
  searchParams: {
    query: string;
    refinementFilters?: string[];
    startRow?: number;
    sortProperty?: string;
    sortDirection?: 'asc' | 'desc';
  },
  fileTypesKey: string,
  sortOrder: 'asc' | 'desc' | 'relevance',
  dependencies: any[],
  selectedScope: string,
  config: Partial<typeof DEFAULT_SEARCH_CONFIG> = {}
) => {
  const sp = getSP();

  const searchConfig = { ...DEFAULT_SEARCH_CONFIG, ...config };

  const buildSearchQuery = (pageParam: number): ISharePointSearchQuery => ({
    Querytext: searchParams.query,
    SelectProperties: SEARCH_PROPERTIES,
    TrimDuplicates: searchConfig.trimDuplicates,
    SummaryLength: searchConfig.summaryLength,
    RowLimit: searchConfig.rowLimit,
    StartRow: pageParam,
    RefinementFilters: searchParams.refinementFilters || [],
    SortList:
      sortOrder === 'relevance'
        ? []
        : [
            {
              Property: searchParams.sortProperty || 'LastModifiedTime',
              Direction: sortOrder === 'desc' ? 1 : 0,
            },
          ],
    EnableQueryRules: searchConfig.enableQueryRules,
    EnableInterleaving: searchConfig.enableInterleaving,
    ProcessBestBets: searchConfig.processBestBets,
    ProcessPersonalFavorites: searchConfig.processPersonalFavorites,
    QueryTemplate: searchConfig.queryTemplate,
    HitHighlightedProperties: searchConfig.hitHighlightedProperties,
    SourceId: searchConfig.sourceId,
  });

  return useInfiniteQuery<ISharePointSearchResults>({
    queryKey: [
      'search-result',
      searchParams.query,
      fileTypesKey,
      sortOrder,
      dependencies,
      selectedScope,
    ],
    queryFn: async ({ pageParam = 0 }) => {
      const currentQuery = buildSearchQuery(pageParam);

      let results: ISharePointSearchResults = {
        queryKeywords: searchParams.query || '',
        refinementResults: [],
        relevantResults: [],
        secondaryResults: [],
        totalRows: 0,
      };

      const searchResponse = await sp.search(currentQuery);

      if (searchResponse && searchResponse.RawSearchResults) {
        const rawSearchResults = searchResponse.RawSearchResults;

        if (rawSearchResults.PrimaryQueryResult) {
          let refinementResults: IDataFilterResult[] = [];

          const properties =
            rawSearchResults.PrimaryQueryResult.RelevantResults?.Properties?.filter(
              (property) => {
                return property.Key === 'QueryModification';
              }
            );

          if (properties && properties.length === 1) {
            results.queryModification = properties[0].Value;
          }

          const resultRows =
            rawSearchResults.PrimaryQueryResult.RelevantResults?.Table?.Rows;
          let refinementResultsRows =
            rawSearchResults.PrimaryQueryResult.RefinementResults;

          const refinementRows: any = refinementResultsRows
            ? refinementResultsRows.Refiners
            : [];

          let searchResults: ISharePointSearchResult[] =
            getSearchResults(resultRows);

          refinementRows.forEach((refiner: any) => {
            let values: IDataFilterResultValue[] = [];
            refiner.Entries.forEach((item: any) => {
              values.push({
                count: parseInt(item.RefinementCount, 10),
                name: item.RefinementValue.replace('string;#', ''),
                value: item.RefinementToken,
                operator: FilterComparisonOperator.Contains,
              } as IDataFilterResultValue);
            });

            refinementResults.push({
              filterName: refiner.Name,
              values: values,
            });
          });

          results.relevantResults = searchResults;
          results.refinementResults = refinementResults;
          results.totalRows =
            rawSearchResults.PrimaryQueryResult.RelevantResults?.TotalRows || 0;

          if (!isEmpty(rawSearchResults.SpellingSuggestion)) {
            results.spellingSuggestion = rawSearchResults.SpellingSuggestion;
          }
        }
      }

      const rawSearchResults = searchResponse.RawSearchResults;

      if (rawSearchResults.SecondaryQueryResults) {
        const secondaryQueryResults = rawSearchResults.SecondaryQueryResults;

        if (
          Array.isArray(secondaryQueryResults) &&
          secondaryQueryResults.length > 0
        ) {
          let promotedResults: ISharePointSearchPromotedResult[] = [];
          let secondaryResults: ISharePointSearchResultBlock[] = [];

          secondaryQueryResults.forEach((e) => {
            if (e.SpecialTermResults) {
              e.SpecialTermResults.Results.forEach((result: any) => {
                promotedResults.push({
                  title: result.Title,
                  url: result.Url,
                  description: result.Description,
                } as ISharePointSearchPromotedResult);
              });
            }

            if (e.RelevantResults) {
              const secondaryResultItems = getSearchResults(
                e.RelevantResults.Table.Rows
              );

              const secondaryResultBlock: ISharePointSearchResultBlock = {
                title: e.RelevantResults.ResultTitle,
                results: secondaryResultItems,
              };

              if (secondaryResultBlock.results.length > 0) {
                secondaryResults.push(secondaryResultBlock);
              }
            }
          });

          results.promotedResults = promotedResults;
          results.secondaryResults = secondaryResults;
        }
      }

      return results;
    },
    getNextPageParam: (lastPage, allPages) => {
      const currentItemCount = flatMap(
        allPages,
        (page) => page.relevantResults
      ).length;
      if (currentItemCount < lastPage.totalRows) {
        return currentItemCount;
      }
      return undefined;
    },
    keepPreviousData: true,
    enabled: !!searchParams.query,
  });
};

function getSearchResults(resultRows: any): ISharePointSearchResult[] {
  let searchResults: ISharePointSearchResult[] = resultRows.map((elt: any) => {
    let result: ISharePointSearchResult = {};

    elt.Cells.map((item: any) => {
      if (item.Key === 'HtmlFileType' && item.Value) {
        result['FileType'] = item.Value;
      } else if (!result[item.Key]) {
        result[item.Key] = item.Value;
      }
    });

    return result;
  });

  return searchResults;
}
