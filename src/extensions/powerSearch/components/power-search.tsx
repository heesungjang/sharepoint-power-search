import React, {
  useRef,
  useEffect,
  useState,
  useCallback,
  useMemo,
} from 'react';
import {
  Text,
  Dialog,
  Button,
  Spinner,
  DialogBody,
  SearchBox,
  ToggleButton,
  InteractionTag,
  DialogSurface,
  InteractionTagPrimary,
  InteractionTagSecondary,
  TeachingPopover,
  TeachingPopoverBody,
  TeachingPopoverHeader,
  TeachingPopoverTitle,
  TeachingPopoverSurface,
  TeachingPopoverTrigger,
  TeachingPopoverFooter,
} from '@fluentui/react-components';

import {
  ArrowExitRegular,
  ChevronLeftFilled,
  AppTitleFilled,
  OpenRegular,
  HistoryRegular,
  ImageRegular,
} from '@fluentui/react-icons';

import {
  getFileTypeIconAsUrl,
  initializeFileTypeIcons,
} from '@fluentui/react-file-type-icons';

import {
  ResultItem,
  SearchResults,
  SearchResultsSkeleton,
} from './search-results';

import { useSearch } from '../hooks/use-search';
import { useStyles, dialogSurfaceStyle } from '../styles';
import { useKeyboardShortcut } from '../hooks/use-keyboard-shortcut';
import { useDebounceValue } from 'usehooks-ts';

initializeFileTypeIcons();

interface FileType {
  label: string;
  extension: string | null;
  icon?: React.ReactNode;
}

interface ViewerState {
  isOpen: boolean;
  url: string;
  title: string;
  type: string;
}

const FILE_TYPES: FileType[] = [
  { label: 'All', extension: null },
  { label: 'docx', extension: 'docx' },
  { label: 'pptx', extension: 'pptx' },
  { label: 'xlsx', extension: 'xlsx' },
  { label: 'pdf', extension: 'pdf' },
  { label: 'page', extension: 'aspx' },
  { label: 'images', extension: null, icon: <ImageRegular /> },
];

const STANDARD_TYPES = {
  docx: ['docx', 'dotx'],
  pptx: ['pptx', 'potx', 'ppsx'],
  xlsx: ['xlsx', 'xltx'],
  pdf: ['pdf'],
  page: ['aspx'],
};

const IMAGE_TYPES = ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'ico', 'tiff', 'webp'];

const LOCAL_STORAGE_KEYS = {
  FILE_TYPES: 'powerSearchFileTypes',
  SORT_ORDER: 'powerSearchSortOrder',
  HISTORY: 'powerSearchHistory',
};

/**
 * PowerSearch Component
 *
 * A Spotlight-like search component that opens with Ctrl+ArrowUp or Command+ArrowUp.
 * Provides a clean, minimal search interface with a semi-transparent dialog.
 *
 * Features:
 * - Keyboard shortcut to open (Ctrl+ArrowUp or Command+ArrowUp)
 * - Clean, minimal UI inspired by Mac's Spotlight
 * - Search input with clear button
 * - Results display when search query is entered
 * - Debounced search to prevent excessive API calls
 *
 * @returns {React.ReactElement} The PowerSearch component
 */
const PowerSearch: React.FC = () => {
  const styles = useStyles();
  const searchInputRef = useRef<HTMLInputElement>(null);

  const [uiState, setUiState] = useState({
    isDialogOpen: false,
    isHovered: false,
    isViewTransitioning: false,
    currentView: null as string | null,
    startRow: 0,
    showTeachingPopover: false,
  });

  const [preferences, setPreferences] = useState({
    selectedFileTypes: (() => {
      const saved = localStorage.getItem(LOCAL_STORAGE_KEYS.FILE_TYPES);
      return saved ? new Set(JSON.parse(saved)) : new Set(['all']);
    })(),
    sortOrder: (() => {
      const saved = localStorage.getItem(LOCAL_STORAGE_KEYS.SORT_ORDER);
      return saved
        ? (JSON.parse(saved) as 'relevance' | 'desc' | 'asc')
        : 'relevance';
    })(),
    recentSearches: (() => {
      const saved = localStorage.getItem(LOCAL_STORAGE_KEYS.HISTORY);
      return saved ? JSON.parse(saved) : [];
    })(),
  });

  const [searchQuery, setSearchQuery] = useState('');
  const [debouncedSearchQuery] = useDebounceValue(searchQuery, 300);
  const [sectionRefinementFilter, setSectionRefinementFilter] = useState<
    string[]
  >([]);

  const [viewerState, setViewerState] = useState<ViewerState>({
    isOpen: false,
    url: '',
    title: '',
    type: '',
  });

  const saveToLocalStorage = useCallback((key: string, value: any) => {
    localStorage.setItem(key, JSON.stringify(value));
  }, []);

  const openDialog = useCallback(() => {
    setUiState((prev) => ({ ...prev, isDialogOpen: true }));
  }, []);

  const closeDialog = useCallback(() => {
    setUiState((prev) => ({ ...prev, isDialogOpen: false }));
    setSearchQuery('');
  }, []);

  useKeyboardShortcut(
    [
      { key: 'k', ctrlKey: true },
      { key: 'k', metaKey: true },
    ],
    openDialog
  );

  const handleSearchChange = useCallback(
    (
      event: React.ChangeEvent<HTMLInputElement> | React.MouseEvent,
      data?: { value: string }
    ) => {
      const newValue =
        data?.value ||
        ('target' in event && event.target instanceof HTMLInputElement
          ? event.target.value
          : '');
      setSearchQuery(newValue);
    },
    []
  );

  const getRefinementFilters = useCallback(
    (selected: Set<string>): string[] => {
      if (selected.has('all')) return [];

      const createTypeFilter = (extensions: string[]): string =>
        extensions.length > 1
          ? `or(${extensions
              .map((ext) => `FileType:equals("${ext}")`)
              .join(',')})`
          : `FileType:equals("${extensions[0]}")`;

      const fileTypeFilters = [];

      if (selected.has('images')) {
        fileTypeFilters.push(createTypeFilter(IMAGE_TYPES));
      }

      const selectedStandardTypes = Object.entries(STANDARD_TYPES)
        .filter(([key]) => selected.has(key))
        .map(([_, extensions]) => createTypeFilter(extensions));

      fileTypeFilters.push(...selectedStandardTypes);

      return fileTypeFilters.length > 1
        ? [`or(${fileTypeFilters.join(',')})`]
        : fileTypeFilters;
    },
    []
  );

  const getAllRefinementFilters = useCallback(() => {
    return uiState.currentView && sectionRefinementFilter.length > 0
      ? sectionRefinementFilter
      : getRefinementFilters(preferences.selectedFileTypes as Set<string>);
  }, [
    uiState.currentView,
    getRefinementFilters,
    sectionRefinementFilter,
    preferences.selectedFileTypes,
  ]);

  const handleFileTypeClick = useCallback(
    (fileType: string) => {
      setPreferences((prev) => {
        const newSelected = new Set(prev.selectedFileTypes);

        if (fileType === 'all') {
          saveToLocalStorage(LOCAL_STORAGE_KEYS.FILE_TYPES, ['all']);
          return { ...prev, selectedFileTypes: new Set(['all']) };
        }

        newSelected.delete('all');

        if (newSelected.has(fileType)) {
          newSelected.delete(fileType);

          if (newSelected.size === 0) {
            newSelected.add('all');
          }
        } else {
          newSelected.add(fileType);

          if (newSelected.size === FILE_TYPES.length - 1) {
            saveToLocalStorage(LOCAL_STORAGE_KEYS.FILE_TYPES, ['all']);
            return { ...prev, selectedFileTypes: new Set(['all']) };
          }
        }

        saveToLocalStorage(
          LOCAL_STORAGE_KEYS.FILE_TYPES,
          Array.from(newSelected)
        );
        return { ...prev, selectedFileTypes: newSelected };
      });
    },
    [saveToLocalStorage]
  );

  const handleSectionClick = useCallback((section: string) => {
    setUiState((prev) => ({
      ...prev,
      isViewTransitioning: true,
      currentView: section,
      startRow: 0,
    }));

    const sectionFilters = {
      documents: [
        'or(FileType:equals("docx"),FileType:equals("doc"),FileType:equals("pdf"))',
      ],
      pages: ['or(FileType:equals("aspx"),FileType:equals("page"))'],
      other: [
        'not(or(FileType:equals("docx"),FileType:equals("doc"),FileType:equals("pdf"),FileType:equals("aspx"),FileType:equals("page")))',
      ],
      default: [],
    };

    setSectionRefinementFilter(
      sectionFilters[section as keyof typeof sectionFilters] ||
        sectionFilters.default
    );
  }, []);

  const handleBackToSearch = useCallback(() => {
    setUiState((prev) => ({
      ...prev,
      isViewTransitioning: true,
      currentView: null,
    }));
    setSectionRefinementFilter([]);
  }, []);

  const addToSearchHistory = useCallback(
    (query: string) => {
      if (!query.trim()) return;

      setPreferences((prev) => {
        const filtered = prev.recentSearches.filter(
          (item: string) => item !== query
        );
        const updated = [query, ...filtered].slice(0, 6);
        saveToLocalStorage(LOCAL_STORAGE_KEYS.HISTORY, updated);
        return { ...prev, recentSearches: updated };
      });
    },
    [saveToLocalStorage]
  );

  const removeFromSearchHistory = useCallback(
    (query: string) => {
      setPreferences((prev) => {
        const updated = prev.recentSearches.filter(
          (item: string) => item !== query
        );
        saveToLocalStorage(LOCAL_STORAGE_KEYS.HISTORY, updated);
        return { ...prev, recentSearches: updated };
      });
    },
    [saveToLocalStorage]
  );

  const refinementKey = uiState.currentView
    ? sectionRefinementFilter.join(',')
    : Array.from(preferences.selectedFileTypes).sort().join(',');

  const searchSortOrder =
    preferences.sortOrder === 'relevance' ? 'desc' : preferences.sortOrder;

  const { data, fetchNextPage, isFetchingNextPage, isFetching, hasNextPage } =
    useSearch(
      {
        query: debouncedSearchQuery,
        refinementFilters: getAllRefinementFilters(),
        startRow: uiState.startRow,
        sortProperty: 'LastModifiedTime',
        sortDirection: preferences.sortOrder === 'desc' ? 'desc' : 'asc',
      },
      refinementKey,
      searchSortOrder,
      [debouncedSearchQuery, preferences.sortOrder, refinementKey],
      'site'
    );

  const searchResults = useMemo(() => {
    if (!data?.pages) return [];

    return data.pages.reduce((allResults, page) => {
      if (page.relevantResults && Array.isArray(page.relevantResults)) {
        return [...allResults, ...page.relevantResults];
      }
      return allResults;
    }, []);
  }, [data]);

  const handleLoadMore = useCallback(() => {
    if (!isFetchingNextPage) {
      fetchNextPage();
    }
  }, [fetchNextPage, isFetchingNextPage]);

  const handleSortOrderChange = useCallback(
    (newOrder: 'relevance' | 'desc' | 'asc') => {
      setPreferences((prev) => {
        saveToLocalStorage(LOCAL_STORAGE_KEYS.SORT_ORDER, newOrder);
        return { ...prev, sortOrder: newOrder };
      });
    },
    [saveToLocalStorage]
  );

  const closeViewer = useCallback(() => {
    setViewerState((prev) => ({ ...prev, isOpen: false }));
  }, []);

  const handleDialogOpenChange = useCallback(
    (event: any, data: { open: boolean }) => {
      if (
        !(
          event.target && event.target.getAttribute('data-dismiss') === 'dialog'
        )
      ) {
        if (data.open) {
          openDialog();
        } else {
          closeDialog();
        }
      }
    },
    [openDialog, closeDialog]
  );

  useEffect(() => {
    if (debouncedSearchQuery.trim()) {
      addToSearchHistory(debouncedSearchQuery);
    }
  }, [debouncedSearchQuery, addToSearchHistory]);

  useEffect(() => {
    if (uiState.isViewTransitioning && !isFetching) {
      setUiState((prev) => ({ ...prev, isViewTransitioning: false }));
    }
  }, [isFetching, uiState.isViewTransitioning]);

  useEffect(() => {
    if (uiState.isDialogOpen && searchInputRef.current) {
      searchInputRef.current?.focus();
    }
  }, [uiState.isDialogOpen]);

  useEffect(() => {
    const handleViewItem = (event: CustomEvent) => {
      setViewerState({
        isOpen: true,
        url: event.detail.url,
        title: event.detail.title,
        type: event.detail.type,
      });
    };

    window.addEventListener(
      'powerSearchViewItem',
      handleViewItem as EventListener
    );
    return () => {
      window.removeEventListener(
        'powerSearchViewItem',
        handleViewItem as EventListener
      );
    };
  }, []);

  useEffect(() => {
    const hasSeenTeachingPopover = localStorage.getItem(
      'powerSearchTeachingPopoverShown'
    );
    if (!hasSeenTeachingPopover) {
      setUiState((prev) => ({ ...prev, showTeachingPopover: true }));
    }
  }, []);

  const dismissTeachingPopover = useCallback(() => {
    setUiState((prev) => ({ ...prev, showTeachingPopover: false }));
    localStorage.setItem('powerSearchTeachingPopoverShown', 'true');
  }, []);

  const getCurrentViewTitle = () => {
    const viewTitles = {
      documents: 'Files',
      pages: 'Pages',
      other: 'Other Items',
      default: '',
    };

    return (
      viewTitles[uiState.currentView as keyof typeof viewTitles] ||
      viewTitles.default
    );
  };

  const DIALOG_SURFACE_STYLE = viewerState.isOpen
    ? {
        ...(dialogSurfaceStyle as any),
        maxHeight: '90vh',
        height: '90vh',
        width: '50vw',
        marginTop: '5vh',
        maxWidth: '1400px',
      }
    : (dialogSurfaceStyle as any);

  return (
    <>
      <TeachingPopover
        open={uiState.showTeachingPopover}
        onOpenChange={(e, data) => {
          if (!data.open) {
            dismissTeachingPopover();
          }
        }}
        positioning={{ position: 'above', align: 'end' }}
      >
        <TeachingPopoverSurface
          className={styles.teachingPopoverSurface}
          style={{
            width: '300px',
            position: 'absolute',
            top: '100px',
            right: '80px',
          }}
        >
          <TeachingPopoverHeader>
            <TeachingPopoverTitle>Welcome to Power Search</TeachingPopoverTitle>
          </TeachingPopoverHeader>
          <TeachingPopoverBody>
            <div className={styles.teachingPopoverContent}>
              <div className={styles.teachingPopoverIcon}>
                <AppTitleFilled fontSize={32} />
              </div>
              <div>
                <Text weight="semibold" block>
                  Quick access to all your content
                </Text>
                <Text block style={{ marginTop: '8px' }}>
                  Press{' '}
                  <span className={styles.keyboardShortcut}>
                    {navigator.platform.includes('Mac') ? '⌘K' : 'Ctrl+K'}{' '}
                  </span>{' '}
                  anytime to instantly search across your organization's files,
                  pages, and more.
                </Text>
              </div>
            </div>
          </TeachingPopoverBody>
          <TeachingPopoverFooter
            primary={{
              children: 'Try it now',
              onClick: () => {
                openDialog();
              },
            }}
            secondary="Dismiss"
          />
        </TeachingPopoverSurface>
      </TeachingPopover>

      <Dialog
        open={uiState.isDialogOpen}
        onOpenChange={handleDialogOpenChange}
        surfaceMotion={null}
      >
        <DialogSurface style={DIALOG_SURFACE_STYLE}>
          <div className={styles.dialogHeader}>
            <div className={styles.headerContent}>
              {uiState.currentView ? (
                <div className={styles.viewHeaderContent}>
                  <Button
                    appearance="transparent"
                    icon={<ChevronLeftFilled />}
                    onClick={handleBackToSearch}
                    aria-label="Back to search"
                  />
                  <Text weight="semibold" size={300}>
                    {getCurrentViewTitle()}
                  </Text>
                  <Text size={200} className={styles.subdueText}>
                    Results for "{searchQuery}"
                  </Text>
                </div>
              ) : (
                <>
                  <AppTitleFilled className={styles.appIcon} />
                  <div className={styles.appTitleContainer}>
                    <Text weight="semibold" size={300}>
                      SharePoint Power Search
                    </Text>
                    <Text size={200} className={styles.subdueText}>
                      Find anything in your organization
                    </Text>
                  </div>
                </>
              )}
            </div>
            <Button
              appearance="transparent"
              icon={<ArrowExitRegular />}
              aria-label="Close"
              data-dismiss="dialog"
              onClick={() => handleDialogOpenChange({} as any, { open: false })}
            />
          </div>

          <DialogBody className={styles.dialogBody}>
            {viewerState.isOpen ? (
              <div className={styles.viewerContainer}>
                <div className={styles.viewerHeader}>
                  <div className={styles.viewHeaderContent}>
                    <Button
                      appearance="transparent"
                      icon={<ChevronLeftFilled />}
                      onClick={closeViewer}
                      aria-label="Back to results"
                    />
                    <Text weight="semibold" size={300}>
                      {viewerState.title}
                    </Text>
                  </div>
                  <Button
                    appearance="transparent"
                    icon={<OpenRegular />}
                    aria-label="Open in new tab"
                    onClick={() => window.open(viewerState.url, '_blank')}
                  />
                </div>
                <div className={styles.viewerContent}>
                  <iframe
                    src={viewerState.url}
                    title={viewerState.title}
                    className={styles.viewerFrame}
                    sandbox="allow-same-origin allow-scripts allow-forms"
                  />
                </div>
              </div>
            ) : (
              <div className={styles.searchContainer}>
                {!uiState.currentView && (
                  <SearchBox
                    placeholder="Search for anything..."
                    value={searchQuery}
                    onChange={handleSearchChange}
                    autoFocus
                    ref={searchInputRef}
                    className={styles.searchBox}
                    size="large"
                    appearance="filled-lighter"
                    style={{
                      boxShadow: uiState.isHovered
                        ? '0 4px 12px rgba(0, 0, 0, 0.08)'
                        : '0 2px 8px rgba(0, 0, 0, 0.05)',
                      borderRadius: '8px',
                      transition: 'box-shadow 0.3s ease',
                    }}
                    onMouseEnter={() =>
                      setUiState((prev) => ({ ...prev, isHovered: true }))
                    }
                    onMouseLeave={() =>
                      setUiState((prev) => ({ ...prev, isHovered: false }))
                    }
                  />
                )}

                {/* Show recent searches only when no query and we have history items */}
                {!searchQuery &&
                  preferences.recentSearches.length > 0 &&
                  !uiState.currentView && (
                    <div className={styles.recentSearchesContainer}>
                      <Text
                        weight="semibold"
                        size={200}
                        className={styles.sectionLabel}
                      >
                        Recent Searches
                      </Text>

                      <div className={styles.recentSearchesList}>
                        {preferences.recentSearches.map(
                          (query: string, index: number) => (
                            <InteractionTag key={index} size="small">
                              <InteractionTagPrimary
                                className={styles.historyTag}
                                icon={<HistoryRegular />}
                                hasSecondaryAction
                                onClick={() => {
                                  handleSearchChange({
                                    target: { value: query },
                                  } as any);
                                }}
                              >
                                {query}
                              </InteractionTagPrimary>
                              <InteractionTagSecondary
                                aria-label="remove"
                                onClick={() => removeFromSearchHistory(query)}
                              />
                            </InteractionTag>
                          )
                        )}
                      </div>
                    </div>
                  )}

                {/* Show placeholder when no search history and no search query */}
                {!searchQuery &&
                  preferences.recentSearches.length === 0 &&
                  !uiState.currentView && (
                    <div className={styles.emptyStateContainer}>
                      <Text size={300}>Empty</Text>
                    </div>
                  )}

                {searchQuery && !uiState.currentView && (
                  <div className={styles.filterToolbar}>
                    <div className={styles.filterSection}>
                      <Text
                        weight="semibold"
                        size={200}
                        className={styles.filterLabel}
                      >
                        File type
                      </Text>
                      {FILE_TYPES.map((fileType) => {
                        const icon =
                          fileType.icon ||
                          (fileType.extension ? (
                            <img
                              src={getFileTypeIconAsUrl({
                                extension: fileType.extension,
                                size: 16,
                              })}
                              alt={fileType.label}
                              className={styles.fileTypeIcon}
                            />
                          ) : undefined);

                        return (
                          <ToggleButton
                            key={fileType.label.toLowerCase()}
                            icon={icon}
                            appearance="subtle"
                            size="small"
                            className={styles.toggleFilterButton}
                            checked={preferences.selectedFileTypes.has(
                              fileType.label.toLowerCase()
                            )}
                            onClick={() =>
                              handleFileTypeClick(fileType.label.toLowerCase())
                            }
                          >
                            {fileType.label}
                          </ToggleButton>
                        );
                      })}
                    </div>

                    <div className={styles.filterSection}>
                      <Text
                        weight="semibold"
                        size={200}
                        className={styles.filterLabel}
                      >
                        Sort by
                      </Text>
                      <ToggleButton
                        appearance="subtle"
                        size="small"
                        className={styles.sortToggleButton}
                        checked={preferences.sortOrder === 'relevance'}
                        onClick={() => {
                          handleSortOrderChange('relevance');
                        }}
                      >
                        Relevance
                      </ToggleButton>
                      <ToggleButton
                        appearance="subtle"
                        size="small"
                        className={styles.sortToggleButton}
                        checked={preferences.sortOrder === 'desc'}
                        onClick={() => {
                          handleSortOrderChange('desc');
                        }}
                      >
                        Newer
                      </ToggleButton>
                      <ToggleButton
                        appearance="subtle"
                        size="small"
                        className={styles.sortToggleButton}
                        checked={preferences.sortOrder === 'asc'}
                        onClick={() => {
                          handleSortOrderChange('asc');
                        }}
                      >
                        Older
                      </ToggleButton>
                    </div>
                  </div>
                )}

                {uiState.currentView ? (
                  <div className={styles.resultsContainer}>
                    {uiState.isViewTransitioning ? (
                      <SearchResultsSkeleton />
                    ) : (
                      <>
                        {searchResults.map((result: any, index: number) => (
                          <ResultItem
                            key={index}
                            result={result}
                            query={searchQuery}
                            index={index}
                          />
                        ))}

                        {hasNextPage && (
                          <div className={styles.loadMoreContainer}>
                            <Button
                              appearance="subtle"
                              className={styles.loadMoreButton}
                              onClick={handleLoadMore}
                              disabled={isFetchingNextPage}
                            >
                              {isFetchingNextPage ? (
                                <Spinner size="tiny" />
                              ) : (
                                'Load more'
                              )}
                            </Button>
                          </div>
                        )}

                        {!hasNextPage && searchResults.length > 0 && (
                          <div className={styles.noMoreResults}>
                            No more results
                          </div>
                        )}
                      </>
                    )}
                  </div>
                ) : (
                  <SearchResults
                    query={searchQuery}
                    data={searchResults}
                    isLoading={isFetching}
                    onSectionClick={handleSectionClick}
                  />
                )}
              </div>
            )}
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </>
  );
};

export default PowerSearch;
