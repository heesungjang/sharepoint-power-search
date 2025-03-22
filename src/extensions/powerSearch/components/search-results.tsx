import React, { useEffect, useRef } from 'react';
import {
  Text,
  Avatar,
  makeStyles,
  shorthands,
  tokens,
  SkeletonItem,
  Menu,
  MenuTrigger,
  MenuList,
  MenuItem,
  MenuPopover,
} from '@fluentui/react-components';
import {
  CalendarRegular,
  ChevronRightRegular,
  DocumentRegular,
  OpenRegular,
  FolderRegular,
  ArrowDownloadRegular,
  LinkRegular,
  EyeRegular,
} from '@fluentui/react-icons';
import {
  getFileTypeIconAsUrl,
  initializeFileTypeIcons,
} from '@fluentui/react-file-type-icons';

initializeFileTypeIcons();

interface SearchResultsProps {
  query: string;
  data?: any[];
  isLoading?: boolean;
  onSectionClick?: (section: string) => void;
}

interface ResultItemProps {
  result: any;
  query: string;
  index?: number;
}

const useStyles = makeStyles({
  resultsContainer: {
    width: '100%',
    marginTop: '12px',
    maxHeight: '500px',
    overflowY: 'auto',
    ...shorthands.padding('8px', '0'),
  },
  resultItem: {
    display: 'flex',
    alignItems: 'start',
    ...shorthands.padding('12px', '16px'),
    ...shorthands.borderRadius('6px'),
    cursor: 'pointer',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
  resultIcon: {
    marginRight: '12px',
    marginTop: '4px',
    color: tokens.colorBrandForeground1,
  },
  resultContent: {
    display: 'flex',
    flexDirection: 'column',
    flexGrow: 1,
    overflow: 'hidden',
  },
  resultTitle: {
    fontWeight: tokens.fontWeightSemibold,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  },
  resultPath: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  },
  resultSummary: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    marginTop: '4px',
    display: '-webkit-box',
    WebkitLineClamp: 2,
    WebkitBoxOrient: 'vertical',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  },
  noResults: {
    ...shorthands.padding('16px'),
    textAlign: 'center',
    color: tokens.colorNeutralForeground3,
  },
  sectionHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    cursor: 'pointer',
    ...shorthands.padding('8px', '12px'),
    ...shorthands.borderRadius('6px'),
    ...shorthands.margin('12px', '4px', '6px', '4px'),
    backgroundColor: tokens.colorNeutralBackground2,
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground2Hover,
    },
  },
  sectionTitleWrapper: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  sectionTitle: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
    backgroundColor: tokens.colorNeutralBackground4,
    ...shorthands.padding('1px', '12px'),
    ...shorthands.borderRadius('5px'),
  },
  sectionCount: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    ...shorthands.padding('1px', '12px'),
    ...shorthands.borderRadius('5px'),
    backgroundColor: tokens.colorNeutralBackground4,
  },
  sectionArrow: {
    color: tokens.colorNeutralForeground3,
  },
  skeletonContainer: {
    display: 'flex',
    width: '100%',
    flexDirection: 'column',
    gap: '12px',
    ...shorthands.padding('16px'),
  },
  skeletonItem: {
    display: 'flex',
    alignItems: 'center',
    ...shorthands.padding('12px', '16px'),
  },
  skeletonIcon: {
    marginRight: '12px',
  },
  skeletonContent: {
    display: 'flex',
    flexDirection: 'column',
    flexGrow: 1,
    gap: '8px',
  },
  resultMetadata: {
    display: 'flex',
    gap: '12px',
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    marginTop: '4px',
    flexWrap: 'wrap',
  },
  metadataItem: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
  },
  metadataIcon: {
    fontSize: '12px',
    color: tokens.colorNeutralForeground3,
  },
  sitePathInfo: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    marginTop: '2px',
  },
  pathDisplayWrapper: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
  },
  menuItem: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  menuIcon: {
    fontSize: '16px',
  },
  sectionMoreText: {
    color: tokens.colorNeutralForeground2,
  },
  moreWrapper: {
    display: 'flex',
    alignItems: 'center',
  },
  avatarMini: {
    marginRight: '4px',
  },
  authorNameSpan: {
    marginBottom: '2px',
  },
  fileIconImg: {
    width: '24px',
    height: '24px',
  },
});

const getFileIcon = (fileType: string | null) => {
  if (fileType === 'event' || fileType === 'ics') {
    return <CalendarRegular />;
  }

  return (
    <img
      src={getFileTypeIconAsUrl({
        extension: fileType || '',
        size: 16,
      })}
      alt={fileType || 'document'}
      style={{ width: '24px', height: '24px' }}
    />
  );
};

const formatDate = (dateString: string) => {
  if (!dateString) return '';
  try {
    const date = new Date(dateString);
    return date.toLocaleDateString(undefined, {
      year: 'numeric',
      month: 'short',
      day: 'numeric',
    });
  } catch (e) {
    return dateString;
  }
};

const getSiteName = (url: string) => {
  try {
    const urlObj = new URL(url);
    const pathParts = urlObj.pathname.split('/').filter(Boolean);
    const siteIndex = pathParts.findIndex(
      (part) => part.toLowerCase() === 'sites' || part.toLowerCase() === 'teams'
    );
    if (siteIndex >= 0 && pathParts.length > siteIndex + 1) {
      return pathParts[siteIndex + 1];
    }
    return urlObj.hostname.split('.')[0];
  } catch (e) {
    return '';
  }
};

const formatPath = (path: string) => {
  try {
    const url = new URL(path);
    return url.pathname;
  } catch {
    return path;
  }
};

const getFolderPath = (fullPath: string) => {
  try {
    const url = new URL(fullPath);
    const urlPath = url.pathname;
    const lastSlashIndex = urlPath.lastIndexOf('/');
    if (lastSlashIndex !== -1) {
      return url.origin + urlPath.substring(0, lastSlashIndex);
    }
    return url.origin;
  } catch {
    const lastSlashIndex = fullPath.lastIndexOf('/');
    if (lastSlashIndex !== -1) {
      return fullPath.substring(0, lastSlashIndex);
    }
    return fullPath;
  }
};

const highlightSearchTerms = (text: string, query: string) => {
  if (!text || !query || query.trim() === '') return text;

  let cleanText = text.replace(/<ddd\/>/g, '...');

  if (cleanText.includes('<c0>')) {
    return cleanText.replace(
      /<c0>(.*?)<\/c0>/g,
      (_, match) => `<span class="highlight">${match}</span>`
    );
  }

  const terms = query
    .trim()
    .split(/\s+/)
    .filter((term) => term.length > 2);
  if (terms.length === 0) return cleanText;

  const escapedTerms = terms.map((term) =>
    term.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  );

  const regex = new RegExp(`(${escapedTerms.join('|')})`, 'gi');
  return cleanText.replace(
    regex,
    (match) => `<span class="highlight">${match}</span>`
  );
};

/**
 * Skeleton loader component for search results
 */
export const SearchResultsSkeleton: React.FC = () => {
  const styles = useStyles();
  const skeletonItems = Array(5).fill(0);

  return (
    <div className={styles.skeletonContainer}>
      {skeletonItems.map((_, index) => (
        <div key={index} className={styles.skeletonItem}>
          <div className={styles.skeletonIcon}></div>
          <div className={styles.skeletonContent}>
            <SkeletonItem
              shape="rectangle"
              style={{ width: '70%', height: '20px' }}
            />
            <SkeletonItem
              shape="rectangle"
              style={{ width: '40%', height: '16px' }}
            />
            <SkeletonItem
              shape="rectangle"
              style={{ width: '90%', height: '32px' }}
            />
          </div>
        </div>
      ))}
    </div>
  );
};

/**
 * Component to display an individual search result
 */
export const ResultItem: React.FC<ResultItemProps> = ({
  result,
  query,
  index = 0,
}) => {
  const styles = useStyles();
  const siteName = result.Path ? getSiteName(result.Path) : '';
  const path = result.Path ? formatPath(result.Path) : '';
  const sitePathDisplay = siteName ? `${siteName} - ${path}` : path;
  const isPage = ['aspx', 'page'].includes(
    result.FileType?.toLowerCase() || ''
  );

  const handleOpenFolder = () => {
    try {
      let folderPath = getFolderPath(result.Path);
      folderPath = folderPath.replace(/\/Forms/g, '');
      const enhancedUrl = `${folderPath}?q=${encodeURIComponent(
        result.Filename
      )}`;
      window.open(enhancedUrl, '_blank');
    } catch (e) {
      let folderPath = getFolderPath(result.Path);
      folderPath = folderPath.replace(/\/Forms/g, '');
      window.open(folderPath, '_blank');
    }
  };

  const handleDownload = () => {
    const link = document.createElement('a');
    link.href = result.Path;
    link.download = result.Filename || 'download';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handleCopyLink = () => {
    navigator.clipboard
      .writeText(result.Path)
      .then(() => console.log('Link copied to clipboard'))
      .catch((err) => console.error('Failed to copy link: ', err));
  };

  const handleView = () => {
    window.dispatchEvent(
      new CustomEvent('powerSearchViewItem', {
        detail: {
          url: result.Path,
          title: result.Filename || 'Page',
          type: result.FileType,
        },
      })
    );
  };

  const getAuthorName = (author: string) => {
    return author.includes('|') ? author.split('|')[1].trim() : author;
  };

  return (
    <Menu>
      <MenuTrigger disableButtonEnhancement>
        <div key={index} className={styles.resultItem}>
          <div className={styles.resultIcon}>
            {getFileIcon(result.FileType)}
          </div>
          <div className={styles.resultContent}>
            <Text className={styles.resultTitle}>{result.Filename}</Text>

            {sitePathDisplay && (
              <div className={styles.pathDisplayWrapper}>
                <Text className={styles.sitePathInfo}>
                  {sitePathDisplay.length > 40
                    ? `${sitePathDisplay.substring(0, 37)}...`
                    : sitePathDisplay}
                </Text>
              </div>
            )}

            <div className={styles.resultMetadata}>
              {result.CreatedBy && (
                <div className={styles.metadataItem}>
                  <Avatar
                    name={result.CreatedBy}
                    size={16}
                    color="colorful"
                    className={styles.avatarMini}
                  />
                  <span className={styles.authorNameSpan}>
                    {getAuthorName(result.CreatedBy)} -{' '}
                    {formatDate(result.LastModifiedTime)}
                  </span>
                </div>
              )}

              {result.LastModifiedTime && (
                <div className={styles.metadataItem}>
                  <CalendarRegular className={styles.metadataIcon} />
                  <span>{formatDate(result.LastModifiedTime)}</span>
                </div>
              )}

              {result.Size && (
                <div className={styles.metadataItem}>
                  <DocumentRegular className={styles.metadataIcon} />
                  <span>{Math.round(result.Size / 1024)} KB</span>
                </div>
              )}
            </div>

            {result.HitHighlightedSummary && (
              <div
                className={styles.resultSummary}
                dangerouslySetInnerHTML={{
                  __html: highlightSearchTerms(
                    result.HitHighlightedSummary,
                    query
                  ),
                }}
              />
            )}
          </div>
        </div>
      </MenuTrigger>
      <MenuPopover>
        <MenuList>
          {isPage && (
            <MenuItem onClick={handleView}>
              <div className={styles.menuItem}>
                <EyeRegular className={styles.menuIcon} />
                <span>View</span>
              </div>
            </MenuItem>
          )}
          <MenuItem onClick={() => window.open(result.Path, '_blank')}>
            <div className={styles.menuItem}>
              <OpenRegular className={styles.menuIcon} />
              <span>Open</span>
            </div>
          </MenuItem>
          <MenuItem onClick={handleOpenFolder}>
            <div className={styles.menuItem}>
              <FolderRegular className={styles.menuIcon} />
              <span>Open in folder</span>
            </div>
          </MenuItem>
          <MenuItem onClick={handleDownload}>
            <div className={styles.menuItem}>
              <ArrowDownloadRegular className={styles.menuIcon} />
              <span>Download</span>
            </div>
          </MenuItem>
          <MenuItem onClick={handleCopyLink}>
            <div className={styles.menuItem}>
              <LinkRegular className={styles.menuIcon} />
              <span>Copy link</span>
            </div>
          </MenuItem>
        </MenuList>
      </MenuPopover>
    </Menu>
  );
};

/**
 * Component to display search results
 */
export const SearchResults: React.FC<SearchResultsProps> = ({
  query,
  data = [],
  isLoading = false,
  onSectionClick,
}) => {
  const styles = useStyles();
  const highlightRef = useRef<HTMLStyleElement | null>(null);

  useEffect(() => {
    if (!highlightRef.current) {
      const style = document.createElement('style');
      style.innerHTML = `
        .highlight {
          background-color: rgba(255, 213, 0, 0.3);
          font-weight: 600;
          padding: 0 1px;
          border-radius: 2px;
        }
      `;
      document.head.appendChild(style);
      highlightRef.current = style;
    }

    return () => {
      if (highlightRef.current) {
        document.head.removeChild(highlightRef.current);
        highlightRef.current = null;
      }
    };
  }, []);

  if (!query) return null;

  if (isLoading && (!data || data.length === 0)) {
    return <SearchResultsSkeleton />;
  }

  if (!data || data.length === 0) {
    return (
      <div className={styles.noResults}>
        <Text>No results found for "{query}"</Text>
      </div>
    );
  }

  const documentResults = data.filter((item) =>
    ['docx', 'doc', 'pdf'].includes(item.FileType?.toLowerCase() || '')
  );
  const pageResults = data.filter((item) =>
    ['aspx', 'page'].includes(item.FileType?.toLowerCase() || '')
  );
  const otherResults = data.filter(
    (item) =>
      !['docx', 'doc', 'pdf', 'aspx', 'page'].includes(
        item.FileType?.toLowerCase() || ''
      )
  );

  const renderSection = (title: string, sectionKey: string, results: any[]) => (
    <>
      <div
        className={styles.sectionHeader}
        onClick={() => onSectionClick?.(sectionKey)}
      >
        <div className={styles.sectionTitleWrapper}>
          <Text className={styles.sectionTitle}>{title}</Text>
          <Text className={styles.sectionCount}>{results.length} +</Text>
        </div>
        <div className={styles.moreWrapper}>
          <Text size={200} className={styles.sectionMoreText}>
            more
          </Text>
          <ChevronRightRegular fontSize={16} className={styles.sectionArrow} />
        </div>
      </div>
      {results.slice(0, 3).map((result, index) => (
        <ResultItem key={index} result={result} query={query} index={index} />
      ))}
    </>
  );

  return (
    <div className={styles.resultsContainer}>
      {documentResults.length > 0 &&
        renderSection('Files', 'documents', documentResults)}

      {pageResults.length > 0 && renderSection('Pages', 'pages', pageResults)}

      {otherResults.length > 0 && renderSection('Other', 'other', otherResults)}
    </div>
  );
};
