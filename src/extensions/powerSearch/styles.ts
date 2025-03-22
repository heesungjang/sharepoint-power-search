import { makeStyles, shorthands, tokens } from '@fluentui/react-components';

export const useStyles = makeStyles({
  dialogPositioning: {
    '& .fui-DialogSurface': {
      marginTop: '15vh',
      marginBottom: 'auto',
    },
  },
  dialogBody: {
    padding: '0px',
    display: 'flex',
    overflow: 'hidden',
    flexDirection: 'column',
    height: '100%',
  },
  searchContainer: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    gap: '12px',
    width: '100%',
    borderRadius: '10px',
    height: '100%',
  },
  searchBox: {
    marginTop: '16px',
    width: '98%',
    maxWidth: 'unset',
    marginBottom: '16px',
  },
  resultsContainer: {
    width: '100%',
    marginTop: '8px',
    maxHeight: '550px',
    overflowY: 'auto',
  },
  dialogTitle: {
    fontSize: '18px',
    fontWeight: '600',
    marginBottom: '12px',
  },
  resultsWrapper: {
    width: '100%',
  },
  filterToolbar: {
    display: 'flex',
    width: '100%',
    flexDirection: 'column',
    alignItems: 'start',
    gap: '14px',
    margin: '0',
    paddingLeft: '32px',
  },
  filterSection: {
    display: 'flex',
    flexDirection: 'row',
    alignItems: 'center',
    gap: '4px',
  },
  sectionResults: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
    paddingRight: '8px',
  },
  loadMoreContainer: {
    display: 'flex',
    justifyContent: 'center',
    ...shorthands.padding('16px'),
  },
  loadMoreButton: {
    minWidth: '120px',
  },
  noMoreResults: {
    textAlign: 'center',
    color: tokens.colorNeutralForeground3,
    ...shorthands.padding('16px'),
    fontSize: tokens.fontSizeBase200,
  },
  detailedResultsWrapper: {
    marginTop: '12px',
    maxHeight: '500px',
    overflowY: 'auto',
    ...shorthands.padding('8px', '0'),
  },
  viewerContainer: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    width: '100%',
  },
  viewerHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    padding: '12px 16px',
    borderBottom: '1px solid rgba(0,0,0,0.1)',
  },
  viewerContent: {
    flex: 1,
    height: 'calc(90vh - 120px)',
    width: '100%',
    overflow: 'hidden',
  },
  viewerFrame: {
    width: '100%',
    height: '100%',
    border: 'none',
    backgroundColor: 'white',
  },
  dialogHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    padding: '12px 16px',
    borderBottom: '1px solid rgba(0,0,0,0.05)',
  },
  headerContent: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  viewHeaderContent: {
    display: 'flex',
    alignItems: 'start',
    gap: '8px',
  },
  appIcon: {
    fontSize: '32px',
  },
  appTitleContainer: {
    display: 'flex',
    flexDirection: 'column',
  },
  subdueText: {
    color: 'rgba(0,0,0,0.6)',
  },
  fileTypeIcon: {
    width: '16px',
    height: '16px',
  },
  recentSearchesContainer: {
    display: 'flex',
    width: '100%',
    alignContent: 'flex-start',
    minHeight: '320px',
    flexDirection: 'column',
  },
  sectionLabel: {
    paddingLeft: '16px',
  },
  recentSearchesList: {
    display: 'flex',
    flexDirection: 'column',
    paddingLeft: '12px',
    marginTop: '16px',
    gap: '12px',
  },
  historyTag: {
    minWidth: '120px',
    cursor: 'pointer',
  },
  emptyStateContainer: {
    display: 'flex',
    width: '100%',
    alignItems: 'center',
    justifyContent: 'center',
    minHeight: '320px',
    flexDirection: 'column',
    color: 'rgba(0,0,0,0.5)',
  },
  filterLabel: {
    marginRight: '8px',
    minWidth: '50px',
  },
  toggleFilterButton: {
    minWidth: '50px',
    fontSize: '11px',
  },
  sortToggleButton: {
    minWidth: '50px',
    fontSize: '11px',
  },
  teachingPopoverTrigger: {
    position: 'fixed',
    bottom: '30px',
    right: '30px',
    width: '1px',
    height: '1px',
    pointerEvents: 'none',
  },
  teachingPopoverSurface: {
    maxWidth: '360px',
    padding: '12px',
    boxShadow: '0 6px 12px rgba(0, 0, 0, 0.1)',
  },
  teachingPopoverContent: {
    display: 'flex',
    gap: '12px',
    alignItems: 'start',
    marginTop: '10px',
  },
  teachingPopoverIcon: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: '40px',
    height: '40px',
    borderRadius: '8px',
    backgroundColor: 'var(--colorBrandBackground)',
    color: 'var(--colorNeutralBackground1)',
    flexShrink: 0,
  },
  keyboardShortcut: {
    display: 'inline-block',
    padding: '1px 6px',
    backgroundColor: 'var(--colorNeutralBackground3)',
    borderRadius: '4px',
    fontWeight: 600,
    fontSize: '14px',
    lineHeight: '20px',
    border: '1px solid var(--colorNeutralStroke2)',
    boxShadow: '0 1px 2px rgba(0, 0, 0, 0.06)',
    margin: '0 2px',
  },
});

export const dialogSurfaceStyle = {
  maxWidth: '650px',
  width: '650px',
  borderRadius: '10px',
  boxShadow:
    '0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04)',
  backgroundColor: 'rgba(255, 255, 255, 0.95)',
  backdropFilter: 'blur(10px)',

  marginTop: '12vh',
  marginBottom: 'auto',
  padding: '12px',
  maxHeight: '80vh',
  display: 'flex',
  flexDirection: 'column',
};

export const fluentOverrides = {
  colorBrandBackground: '#242424',
  colorCompoundBrandBackground: '#242424',
  colorBrandBackgroundHover: '#363636', // slightly lighter
  colorBrandBackgroundPressed: '#121212', // slightly darker
  colorBrandForeground1: '#242424',
  colorBrandForeground2: '#363636',
  colorCompoundBrandStroke: '#242424',
  colorCompoundBrandStrokeHover: '#242424',
  colorBrandStroke1: '#242424',
  colorBrandStroke2Contrast: '#D8D9DA', //slightly lighter
  colorStatusSuccessForeground2: '#242424',
  colorStatusSuccessForeground1: '#242424',
  colorNeutralForeground2BrandHover: '#242424',
};
