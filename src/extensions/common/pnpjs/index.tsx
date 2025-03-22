/* eslint-disable no-var */
import { ExtensionContext } from '@microsoft/sp-extension-base';
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';

// import pnp and pnp logging system
import '@pnp/sp/presets/all';
import { spfi, SPFI, SPFx as spSPFx } from '@pnp/sp';
import { LogLevel, PnPLogging } from '@pnp/logging';

import '@pnp/graph/presets/all';
import { graphfi, GraphFI, SPFx as graphSPFx } from '@pnp/graph';

// eslint-disable-next-line no-var
var _sp: SPFI;
var _graph: GraphFI;

export const getSP = (
  context?: ApplicationCustomizerContext | ExtensionContext
): SPFI => {
  if (context != null) {
    _sp = spfi().using(spSPFx(context)).using(PnPLogging(LogLevel.Warning));
  }

  return _sp;
};

export const getGraph = (
  context?: ApplicationCustomizerContext | ExtensionContext
): GraphFI => {
  if (context != null) {
    _graph = graphfi()
      .using(graphSPFx(context))
      .using(PnPLogging(LogLevel.Warning));
  }

  return _graph;
};
