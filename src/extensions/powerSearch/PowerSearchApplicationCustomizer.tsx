import {
  PlaceholderContent,
  BaseApplicationCustomizer,
  PlaceholderName,
} from '@microsoft/sp-application-base';

import {
  FluentProvider,
  IdPrefixProvider,
  webLightTheme,
} from '@fluentui/react-components';

import * as React from 'react';
import { SPFI } from '@pnp/sp';
import ReactDOM from 'react-dom';
import { GraphFI } from '@pnp/graph';
import { queryClient } from '../queryClient';
import { getSP } from '../common/pnpjs';
import { getGraph } from '../common/pnpjs';
import { QueryClientProvider } from '@tanstack/react-query';
import PowerSearch from './components/power-search';
import { fluentOverrides } from './styles';

export interface IPowerSearchApplicationCustomizerProperties {
  testMessage: string;
}

export default class PowerSearchApplicationCustomizer extends BaseApplicationCustomizer<IPowerSearchApplicationCustomizerProperties> {
  private sp: SPFI;
  private graph: GraphFI;
  private appBarPlaceholder: HTMLElement | undefined;
  private bottomPlaceholder: PlaceholderContent | undefined;

  public onInit(): Promise<void> {
    this.sp = getSP(this.context);
    this.graph = getGraph(this.context);

    if (!window.isNavigatedEventSubscribed) {
      this.context.application.navigatedEvent.add(this, this.renderPowerSearch);
      window.isNavigatedEventSubscribed = true;
    }

    return Promise.resolve();
  }

  private renderPowerSearch = () => {
    if (!this.bottomPlaceholder) {
      this.bottomPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Bottom,
          {
            onDispose: this.onDispose,
          }
        );

      if (!this.bottomPlaceholder) {
        console.error(
          `Entry button cannot be displayed as 
            the bottom placeholder cannot be found.`
        );
        return;
      }

      if (this.bottomPlaceholder.domElement) {
        ReactDOM.render(
          <IdPrefixProvider value="PowerSearch">
            <QueryClientProvider client={queryClient}>
              <FluentProvider
                theme={{
                  ...webLightTheme,
                  ...fluentOverrides,
                }}
              >
                <PowerSearch />
              </FluentProvider>
            </QueryClientProvider>
          </IdPrefixProvider>,
          this.bottomPlaceholder.domElement
        );
      }
    }
  };
}
