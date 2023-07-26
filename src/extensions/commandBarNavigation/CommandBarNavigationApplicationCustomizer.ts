import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,  
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { ITopCommandBarProps } from './components/ITopCommandBarProps';
import { TopCommandBar } from './components/TopCommandBar';

const LOG_SOURCE: string = 'CommandBarNavigationApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICommandBarNavigationApplicationCustomizerProperties {
  useTeamsites: boolean;
  // teamSitesMoreUrl: string;
  useCommsites: boolean;
  // commSitesMoreUrl: string;
  // useClassicsites: boolean;
  // classicSitesMoreUrl: string;
  // useHubsites: boolean;
  // hubSitesMoreUrl: string;
  // useTeams: boolean;
  // teamsMoreUrl: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class CommandBarNavigationApplicationCustomizer
  extends BaseApplicationCustomizer<ICommandBarNavigationApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized GlobalNavigation`);

    this.context.placeholderProvider.changedEvent.add(this, this.renderGlobalNavigation);
    
    this.renderGlobalNavigation();
    return Promise.resolve();
  }

  private renderGlobalNavigation() {
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, { onDispose: this._onDispose });
    
      if (!this._topPlaceholder) {
        console.error('The expected placeholder was not found!');
      }
      // Default handling config properties. ToDo: WHERE BETTER?
      if (typeof this.properties.useTeamsites !== 'boolean') {
        this.properties.useTeamsites = true;
      }
      if (typeof this.properties.useCommsites !== 'boolean') {
        this.properties.useCommsites = true;
      }
      // if (typeof this.properties.useClassicsites !== 'boolean') {
      //   this.properties.useClassicsites = true;
      // }
      // if (typeof this.properties.useHubsites !== 'boolean') {
      //   this.properties.useHubsites = true;
      // }
      // if (typeof this.properties.useTeams !== 'boolean') {
      //   this.properties.useTeams = true;
      // }
      if (this._topPlaceholder && this._topPlaceholder.domElement) {
        const element: React.ReactElement<ITopCommandBarProps> = React.createElement(
          TopCommandBar,
          {
            currentSiteUrl: this.context.pageContext.site.absoluteUrl,
            serviceScope: this.context.serviceScope,
            msGraphClientFactory: this.context.msGraphClientFactory,
            useGraph: true,               // ToDo
            useTeamsites: this.properties.useTeamsites,
            useCommsites: this.properties.useCommsites,
            // useClassicsites: this.properties.useClassicsites,
            // useHubsites: this.properties.useHubsites,
            // useTeams: this.properties.useTeams,
            // teamSitesMoreUrl: this.properties.teamSitesMoreUrl,
            // commSitesMoreUrl: this.properties.commSitesMoreUrl,
            // classicSitesMoreUrl: this.properties.classicSitesMoreUrl,  
            // hubSitesMoreUrl: this.properties.hubSitesMoreUrl,
            // teamsMoreUrl: this.properties.teamsMoreUrl
          }
        );
        ReactDom.render(element, this._topPlaceholder.domElement); 
      }
    }
  }

  private _onDispose(): void {
    console.log('Disposed custom top navigation placeholder.');
  }
}
