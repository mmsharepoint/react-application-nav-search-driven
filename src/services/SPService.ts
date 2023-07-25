import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { SPHttpClient, SPHttpClientResponse, ODataVersion, ISPHttpClientConfiguration } from '@microsoft/sp-http';
import { ISharePointSearchResults, ISharePointSearchResultsTable } from "../models/ISearchHandler";
import { IMenuItem } from "../models/IMenuItem";
import { IPermissionItem } from "../models/IPermissionItem";

export interface ISPService {
  readTeamsites(searchText: string, start: number, currentSiteUrl: string): Promise<IMenuItem[]>;
}

export class SPService implements ISPService {
  public static readonly serviceKey: ServiceKey<SPService> =
    ServiceKey.create<SPService>('react-application-nav', SPService);
  private _spHttpClient: SPHttpClient;
  private currentSiteUrl: string;
  private pagesize: number = 15;

  constructor(serviceScope: ServiceScope) {  
    serviceScope.whenFinished(() => {
      this._spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
      // this.graphServiceInstance = serviceScope.consume(GraphService.serviceKey);
    });
  }

  /**
   * This functions executes a SP search for sites and returns a non transformed search result
   * @param queryText Keqword query text for site search
   */
  private searchSites(queryText: string, start: number): Promise<ISharePointSearchResults> {
    const spSearchConfig: ISPHttpClientConfiguration = {
      defaultODataVersion: ODataVersion.v3
    };
    const requestUrl = this.currentSiteUrl + `/_api/search/query?querytext='${queryText}'&selectproperties='Title,Path,SiteLogo,Description,SiteId'&startrow=${start}&rowlimit=${this.pagesize}&sortlist='Created:descending'`;
    return this._spHttpClient.get(requestUrl, SPHttpClient.configurations.v1.overrideWith(spSearchConfig))
        .then((response: SPHttpClientResponse) => {
            return response.json();
        });
  }

  /**
   * This functions transforms a standard SP search result to custom data model
   * @param response: ISharePointSearchResults The standard SP search result
   */
  private transformSearchSites(response:ISharePointSearchResults): IMenuItem[] {
    const relevantResults: ISharePointSearchResultsTable | null = (response.PrimaryQueryResult !== null) ? response.PrimaryQueryResult.RelevantResults.Table : null;     
    const items: Array<IMenuItem> = new Array<IMenuItem>();
    if (relevantResults !== null){
      relevantResults.Rows.forEach((r: any) => {          
        items.push({ displayName: r.Cells[0].Value, url: r.Cells[1].Value, iconUrl: r.Cells[2].Value, description: r.Cells[3].Value, key: r.Cells[4].Value });        
      });
      return items;
    }
    else {
      return [];
    }
  }
  
  public async readTeamsites(searchText: string, start: number, currentSiteUrl: string): Promise<IMenuItem[]> {
    let queryText = `contentclass:STS_Site AND WebTemplate:Group`;
    this.currentSiteUrl = currentSiteUrl;
    if (searchText !== null && searchText !== '') {
      queryText += ` AND ${searchText}`;
    }
    const searchResponse = await this.searchSites(queryText, start);
    return this.transformSearchSites(searchResponse);
  }

  public async readCommsites(searchText: string, start: number, currentSiteUrl: string): Promise<IMenuItem[]> {
    let queryText = `contentclass:STS_Site AND WebTemplate:SITEPAGEPUBLISHING`;
    this.currentSiteUrl = currentSiteUrl;
    if (searchText !== null && searchText !== '') {
      queryText += ` AND ${searchText}`;
    }
    const searchResponse = await this.searchSites(queryText, start);
    return this.transformSearchSites(searchResponse);
  }

  /**
   * This function evaluates if external sharing is enabled for the current site
   * @returns Promise<boolean> If external sharing enabled
   */
  public async evalExternalSharingEnabled(): Promise<boolean> {
    const requestUrl = this.currentSiteUrl + '/_api/site?$select=ShareByEmailEnabled';
    return this._spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((jsonResponse: any): boolean => {
        return jsonResponse.ShareByEmailEnabled;
      });
  }

  public async getSitePermissions(currentSiteUrl: string) {
    this.currentSiteUrl = currentSiteUrl;
    const requestUrl = this.currentSiteUrl + '/_api/web/roleassignments?$expand=Member/users,RoleDefinitionBindings';
    return this._spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((jsonResponse: any) => {
        const permissionItems: IPermissionItem[] = [];
        jsonResponse.value.forEach((l: any) => {
          permissionItems.push({ key: l.PrincipalId, name: l.Member.Title, permission: l.RoleDefinitionBindings[0].Name, description: l.RoleDefinitionBindings[0].Description, url: this.currentSiteUrl + `/_layouts/15/people.aspx?MembershipGroupId=${l.PrincipalId}` });
        });
        console.log(jsonResponse.value);
        return permissionItems;
      });
  }

  public async evalSiteListsPermInheritance(currentSiteUrl: string): Promise<IPermissionItem[]> {
    this.currentSiteUrl = currentSiteUrl;
    const requestUrl = this.currentSiteUrl + '/_api/web/lists?$select=HasUniqueRoleAssignments,Title,Id,BaseTemplate,RootFolder/ServerRelativeUrl&$expand=RootFolder&$filter=BaseTemplate eq 101 or BaseTemplate eq 100';
    return this._spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((jsonResponse: any) => {
        const permissionItems: IPermissionItem[] = [];
        jsonResponse.value.forEach((l: any) => {
          permissionItems.push({ key: l.Id, name: l.Title, permission: l.HasUniqueRoleAssignments ? 'Unique':'Inherits', description: '', url: l.RootFolder.ServerRelativeUrl });
        });
        return permissionItems;
      });
  }
}