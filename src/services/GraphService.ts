import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { MSGraphClientFactory, MSGraphClientV3 } from "@microsoft/sp-http";
import { ITeam } from "../models/ITeam";
import { IMenuItem } from "../models/IMenuItem";

export default class GraphService {
	private msGraphClientFactory: MSGraphClientFactory;
  private client: MSGraphClientV3;

  public static readonly serviceKey: ServiceKey<GraphService> =
    ServiceKey.create<GraphService>('react-application-nav-graph', GraphService);

  constructor(serviceScope: ServiceScope) {  
    serviceScope.whenFinished(async () => {
      this.msGraphClientFactory = serviceScope.consume(MSGraphClientFactory.serviceKey);      
    });
  }

  public async readTeamsites(searchText: string, start: number): Promise<IMenuItem[]> {
    let queryText = `WebTemplate:Group`;
    if (searchText !== null && searchText !== '') {
      queryText += ` AND ${searchText}`;
    }
    const searchResponse = await this.searchSites(queryText, start);    
    return this.transformSearchSites(searchResponse);
  }

  public async readCommsites(searchText: string, start: number): Promise<IMenuItem[]> {
    let queryText = `WebTemplate:SITEPAGEPUBLISHING`;
    if (searchText !== null && searchText !== '') {
      queryText += ` AND ${searchText}`;
    }
    const searchResponse = await this.searchSites(queryText, start);    
    return this.transformSearchSites(searchResponse);
  }

  private async searchSites(queryText: string, start: number): Promise<any> {
    this.client = await this.msGraphClientFactory.getClient('3');
    const reqeustBody = {
      requests: [
          {
              entityTypes: [
                  "site"
              ],
              query: {
                  "queryString": `${queryText}`
              }
          }
      ]
    };

    const response = await this.client
            .api(`search/query`)
            .version('v1.0')
            .skip(start)
            .top(20)   // Limit in batching!      
            .post(reqeustBody);
    return response.value[0].hitsContainers[0].hits;
  }

  private transformSearchSites(response: any[]): IMenuItem[] {    
    const items: Array<IMenuItem> = new Array<IMenuItem>();
    if (response !== null && response.length > 0) {
      response.forEach((r: any) => {          
        items.push({ displayName: r.resource.displayName, url: r.resource.webUrl, iconUrl: '', description: r.resource.description, key: r.resource.id });        
      });
      console.log(response);
      console.log(items);
      return items;
    }
    else {
      return [];
    }
  }

  public async getTopTeams(): Promise<IMenuItem[]> {
    const rawTeams = await this.getTeams();
    const teamsMenuItems = await this.transformTeams(rawTeams);
    return teamsMenuItems;
  }
  /**
  * This function retrievs the user's membership teams from Graph
  */
  private async getTeams(): Promise<ITeam[]> {
    this.client = await this.msGraphClientFactory.getClient('3');
    const response = await this.client
            .api(`groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$expand=members`) // /me/joinedTeams
            .version('v1.0')
            .top(20)   // Limit in batching!      
            .get();
    let teams: Array<ITeam> = new Array<ITeam>();
    response.value.forEach((item: ITeam|any) => {
      teams.push({ id: item.id, displayName: item.displayName, description: item.description, mailNickname: item.mailNickname, visibility: '', url: `/teams/${item.mailNickname}`, photo: '', members: item.members.length, createdDateTime: '', teamUrl: '', owners: []});
    });      
    return teams;
  }

  /**
   * This functions transforms a standard SP search result to custom data model
   * @param response: ISharePointSearchResults The standard SP search result
   */
  private transformTeams(teams: ITeam[]): IMenuItem[] {    
    const items: Array<IMenuItem> = new Array<IMenuItem>();
    if (teams !== null){
      teams.forEach((t: ITeam) => {          
        items.push({ displayName: t.displayName, url: t.url, iconUrl: '', description: t.description, key: t.id });        
      });
      return items;
    }
    else {
      return [];
    }
  }
}