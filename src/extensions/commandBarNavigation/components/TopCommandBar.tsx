import * as React from "react";
import { CommandBar } from '@fluentui/react/lib/CommandBar';
import { IContextualMenuItem } from '@fluentui/react/lib/ContextualMenu';
import { Panel } from '@fluentui/react/lib/Panel';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import styles from './TopCommandBar.module.scss';
import { ITopCommandBarProps } from "./ITopCommandBarProps";
import { IMenuItem } from "../../../models/IMenuItem";
import GraphService from "../../../services/GraphService";
import { SPService } from "../../../services/SPService";
import { evaluateCommandItems, evaluateFarItems } from "../../../services/MenuItemsService";
import { ListPermissions } from "./permissions/ListPermissions";
import { SitePermissions } from "./permissions/SitePermissions";
import { SharingLinks } from "./permissions/SharingLinks";

export const TopCommandBar: React.FC<ITopCommandBarProps> = (props) => {
  const [teamsites, setTeamsites] = React.useState<IMenuItem[]>([]);
  const [commsites, setCommsites] = React.useState<IMenuItem[]>([]);
  const [hubites, setHubsites] = React.useState<IMenuItem[]>([]);
  const [teams, setTeams] = React.useState<IMenuItem[]>([]);
  const [hubsiteId, setHubsiteId] = React.useState<string|null>(null);
  const [homesite, setHomesite] = React.useState<{url:string,displayName:string}|null>(null);
  const [externalSharingEnabled, setExternalSharingEnabled] = React.useState<boolean>(false);
  const [commandItems, setCommandItems] = React.useState<IContextualMenuItem[]>([]);
  const [farItems, setFarItems] = React.useState<IContextualMenuItem[]>([]);
  const [permissionPanelOpen, setPermissionPanelOpen] = React.useState<boolean>(false);
  const spService = new SPService(props.serviceScope);
  const graphService = new GraphService(props.serviceScope);

  const getTeamsites = async (): Promise<void> => {
    if (props.useGraph) {
      const response: IMenuItem[] = await graphService.readTeamsites('', 0);
      setTeamsites(response);                   
    }
    else {
      const response: IMenuItem[] = await spService.readTeamsites("", 0, props.currentSiteUrl);
      setTeamsites(response);
    }
  };

  const getCommsites = async (): Promise<void> => {
    if (props.useGraph) {
      const response: IMenuItem[] = await graphService.readCommsites('', 0);
      setCommsites(response);
    }
    else {
      const response: IMenuItem[] = await spService.readCommsites("", 0, props.currentSiteUrl);
      setCommsites(response);
    }    
  };

  const getHubsites = async (): Promise<void> => {
    const response: string|null = await spService.getHubSiteId(props.currentSiteUrl);
    setHubsiteId(response);
    if (props.useGraph) {
      const response: IMenuItem[] = await graphService.readHubsites('', 0);
      setHubsites(response);                   
    }
    else {
      const response: IMenuItem[] = await spService.readHubsites('', 0, props.currentSiteUrl);
      setHubsites(response);
    }   
  };

  const getTeams = async (): Promise<void> => { 
    const response: IMenuItem[] = await graphService.getTopTeams();
    setTeams(response);
  };

  const evalSharing = async (): Promise<void> => {
    const response: boolean = await spService.evalExternalSharingEnabled(props.currentSiteUrl)
    setExternalSharingEnabled(response);
  };

  React.useEffect((): void => {
    const renderedItems = evaluateCommandItems(teamsites, commsites, hubites, teams, homesite, props.useTeamsites, props.useCommsites, props.useHubsites, props.useTeams);
    setCommandItems(renderedItems);    
  }, [teamsites, commsites, hubites, teams, homesite]);

  React.useEffect((): void => {
    const rightItems = evaluateFarItems(externalSharingEnabled, togglePermissions);
    setFarItems(rightItems);
  }, [externalSharingEnabled]);

  
  React.useEffect((): void => {
    if (props.useHubsites && hubsiteId !== null && hubites.length >0) {
      hubites.forEach((h) => {
        if (h.key.indexOf(hubsiteId) > -1) {
          setHomesite({ url: h.url, displayName: h.displayName});
        }
      })
    } 
  }, [hubsiteId, hubites]);

  React.useEffect((): void => {
    if (props.useTeamsites) {
      getTeamsites();
    }
    if (props.useCommsites) {
      getCommsites();
    }
    if (props.useHubsites) {
      getHubsites();
    }
    if (props.useTeams) {
      getTeams();
    }    
    evalSharing();
  }, []);

  const togglePermissions = React.useCallback(() => {
    setPermissionPanelOpen(!permissionPanelOpen);
  }, [permissionPanelOpen]);

  return (
    <div className={styles.app}>
      <CommandBar          
        className={styles.top}    
        items={ commandItems }
        farItems={ farItems }
      />
      <Panel
          headerText="Permissions"
          isOpen={permissionPanelOpen}
          onDismiss={togglePermissions}
          closeButtonAriaLabel="Close"
        >
        <Pivot aria-label="Basic Pivot Example">
          <PivotItem
            headerText="Site"
            headerButtonProps={{
              'data-order': 1,
              'data-title': 'My Files Title',
            }}
          >
            <SitePermissions serviceScope={props.serviceScope} currentSiteUrl={props.currentSiteUrl} isSiteOwner={props.isSiteOwner} />
          </PivotItem>
          <PivotItem headerText="Lists">
            <ListPermissions serviceScope={props.serviceScope} currentSiteUrl={props.currentSiteUrl} isSiteOwner={props.isSiteOwner} />
          </PivotItem>
          <PivotItem headerText="Sharing Links">
            <SharingLinks serviceScope={props.serviceScope} currentSiteUrl={props.currentSiteUrl} siteId={props.siteId} isSiteOwner={props.isSiteOwner} />
          </PivotItem>      
        </Pivot>        
      </Panel>
    </div>
  );
}