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

  const getTeamsites = () => {
    if (props.useGraph) {
      graphService.readTeamsites('', 0).then((response: IMenuItem[]) => {
        setTeamsites(response);                   
      });;
    }
    else {
      spService.readTeamsites("", 0, props.currentSiteUrl).then((response: IMenuItem[]) => {
        setTeamsites(response);                   
      });
    }
  };

  const getCommsites = () => {
    if (props.useGraph) {
      graphService.readCommsites('', 0).then((response: IMenuItem[]) => {
        setCommsites(response);                   
      });;
    }
    else {
      spService.readCommsites("", 0, props.currentSiteUrl) 
        .then((response: IMenuItem[]) => {
          setCommsites(response);
        });
    }    
  };

  const getHubsites = () => {
    spService.getHubSiteId(props.currentSiteUrl).then((response: string|null) => {
      setHubsiteId(response);
      if (props.useGraph) {
        graphService.readHubsites('', 0).then((response: IMenuItem[]) => {
          setHubsites(response);                   
        });
      }
      else {
        spService.readHubsites('', 0, props.currentSiteUrl).then((response: IMenuItem[]) => {
          setHubsites(response);
        });
      }
    });    
  };

  const getTeams = () => { 
    graphService.getTopTeams() 
      .then((response: IMenuItem[]) => {
        setTeams(response);                   
      });        
  };

  const evalSharing = () => {
    spService.evalExternalSharingEnabled()
      .then((response: boolean) => {
        setExternalSharingEnabled(response);
      });
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
            <SitePermissions serviceScope={props.serviceScope} currentSiteUrl={props.currentSiteUrl} />
          </PivotItem>
          <PivotItem headerText="Lists">
            <ListPermissions serviceScope={props.serviceScope} currentSiteUrl={props.currentSiteUrl} />
          </PivotItem>
          <PivotItem headerText="Sharing Links">
            <SharingLinks serviceScope={props.serviceScope} currentSiteUrl={props.currentSiteUrl} />
          </PivotItem>      
        </Pivot>        
      </Panel>
    </div>
  );
}