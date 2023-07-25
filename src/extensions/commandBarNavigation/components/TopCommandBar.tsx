import * as React from "react";
import { CommandBar } from '@fluentui/react/lib/CommandBar';
import { IContextualMenuItem } from '@fluentui/react/lib/ContextualMenu';
import { Panel } from '@fluentui/react/lib/Panel';
import styles from './TopCommandBar.module.scss';
import { ITopCommandBarProps } from "./ITopCommandBarProps";
import { IMenuItem } from "../../../models/IMenuItem";
import GraphService from "../../../services/GraphService";
import { SPService } from "../../../services/SPService";
import { evaluateCommandItems, evaluateFarItems } from "../../../services/MenuItemsService";
import { ListPermissions } from "./permissions/ListPermissions";

export const TopCommandBar: React.FC<ITopCommandBarProps> = (props) => {
  const [teamsites, setTeamsites] = React.useState<IMenuItem[]>([]);
  const [commsites, setCommsites] = React.useState<IMenuItem[]>([]);
  const [teams, setTeams] = React.useState<IMenuItem[]>([]);
  const [externalSharingEnabled, setExternalSharingEnabled] = React.useState<boolean>(false);
  const [commandItems, setCommandItems] = React.useState<IContextualMenuItem[]>([]);
  const [farItems, setFarItems] = React.useState<IContextualMenuItem[]>([]);
  const [permissionPanelOpen, setPermissionPanelOpen] = React.useState<boolean>(false);
  const spService = new SPService(props.serviceScope);
  const graphService = new GraphService(props.serviceScope);

  const getTeamsites = () => { 
    spService.readTeamsites("", 0, props.currentSiteUrl) 
      .then((response: IMenuItem[]) => {
        setTeamsites(response);                   
      });        
  };

  const getCommsites = () => {    
    spService.readCommsites("", 0, props.currentSiteUrl) 
      .then((response: IMenuItem[]) => {
        setCommsites(response);                   
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
    const renderedItems = evaluateCommandItems(teamsites, commsites, teams);
    setCommandItems(renderedItems);    
  }, [teamsites, commsites, teams]);

  React.useEffect((): void => {
    const rightItems = evaluateFarItems(externalSharingEnabled, togglePermissions);
    setFarItems(rightItems);
  }, [externalSharingEnabled]);

  React.useEffect((): void => {
    getTeamsites();
    getCommsites();
    getTeams();
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
        <ListPermissions serviceScope={props.serviceScope} currentSiteUrl={props.currentSiteUrl} />
      </Panel>
    </div>
  );
}