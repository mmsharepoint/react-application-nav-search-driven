import * as React from 'react';
import { ICommandBarItemProps } from '@fluentui/react/lib/CommandBar';
import { IContextualMenuItem } from '@fluentui/react/lib/ContextualMenu';
import { IMenuItem } from '../models/IMenuItem';
import { SubMenuSites } from '../extensions/commandBarNavigation/components/subMenu/SubMenuSites';
import { ISubMenuSitesProps } from '../extensions/commandBarNavigation/components/subMenu/ISubMenuSitesProps';


const renderSubmenu = (item: any) => {    
  let element: React.ReactElement<ISubMenuSitesProps> = React.createElement(SubMenuSites, item);    
  return element;
}

const dismissProjects = () => {
  // What if submenu gets closed?
}

const closeMenu = () => {
  return true;
}

const teamSitesItem: ICommandBarItemProps = {
  key: 'Teamsites',
  name: 'Teamsites',
  className: `ms-CommandBarItem`,
  iconProps: {
    iconName: 'Group'
  },
  subMenuProps: {
    items: [{
      key: 'Teamsites',
      name: 'Teamsites',   
      label: 'Modern Teamsites',     
      dataItems: [],
      // searchCallback: this.searchTeamsites.bind(this),
      onRender: renderSubmenu
    }],
    onDismiss: dismissProjects,
    onItemClick: closeMenu                
  }
  };

  const commSitesItem: ICommandBarItemProps = {
    key: 'Commsites',
    name: 'Communication Sites',
    className: `ms-CommandBarItem`,
    iconProps: {
      iconName: 'Communications'
    },
    subMenuProps: {
      items: [{
        key: 'ProjectSub',
        name: 'ProjectSub', 
        label: 'Modern Communication Sites',       
        dataItems: [],
        //searchCallback: this.searchCommSites.bind(this),
        onRender: renderSubmenu
      }],
      onDismiss: dismissProjects,
      onItemClick: closeMenu                 
    }
  };

  const externalSharingItem: ICommandBarItemProps = {
    key: 'ExtNo',
    title : '',
    iconProps: {
      iconName: 'Info'
    },
    iconOnly: true
  };

  const permissionItem: ICommandBarItemProps = {
    key: 'permission',
    name: 'Permissions',
    iconProps: {
      iconName: 'Repair'
    },
    iconOnly: true,
    subMenuProps: {
      items: []
    }    
  };

  const permissionPanelItem: IContextualMenuItem = {    
    key: 'ListPermissions',
    name: 'List Permissions',
  };

  export const evaluateCommandItems = (teamsites: IMenuItem[], commsites: IMenuItem[]): ICommandBarItemProps[] => {    
    teamSitesItem.subMenuProps!.items[0].dataItems = teamsites;
    commSitesItem.subMenuProps!.items[0].dataItems = commsites;
    let commandBarItems: ICommandBarItemProps[] = [];
    // if (this.useTeamsites) {      
      commandBarItems.push(teamSitesItem);
    // }
    // if (this.useCommsites) {      
      commandBarItems.push(commSitesItem);
    // }
    
    return commandBarItems;
  }

  export const evaluateFarItems = (externalSharingEnabled: boolean, showPermissions: () => void): ICommandBarItemProps[] => {
    let farItems: ICommandBarItemProps[] = [];
    if (externalSharingEnabled !== null) {
      if (externalSharingEnabled) {
        externalSharingItem.title = 'External Sharing enabled';
        externalSharingItem.iconProps!.iconName = 'World';
      }
      else {
        externalSharingItem.title = 'External Sharing not enabled';
        externalSharingItem.iconProps!.iconName = 'LifesaverLock';
      }
      farItems.push(externalSharingItem);      
    }
    permissionItem.subMenuProps!.items = [];
    permissionPanelItem.onClick = () => { showPermissions(); };
    permissionItem.subMenuProps?.items.push(permissionPanelItem);
    farItems.push(permissionItem);
    return farItems;
  }