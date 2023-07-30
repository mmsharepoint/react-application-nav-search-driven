import * as React from "react";
import { IconButton } from '@fluentui/react/lib/Button';
import { List } from '@fluentui/react/lib/List';
import { IIconProps } from '@fluentui/react';
import { SPService } from "../../../../services/SPService";
import { IPermissionItem } from "../../../../models/IPermissionItem";
import styles from './SitePermissions.module.scss';
import { ISitePermissionsProps } from "./ISitePermissionsProps";

export const SitePermissions: React.FC<ISitePermissionsProps> = (props) => {
  const [items, setItems] = React.useState<IPermissionItem[]>([]);
  const cancelBtn: IIconProps = { iconName: 'Cancel' };
  const spService = new SPService(props.serviceScope);
  
  const evalSitePermissions = async (): Promise<void> => {
    const respItems = await spService.getSitePermissions(props.currentSiteUrl);        
    setItems(respItems);
  };

  const onRenderCell = (item: IPermissionItem, index: number): JSX.Element => {
    return (
      <div data-is-focusable={true}>       
        <div className={index! % 2 === 0 ? styles.evenRow : styles.oddRow}>
          <div className={styles.itemName}><a href={item.url}>{item.name}</a></div>
          <div className={styles.itemIndex}>
            <span>{item.permission}</span>
            {!item.isDefault && props.isSiteOwner &&
            <span>
              <IconButton iconProps={ cancelBtn } title='Remove permission'  />
            </span>}
          </div>
        </div>
      </div>
    );
  };

  React.useEffect((): void => {
    evalSitePermissions();
  }, []);

  return (
    <div className={styles.sitePermissions}>
      <h4>Site</h4>
      <List items={items} onRenderCell={onRenderCell} />
    </div>
  )
}