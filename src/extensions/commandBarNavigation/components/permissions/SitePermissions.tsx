import * as React from "react";
import { List } from '@fluentui/react/lib/List';
import { SPService } from "../../../../services/SPService";
import { IPermissionItem } from "../../../../models/IPermissionItem";
import styles from './SitePermissions.module.scss';
import { ISitePermissionsProps } from "./ISitePermissionsProps";

export const SitePermissions: React.FC<ISitePermissionsProps> = (props) => {
  const [items, setItems] = React.useState<IPermissionItem[]>([]);
  const spService = new SPService(props.serviceScope);
  
  const evalSitePermissions = () => {
    spService.getSitePermissions(props.currentSiteUrl)
      .then((respItems) => {
        console.log(respItems);
        setItems(respItems);
      });
  };

  const onRenderCell = (item: IPermissionItem, index: number): JSX.Element => {
    return (
      <div data-is-focusable={true}>       
        <div className={index! % 2 === 0 ? styles.evenRow : styles.oddRow}>
          <div className={styles.itemName}><a href={item.url}>{item.name}</a></div>
          <div className={styles.itemIndex}>{item.permission}</div>
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