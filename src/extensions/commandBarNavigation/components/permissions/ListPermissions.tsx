import * as React from "react";
import { List } from '@fluentui/react/lib/List';
import { SPService } from "../../../../services/SPService";
import { IPermissionItem } from "../../../../models/IPermissionItem";
import styles from './ListPermissions.module.scss';
import { IListPermissionsProps } from "./IListPermissionsProps";

export const ListPermissions: React.FC<IListPermissionsProps> = (props) => {
  const [items, setItems] = React.useState<IPermissionItem[]>([]);
  const spService = new SPService(props.serviceScope);
  const evalListPermissions = () => {
    spService.evalSiteListsPermInheritance(props.currentSiteUrl)
      .then((respItems) => {
        console.log(respItems);
        setItems(respItems);
      });
  };

  const onRenderCell = (item: IPermissionItem, index: number | undefined): JSX.Element => {
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
    evalListPermissions();
  }, []);
  return (
    <div className={styles.listPermissions}>
      <h4>Lists</h4>
      <List items={items} onRenderCell={onRenderCell} />
    </div>
  )
}