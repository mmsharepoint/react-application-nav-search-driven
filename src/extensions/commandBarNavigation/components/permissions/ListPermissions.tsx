import * as React from "react";
import { IconButton } from '@fluentui/react/lib/Button';
import { IIconProps } from '@fluentui/react';
import { List } from '@fluentui/react/lib/List';
import { SPService } from "../../../../services/SPService";
import { IPermissionItem } from "../../../../models/IPermissionItem";
import styles from './ListPermissions.module.scss';
import { IListPermissionsProps } from "./IListPermissionsProps";

export const ListPermissions: React.FC<IListPermissionsProps> = (props) => {
  const [items, setItems] = React.useState<IPermissionItem[]>([]);
  const spService = new SPService(props.serviceScope);
  const redoBtn: IIconProps = { iconName: 'Redo' };
  const cancelBtn: IIconProps = { iconName: 'Cancel' };
  
  const evalListPermissions = () => {
    spService.evalSiteListsPermInheritance(props.currentSiteUrl)
      .then((respItems) => {
        setItems(respItems);
      });
  };

  const breakPermissionInheritance = React.useCallback((listID: string) => {    
    spService.breakInheritListPermissions(props.currentSiteUrl, listID).then((resp) => {
      if (resp) {
        evalListPermissions();
      }
      else {
        // error
      }
    }); 
  }, [items]);

  const inheritPermissions = React.useCallback((listID: string) => {
    spService.reInheritListPermissions(props.currentSiteUrl, listID).then((resp) => {
      if (resp) {
        evalListPermissions();
      }
      else {
        // error
      }
    });
  }, [items]);

  const onRenderCell = (item: IPermissionItem, index: number | undefined): JSX.Element => {
    return (
      <div data-is-focusable={true}>       
        <div className={index! % 2 === 0 ? styles.evenRow : styles.oddRow}>
          <div className={styles.itemName}><a href={item.url}>{item.name}</a></div>
          <div className={styles.itemPermission}>
            <span className={styles.txtPermission}>{item.permission}</span>
            {item.permission === "Unique"?
              <IconButton iconProps={ redoBtn } title='Re-inherit permissions' onClick={ () => inheritPermissions(item.key) } />:
              <IconButton iconProps={ cancelBtn } title='Stop inherit permissions' onClick={ () => breakPermissionInheritance(item.key) } />}
          </div>
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