import * as React from "react";
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { IIconProps } from '@fluentui/react';
import { List } from '@fluentui/react/lib/List';
import { SPService } from "../../../../services/SPService";
import { IPermissionItem } from "../../../../models/IPermissionItem";
import styles from './ListPermissions.module.scss';
import { IListPermissionsProps } from "./IListPermissionsProps";

export const ListPermissions: React.FC<IListPermissionsProps> = (props) => {
  const [items, setItems] = React.useState<IPermissionItem[]>([]);
  const spService = new SPService(props.serviceScope);
  const [dialog, setDialog] = React.useState<JSX.Element>();
  const redoBtn: IIconProps = { iconName: 'Redo' };
  const cancelBtn: IIconProps = { iconName: 'Cancel' };
  
  const dialogContentProps = {
    type: DialogType.normal,
    title: 'Confirm Permission Change',
    closeButtonAriaLabel: 'Close',
    subText: '',
  };

  const evalListPermissions = () => {
    spService.evalSiteListsPermInheritance(props.currentSiteUrl)
      .then((respItems) => {
        setItems(respItems);
      });
  };

  const breakPermissionInheritance = (listID: string) => { 
    hideDialog();   
    spService.breakInheritListPermissions(props.currentSiteUrl, listID).then((resp) => {
      if (resp) {
        evalListPermissions();
      }
      else {
        // error
      }
    }); 
  };

  const hideDialog = () => {
    setDialog(<React.Fragment></React.Fragment>);
  };

  const confirmBreakPermissions = React.useCallback((listID: string) => {
    dialogContentProps.subText = 'Do you really want to break inherited list permissions?'
    setDialog(<Dialog
              hidden={false}
              onDismiss={hideDialog}
              dialogContentProps={dialogContentProps}
            >
              <DialogFooter>
                <PrimaryButton onClick={() => breakPermissionInheritance(listID)} text="OK" />
                <DefaultButton onClick={hideDialog} text="Cancel" />
              </DialogFooter>
            </Dialog>);
  }, [items]);

  const inheritPermissions = (listID: string) => {
    hideDialog();
    spService.reInheritListPermissions(props.currentSiteUrl, listID).then((resp) => {
      if (resp) {
        evalListPermissions();
      }
      else {
        // error
      }
    });
  };

  const confirmInheritPermissions = React.useCallback((listID: string) => {
    dialogContentProps.subText = 'Do you really want to re-inherit list permissions?'
    setDialog(<Dialog
              hidden={false}
              onDismiss={hideDialog}
              dialogContentProps={dialogContentProps}
            >
              <DialogFooter>
                <PrimaryButton onClick={() => inheritPermissions(listID)} text="OK" />
                <DefaultButton onClick={hideDialog} text="Cancel" />
              </DialogFooter>
            </Dialog>);
  }, [items]);

  const onRenderCell = (item: IPermissionItem, index: number | undefined): JSX.Element => {
    return (
      <div data-is-focusable={true}>       
        <div className={index! % 2 === 0 ? styles.evenRow : styles.oddRow}>
          <div className={styles.itemName}><a href={item.url}>{item.name}</a></div>
          <div className={styles.itemPermission}>
            <span className={styles.txtPermission}>{item.permission}</span>
            {item.permission === "Unique"?
              <IconButton iconProps={ redoBtn } title='Re-inherit permissions' onClick={ () => confirmInheritPermissions(item.key) } />:
              <IconButton iconProps={ cancelBtn } title='Stop inherit permissions' onClick={ () => confirmBreakPermissions(item.key) } />}
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

      {dialog}
    </div>
  )
}