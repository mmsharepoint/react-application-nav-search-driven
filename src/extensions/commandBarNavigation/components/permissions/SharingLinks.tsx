import * as React from "react";
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { IIconProps } from '@fluentui/react';
import { List } from '@fluentui/react/lib/List';
import { SPService } from "../../../../services/SPService";
import styles from './ListPermissions.module.scss';
import { ISharingLinksProps } from "./ISharingLinksProps";
import { ISharingLink } from "../../../../models/ISharingLink";

export const SharingLinks: React.FC<ISharingLinksProps> = (props) => {
  const [items, setItems] = React.useState<ISharingLink[]>([]);
  const spService = new SPService(props.serviceScope);
  const [dialog, setDialog] = React.useState<JSX.Element>();
  const cancelBtn: IIconProps = { iconName: 'Cancel' };
  
  const dialogContentProps = {
    type: DialogType.normal,
    title: 'Confirm Permission Change',
    closeButtonAriaLabel: 'Close',
    subText: '',
  };

  const getSharingLinks = () => {
    spService.getSharingLinks(props.currentSiteUrl)
      .then((respItems) => {
        setItems(respItems);
        console.log(respItems);
        // spService.enrichSharingLinksByDoc(props.currentSiteUrl, respItems)
        //   .then((respItems) => {
        //     setItems(respItems);
        //     console.log(respItems);
        //   });
      });
  };

  const breakPermissionInheritance = (listID: string) => { 
    hideDialog();   
    spService.breakInheritListPermissions(props.currentSiteUrl, listID).then((resp) => {
      if (resp) {
        getSharingLinks();
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

  const onRenderCell = (item: ISharingLink, index: number | undefined): JSX.Element => {
    return (
      <div data-is-focusable={true}>       
        <div className={index! % 2 === 0 ? styles.evenRow : styles.oddRow}>
          <div className={styles.itemName}><a href={item.url}>{item.name}</a></div>
          <div className={styles.itemPermission}>
            <span className={styles.txtPermission}>{item.role}</span>
              <IconButton iconProps={ cancelBtn } title='Stop sharing!' onClick={ () => confirmBreakPermissions(item.key) } />
          </div>
        </div>
      </div>
    );
  };

  React.useEffect((): void => {
    getSharingLinks();
  }, []);

  return (
    <div className={styles.listPermissions}>
      <h4>Shared Files</h4>
      <List items={items} onRenderCell={onRenderCell} />

      {dialog}
    </div>
  )
}