import * as React from "react";
import { TextField } from '@fluentui/react/lib/TextField';
import styles from './SubMenuSites.module.scss';
import { ISubMenuSitesProps } from "./ISubMenuSitesProps";
import { SiteIcon } from "../siteIcon/SiteIcon";

export const SubMenuSites: React.FC<ISubMenuSitesProps> = (props) => {
  let listElements: JSX.Element[] = [];
  if (props.dataItems.length > 0) {
    listElements = props.dataItems.map((item) => {
      return (<li className={styles.menuListItem}>
                <SiteIcon
                  siteTitle={item.displayName}
                  iconAcronym={item.iconAcronym?item.iconAcronym:''}
                  iconColor={item.iconColor?item.iconColor:''}
                  iconUrl={item.iconUrl?item.iconUrl:''} />
                <a className="ms-fontColor-neutralPrimary ms-fontColor-themePrimary--hover" onClick={() => window.open(item.url , '_self')} >{item.displayName}</a>
              </li>);
    });
  }
  else {
    
  }
  return (
    <div className={styles.subMenu}>
      <div className="ms-Grid-row">
        <div className="ms-Grid-col ms-sm11 ms-md7 ms-lg7 ms-xl7 ms-xxl7 ms-xxxl7">
          <TextField
            inputClassName="ms-bgColor-themeLighter ms-fontColor-themePrimary"
            borderless
            placeholder={`Search ${props.label}`}
            // iconClass="ms-fontColor-themePrimary"
            iconProps={ { iconName: 'Search' } }
            // value={this.state.searchInputText}
            // onChanged={this.handleSearchTextChange}
            // onKeyDown={this.searchSites.bind(this)}              
          />
        </div>
      </div>        
      <div className="ms-Grid-row">
        <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6 ms-xl6 ms-xxl6 ms-xxxl6">
          <div className={styles.menuHeader}>
            <span className="ms-font-xl ms-fontColor-themePrimary">{props.label}</span>
          </div>            
          <ul className={styles.menuList}>{listElements}</ul>
        </div>          
      </div>
    </div>
  );
}