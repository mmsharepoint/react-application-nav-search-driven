import { ServiceScope } from "@microsoft/sp-core-library";

export interface ISharingLinksProps {
    currentSiteUrl: string;
    serviceScope: ServiceScope;
  }