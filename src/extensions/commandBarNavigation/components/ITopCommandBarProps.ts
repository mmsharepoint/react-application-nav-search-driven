import { ServiceScope } from "@microsoft/sp-core-library";
import { MSGraphClientFactory } from "@microsoft/sp-http";

export interface ITopCommandBarProps {
  currentSiteUrl: string;
  serviceScope: ServiceScope;
  msGraphClientFactory: MSGraphClientFactory;
  useTeamsites: boolean;
  useCommsites: boolean;
}