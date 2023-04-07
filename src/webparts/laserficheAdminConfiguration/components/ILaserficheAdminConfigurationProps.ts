import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ILaserficheAdminConfigurationProps {
  webPartTitle: string;
  laserficheRedirectPage: string;
  context: WebPartContext;
  devMode: boolean;
}
