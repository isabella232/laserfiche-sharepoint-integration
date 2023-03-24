import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IAddNewManageConfigurationProps {
    context:WebPartContext;
    laserficheRedirectPage:string;
    devMode: boolean;
  }