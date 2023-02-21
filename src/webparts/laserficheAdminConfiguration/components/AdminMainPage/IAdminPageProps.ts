import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IAdminPageProps {
    context:WebPartContext;
    webPartTitle:string;
    laserficheRedirectPage:string;
    region:string;
  }
  