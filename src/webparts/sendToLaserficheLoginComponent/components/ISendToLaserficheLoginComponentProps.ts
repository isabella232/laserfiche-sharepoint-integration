import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISendToLaserficheLoginComponentProps {
  laserficheRedirectPage: string;
  context: WebPartContext;
  devMode: boolean;
}
