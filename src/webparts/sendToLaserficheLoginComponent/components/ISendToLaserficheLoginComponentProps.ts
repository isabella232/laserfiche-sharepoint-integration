import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISendToLaserficheLoginComponentProps {
  laserficheRedirectUrl: string;
  context: WebPartContext;
  devMode: boolean;
}
