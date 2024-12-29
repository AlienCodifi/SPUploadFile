import { SPHttpClient } from "@microsoft/sp-http";

export interface IUploadAttachmentProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  absoluteURL:string;
}


export interface IUploadAttachmentControlProps {
  absoluteURL:string;
  spHttpClient: SPHttpClient;
}

