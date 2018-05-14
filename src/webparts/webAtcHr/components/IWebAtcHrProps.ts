import { SPHttpClient } from '@microsoft/sp-http';
import { SharePointUserPersona } from '../models/OfficeUiFabricPeoplePicker';
export interface IWebAtcHrProps {
  description: string;
  PassportRequest:number;
  LeaveRequest:number;
  FormIsEnabled:number;
  RequestTypeString:string;
  spHttpClient: SPHttpClient;
  currentPicker: number,
  delayResults: boolean,
  selectedItems: Array<string>[];
  descriptionpicker: string;
  siteUrlpicker: string;
  typePicker: string;
  principalTypeUser: boolean;
  principalTypeSharePointGroup: boolean;
  principalTypeSecurityGroup: boolean;
  principalTypeDistributionList: boolean;
  numberOfItems: number;
  onChange?: (items: SharePointUserPersona[]) => void;
  siteUrl:string;
  
}
