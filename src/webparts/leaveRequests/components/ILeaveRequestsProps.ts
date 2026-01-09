
import{SPHttpClient} from '@microsoft/sp-http'
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ViewType } from './ViewType';
export interface ILeaveRequestsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  siteUrl: string;
  spHttpClient: SPHttpClient;
  context:WebPartContext;
  onViewChange:(view:ViewType.myLeaves |ViewType.myApproval)=> void;
  isManagerView:boolean;
}
