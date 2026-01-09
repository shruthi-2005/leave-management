// src/webparts/leaveRequests/components/IApplyLeaveProps.ts

import { ViewType } from "./ViewType";

export interface IApplyLeaveProps {
  item: {
    Id: number;
    EmployeeName: string;
    LeaveType: string;
    StartDate: string;
    EndDate: string;
    Days: number;
    Status: string;
    Reason: string;
    Manager: any;
    RejectReason: string;
  
  };
  viewType: ViewType;
  onViewChange: (view: ViewType) => void;
  context: any;
  formatDate: (date: string) => string;
  isManagerView:(dataString:string)=>string;
  
}