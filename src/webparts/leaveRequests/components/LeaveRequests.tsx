import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import ApplyLeave from './ApplyLeave';
import { ViewType } from './ViewType';

import * as $ from 'jquery';
import 'datatables.net-bs5';
import 'datatables.net-bs5/css/dataTables.bootstrap5.min.css';

export interface ILeaveRequestsProps {
  context: WebPartContext;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  viewType: ViewType;
  onViewChange: (view: ViewType) => void;
  onSelectItem: (item: any) => void;
}

export interface ILeaveRequestItem {
  Id: number;
  EmployeeName: string;
  LeaveType: string;
  StartDate: string;
  EndDate: string;
  Days: number;
  Status: string;
  Reason: string;
  Manager: {
    EMail?: string;
    Title?: string;
  };
}

const hardcodedSiteUrl = "https://elevix.sharepoint.com/sites/Trainingportal";

const LeaveRequests: React.FC<ILeaveRequestsProps> = (props) => {
  const [allRequests, setAllRequests] = useState<ILeaveRequestItem[]>([]);
  const [selectedItem, setSelectedItem] = useState<ILeaveRequestItem | null>(null);
  const [statusFilter, setStatusFilter] = useState<string>("");
  const [reloadFlag, setReloadFlag] = useState<boolean>(false);

  const tableRef = useRef<HTMLTableElement>(null);
  const dataTableInstance = useRef<any>(null);

  // 🔹 Fetch data whenever viewType or reloadFlag changes
  useEffect(() => {
    const fetchData = async () => {
      try {
        const userEmail = props.context.pageContext.user.email.toLowerCase();
        const displayName = props.context.pageContext.user.displayName;

        let baseUrl = `${hardcodedSiteUrl}/_api/web/lists/getbytitle('LeaveRequests')/items?$filter=`;

        if (props.viewType === 'myLeaves') {
          baseUrl += `EmployeeName eq '${displayName}'`;
        } else if (props.viewType === 'myApproval') {
          baseUrl += `Manager/EMail eq '${userEmail}'`;
        }

        const selectExpand = "&$select=Id,EmployeeName,LeaveType,StartDate,EndDate,Days,Status,Reason,Manager/Title,Manager/EMail&$expand=Manager&$orderby=Id desc";
        const requestUrl = baseUrl + selectExpand;

        console.log("📡 Fetching from:", requestUrl);

        const response = await props.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1);
        const data = await response.json();

        if (response.ok) {
          console.log("✅ Fetched Leave Requests:", data.value.length);
          setAllRequests(data.value || []);
        } else {
          console.error("❌ Error response:", response.statusText);
          setAllRequests([]);
        }
      } catch (error) {
        console.error("❌ Error fetching leave requests:", error);
        setAllRequests([]);
      }
    };

    fetchData();
  }, [props.viewType, reloadFlag]);

  // 🔹 Initialize / re-initialize DataTable safely
  useEffect(() => {
    if (!tableRef.current) return;

    // Destroy previous instance if exists
    if (dataTableInstance.current) {
      dataTableInstance.current.destroy();
    }

    dataTableInstance.current = ($(tableRef.current) as any).DataTable({
    paging: true,
    searching: true,
    ordering: true, 
   
    order: [], 
    autoWidth: false,
  });
}, [allRequests]);
  
  useEffect(() => {
    if (!dataTableInstance.current) return;

    dataTableInstance.current.clear();

    const filtered = statusFilter === ""
      ? allRequests
      : allRequests.filter(r => r.Status?.trim().toLowerCase() === statusFilter.toLowerCase());

    filtered.forEach(req => {
      dataTableInstance.current.row.add([
        req.EmployeeName,
        req.LeaveType,
        formatDate(new Date(req.StartDate)),
        formatDate(new Date(req.EndDate)),
        req.Days,
        req.Status,
        req.Manager?.Title || '-'
      ]);
    });

    dataTableInstance.current.draw();
  }, [allRequests, statusFilter]);
 
  const formatDate = (date: Date): string => {
    const day = date.getDate();
    const month = date.getMonth() + 1;
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
  };


  if (selectedItem) {
    const status = selectedItem.Status?.trim().toLowerCase();
    const isEmployee = props.viewType === ViewType.myLeaves;
    const isManager = props.viewType === ViewType.myApproval;

    const showCancelButton = isEmployee && status === "pending";
    const showManagerActions = isManager && status === "pending";

    return (
      <ApplyLeave
        context={props.context}
        item={selectedItem}
        spHttpClient={props.spHttpClient}
        siteUrl={hardcodedSiteUrl}
        onViewChange={props.onViewChange}
        viewType={props.viewType}
        sourceView={props.viewType}
        showCancelButton={showCancelButton}
        showManagerActions={showManagerActions}
        onBack={() => { 
          setSelectedItem(null);
          setReloadFlag(prev => !prev); // 🔹 Trigger refetch
        }}
      />
    );
  }

  // --- Main Table Screen ---
  return (
    <div className="container-fluid mt-3 px-2">
      <div className="d-flex justify-content-between align-items-center mb-3 flex-wrap">
        {props.viewType === 'myApproval' && (
          <div className="mb-2">
            <label className="me-2 fw-semibold">Status:</label>
            <select
              className="form-select d-inline-block"
              style={{ width: 'auto', minWidth: '120px' }}
              value={statusFilter}
              onChange={e => setStatusFilter(e.target.value)}
            >
              <option value="">All</option>
              <option value="Pending">Pending</option>
              <option value="Approved">Approved</option>
              <option value="Rejected">Rejected</option>
            </select>
          </div>
        )}
      </div>

      <div className="table-responsive shadow-sm rounded-3" style={{ overflowX: 'auto' }}>
        <table
          className="table table-striped table-bordered align-middle mb-0"
          style={{ width: "100%", cursor: 'pointer', fontSize: '0.9rem' }}
          ref={tableRef}
          onClick={(e) => {
            const target = e.target as HTMLElement;
            const tr = target.closest('tr');
            if (!tr) return;

            const rowIndex = dataTableInstance.current.row(tr).index();
            const filtered = statusFilter === ""
              ? allRequests
              : allRequests.filter(r => r.Status?.trim().toLowerCase() === statusFilter.toLowerCase());

            const dataItem = filtered[rowIndex];
            if (dataItem) setSelectedItem(dataItem);
          }}
        >
          <thead className="table-light">
            <tr>
              <th>Employee</th>
              <th>Leave Type</th>
              <th>Start Date</th>
              <th>End Date</th>
              <th>Days</th>
              <th>Status</th>
              <th>Manager</th>
            </tr>
          </thead>
          <tbody></tbody>
        </table>
      </div>
    </div>
  );
};

export default LeaveRequests;