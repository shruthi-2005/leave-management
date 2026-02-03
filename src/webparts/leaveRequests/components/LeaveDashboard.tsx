import * as React from 'react';
import { useEffect, useState } from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import Leaves from './Leaves';
import LeaveRequests from './LeaveRequests';
import ApplyLeave from './ApplyLeave';
import { ViewType } from './ViewType';
import 'bootstrap/dist/css/bootstrap.min.css';

export interface ILeaveDashboardProps {
  context: WebPartContext;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  onViewChange: (view: ViewType) => void;
}

interface ILeaveInfo {
  Id: number;
  EmployeeName: string;
  LeaveType: string;
  TotalLeaves: number;
  RemainingLeaves: number;
  LeaveTaken: number;
}

const LeaveDashboard: React.FC<ILeaveDashboardProps> = ({
  context,
  spHttpClient,
  siteUrl,
  onViewChange,
}) => {
  const [userName, setUserName] = useState('');
  const [leaveData, setLeaveData] = useState<ILeaveInfo[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [currentView, setCurrentView] = useState<ViewType>(ViewType.home);

  const [isDeeplinkLoading, setIsDeeplinkLoading] = useState(true);
  const [selectedItem, setSelectedItem] = useState<any>(null);

  // Manager filter states
  const [isManager, setIsManager] = useState(false);
  const [assignedEmployees, setAssignedEmployees] = useState<string[]>([]);
  const [selectedEmployee, setSelectedEmployee] = useState<string>(''); // Empty → means show manager’s own balance
const [sourceView, setSourceView] = useState<ViewType>(ViewType.myLeaves);
  // 🔹 Hardcoded site URL for Teams
  const hardcodedSiteUrl = "https://elevix.sharepoint.com/sites/Trainingportal";
useEffect(() => {
  const teamsContext = context.sdks.microsoftTeams?.context;

  // ✅ CASE 1: SharePoint WebPart (NO Teams at all)
  if (!teamsContext) {
    setCurrentView(ViewType.home);
    setIsDeeplinkLoading(false);
    return;
  }

  // ✅ CASE 2: Teams open WITHOUT deeplink
  if (!teamsContext.subEntityId) {
    setCurrentView(ViewType.home);
    setIsDeeplinkLoading(false);
    return;
  }

  // 🔥 CASE 3: Teams DEEPLINK
  const itemId = Number(teamsContext.subEntityId);
  console.log("TEAMS subEntityId:", itemId);

  spHttpClient
    .get(
      `${hardcodedSiteUrl}/_api/web/lists/getbytitle('LeaveRequests')/items(${itemId})
?$select=Id,EmployeeName,LeaveType,StartDate,EndDate,Days,Status,Reason,Manager/Title
&$expand=Manager`,
      SPHttpClient.configurations.v1
    )
    .then(res => res.json())
    .then(data => {
      setSelectedItem(data);

      const status = data.Status?.trim().toLowerCase();
      const isSelfApproval =
        data.EmployeeName?.trim().toLowerCase() ===
        data.Manager?.Title?.trim().toLowerCase();

      setSourceView(
        status === "pending" && (isManager || isSelfApproval)
          ? ViewType.myApproval
          : ViewType.myLeaves
      );

      setCurrentView(ViewType.details);
      setIsDeeplinkLoading(false);
    })
    .catch(err => {
      console.error("Fetch error", err);
      setCurrentView(ViewType.home);
      setIsDeeplinkLoading(false);
    });

}, [context.sdks.microsoftTeams, isManager]);

  // 🔹 Fetch current logged-in user and assigned employees
  useEffect(() => {
    const loggedInUser = context.pageContext.user.displayName;
    setUserName(loggedInUser);

    const fetchAssignedEmployees = async () => {
      try {
        const url =
          `${hardcodedSiteUrl}/_api/web/lists/getbytitle('LeaveRequests')/items` +
          `?$select=EmployeeName,Manager/Title&$expand=Manager` +
          `&$filter=Manager/Title eq '${loggedInUser}'`;

        const response = await spHttpClient.get(url, SPHttpClient.configurations.v1);
        const data = await response.json();

        if (data.value && data.value.length > 0) {
          const uniqueEmployees: string[] = Array.from(
            new Set(data.value.map((item: any) => String(item.EmployeeName)))
          );
          setAssignedEmployees(uniqueEmployees);
          setIsManager(true);
        } else {
          setIsManager(false);
          setAssignedEmployees([]);
        }
      } catch (err) {
        console.error('Failed to fetch assigned employees', err);
      }
    };

    fetchAssignedEmployees();
  }, [context, spHttpClient]);

  // 🔹 Fetch Leave Data (Own or Selected Employee)
  useEffect(() => {
    const employeeToFetch =
      isManager && selectedEmployee ? selectedEmployee : userName;

    if (!employeeToFetch) return;

    const fetchLeaveData = async () => {
      setLoading(true);
      setError(null);
      try {
        const url =
          `${hardcodedSiteUrl}/_api/web/lists/getbytitle('LeaveInformations')/items` +
          `?$filter=EmployeeName eq '${employeeToFetch}'`;

        const response = await spHttpClient.get(url, SPHttpClient.configurations.v1);
        const data = await response.json();

        if (data.value && data.value.length > 0) {
          setLeaveData(data.value);
        } else {
          setLeaveData([]);
          setError('No leave data found.');
        }
      } catch (err) {
        console.error(err);
        setError('Failed to fetch leave data.');
      } finally {
        setLoading(false);
      }
    };

    fetchLeaveData();
  }, [userName, selectedEmployee, isManager, spHttpClient]);

  // 🔹 Navigation Handler
  const handleNavClick = (view: ViewType) => {
    setCurrentView(view);
    setSelectedItem(null);
    if (view === ViewType.apply && onViewChange) {
      onViewChange(view);
    }
  };

// 🔹 Block UI until deeplink decision is made
  if (isDeeplinkLoading || currentView === null) {
    return (
      <div className="container mt-4">
        Loading...
      </div>
    );
  }

  return (
    <div className="container mt-4" style={{ fontFamily: 'Segoe UI' }}>
      {/* existing JSX */}
      <div className="d-flex flex-wrap justify-content-between align-items-center mb-3">
        <h4 className="mb-2">Leave Management</h4>
        <button
          className="btn btn-success mb-2"
          onClick={() => handleNavClick(ViewType.apply)}
        >
          + Apply Leave
        </button>
      </div>

      {/* 🔹 Tabs */}
<ul className="nav nav-tabs flex-wrap mb-4">
  <li className="nav-item">
    <button
      className={`nav-link ${
        currentView === ViewType.home ? 'active' : ''
      }`}
      onClick={() => handleNavClick(ViewType.home)}
    >
      My Information
    </button>
  </li>

  <li className="nav-item">
    <button
      className={`nav-link ${
        currentView === ViewType.details
          ? sourceView === ViewType.myLeaves
            ? 'active'
            : ''
          : currentView === ViewType.myLeaves
          ? 'active'
          : ''
      }`}
      onClick={() => handleNavClick(ViewType.myLeaves)}
    >
      My Leaves
    </button>
  </li>

  <li className="nav-item">
    <button
      className={`nav-link ${
        currentView === ViewType.details
          ? sourceView === ViewType.myApproval
            ? 'active'
            : ''
          : currentView === ViewType.myApproval
          ? 'active'
          : ''
      }`}
      onClick={() => handleNavClick(ViewType.myApproval)}
    >
      My Approvals
    </button>
  </li>
</ul>

      {/* 🔹 My Information */}
      {currentView === ViewType.home && (
        <>
          <h5 className="mb-3">
            Welcome, {userName}{isManager ? ' (Manager)' : ''}
          </h5>

          {/* 🔹 Manager Employee Filter */}
          {isManager && assignedEmployees.length > 0 && (
            <div className="mb-3">
              <label className="form-label">Select Employee:</label>
              <select
                className="form-select w-auto"
                value={selectedEmployee}
                onChange={(e) => setSelectedEmployee(e.target.value)}
              >
                <option value="">-- My Own Balance --</option>
                {assignedEmployees.map((emp) => (
                  <option key={emp} value={emp}>
                    {emp}
                  </option>
                ))}
              </select>
            </div>
          )}

          {/* 🔹 Leave Data Table */}
          {loading ? (
            <p>Loading...</p>
          ) : error ? (
            <p className="text-danger">{error}</p>
          ) : (
            <div className="table-responsive">
              <table className="table table-striped table-hover align-middle">
                <thead className="table-light">
                  <tr>
                    <th>Leave Type</th>
                    <th>Total</th>
                    <th>Leave Taken</th>
                    <th>Remaining</th>
                  </tr>
                </thead>
                <tbody>
                  {leaveData.map((item) => (
                    <tr key={item.Id}>
                      <td>{item.LeaveType}</td>
                      <td>{item.TotalLeaves}</td>
                      <td>{item.LeaveTaken}</td>
                      <td>{item.RemainingLeaves}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </>
      )}

      {/* 🔹 My Leaves */}
      {currentView === ViewType.myLeaves && (
        <LeaveRequests
          context={context}
          spHttpClient={spHttpClient}
          siteUrl={hardcodedSiteUrl}
          viewType={ViewType.myLeaves}
          onSelectItem={(item) => setSelectedItem(item)}
          onViewChange={(view) => setCurrentView(view)}
        />
      )}

      {/* 🔹 My Approvals */}
      {currentView === ViewType.myApproval && (
        <LeaveRequests
          context={context}
          spHttpClient={spHttpClient}
          siteUrl={hardcodedSiteUrl}
          viewType={ViewType.myApproval}
          onSelectItem={(item) => setSelectedItem(item)}
          onViewChange={(view) => setCurrentView(view)}
        />
      )}

      {/* 🔹 Apply Leave */}
      {currentView === ViewType.apply && (
        <Leaves
          context={context}
          spHttpClient={spHttpClient}
          siteUrl={hardcodedSiteUrl}
        />
      )}

      {/* 🔹 Leave Details */}
      {currentView === ViewType.details && selectedItem && (
  <ApplyLeave
    context={context}
    spHttpClient={spHttpClient}
    siteUrl={hardcodedSiteUrl}
    item={selectedItem}
    viewType={currentView}
    sourceView={sourceView} 
    onBack={() => {
      setSelectedItem(null);
      setCurrentView(sourceView);
    }}
    onViewChange={(view) => setCurrentView(view)}
  />
)}
    </div>
  );
};

export default LeaveDashboard;