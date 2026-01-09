import * as React from 'react';
import { useState, useEffect } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { ViewType } from './ViewType';

export interface IApplyLeaveProps {
  context: WebPartContext;
  item: {
    Id: number;
    EmployeeName: string;
    LeaveType: string;
    StartDate: string;
    EndDate: string;
    Days: number;
    Status: string;
    Reason: string;
    Manager: { Email?: string; Title?: string };
  };
  onBack: () => void;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  onViewChange: (view: ViewType) => void;
  viewType: ViewType;
  showCancelButton?: boolean;
  showManagerActions?: boolean;
}

const hardcodedSiteUrl = "https://elevix.sharepoint.com/sites/Trainingportal";

const formatDate = (rawDate: string): string => {
  if (!rawDate) return '';
  const date = new Date(rawDate);
  return date.toLocaleDateString('en-GB', {
    day: '2-digit',
    month: 'short',
    year: 'numeric',
  });
};

const ApplyLeave: React.FC<IApplyLeaveProps> = ({
  item,
  onBack,
  spHttpClient,
  viewType,
  onViewChange
}) => {

  const [showDialog, setShowDialog] = useState(false);
  const [dialogMessage, setDialogMessage] = useState('');

  // 🔹 Comment & History
  const [comments, setComments] = useState('');
  const [commentsHistory, setCommentsHistory] = useState<string[]>([]);
  const [commentError, setCommentError] = useState('');

  // ✅ Load comments history for this leave from localStorage
  useEffect(() => {
    const stored = localStorage.getItem('leaveComments');
    if (stored) {
      const allComments = JSON.parse(stored);
      if (allComments[item.Id]) {
        setCommentsHistory(allComments[item.Id]);
      }
    }
  }, [item.Id]);

  // ✅ Save updated comments history to localStorage
  const saveCommentsToLocal = (updated: string[]) => {
    const stored = localStorage.getItem('leaveComments');
    const allComments = stored ? JSON.parse(stored) : {};
    allComments[item.Id] = updated;
    localStorage.setItem('leaveComments', JSON.stringify(allComments));
  };

  const handleApproval = async (actionStatus: string) => {
  // ✅ Only require comment for Approve/Reject
  if ((actionStatus === "Approved" || actionStatus === "Rejected") && comments.trim() === '') {
    setCommentError('Please enter a comment before submitting.');
    return;
  }
  setCommentError('');

  try {
    // ✅ Update LeaveRequests list
    const body = JSON.stringify({
      __metadata: { type: "SP.Data.LeaveRequestsListItem" },
      Status: actionStatus
    });

    await spHttpClient.post(
      `${hardcodedSiteUrl}/_api/web/lists/getbytitle('LeaveRequests')/items(${item.Id})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=verbose',
          'odata-version': '',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE',
        },
        body: body
      }
    );

    // ✅ Update LeaveInformations if Rejected or Cancelled
    if (actionStatus === "Rejected" || actionStatus === "Cancelled") {
      try {
        const infoResponse = await spHttpClient.get(
          `${hardcodedSiteUrl}/_api/web/lists/getbytitle('LeaveInformations')/items?$filter=EmployeeName eq '${item.EmployeeName}' and LeaveType eq '${item.LeaveType}'`,
          SPHttpClient.configurations.v1
        );
        const infoData = await infoResponse.json();
        const leaveInfo = infoData.value && infoData.value[0];

        if (leaveInfo) {
          const updatedRemaining = (leaveInfo.RemainingLeaves || 0) + item.Days;
          const leaveTaken = (leaveInfo.LeaveTaken || 0) - item.Days;
          const finalLeaveTaken = leaveTaken < 0 ? 0 : leaveTaken;

          await spHttpClient.post(
            `${hardcodedSiteUrl}/_api/web/lists/getbytitle('LeaveInformations')/items(${leaveInfo.Id})`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                'Accept': 'application/json;odata=nometadata',
                'Content-Type': 'application/json;odata=nometadata',
                'IF-MATCH': '*',
                'X-HTTP-Method': 'MERGE'
              },
              body: JSON.stringify({
                RemainingLeaves: updatedRemaining,
                LeaveTaken: finalLeaveTaken
              })
            }
          );
        }
      } catch (err) {
        console.error(`Error updating LeaveInformations for ${actionStatus}:`, err);
      }
    }

    // ✅ Add comment entry to local history (allow empty for Cancel)
    const entry = comments ? `[${new Date().toLocaleString()}] ${actionStatus}: ${comments}` 
                           : `[${new Date().toLocaleString()}] ${actionStatus}`;
    const updatedHistory = [entry, ...commentsHistory];
    setCommentsHistory(updatedHistory);
    saveCommentsToLocal(updatedHistory);
    setComments('');

    setDialogMessage(`Leave ${actionStatus} successfully.`);
    setShowDialog(true);

  } catch (error) {
    console.error(`${actionStatus} failed:`, error);
    setDialogMessage(`Failed to ${actionStatus.toLowerCase()} leave. Please try again.`);
    setShowDialog(true);
  }
};

  const handleDialogClose = () => {
    setShowDialog(false);
    if (dialogMessage.includes("successfully")) {
      onBack();
      
    }
  };

  const status = item.Status?.trim().toLowerCase();
  const isEmployee = viewType === ViewType.myLeaves;
  const isManager = viewType === ViewType.myApproval;

  const showCancel = isEmployee && status === "pending";
  const showApproveReject = isManager && status === "pending";

  return (
    <div className="container-fluid mt-3 px-2">
      <div className="card shadow-sm border-0 rounded-3">
        <div className="card-header bg-primary text-white text-center text-md-start">
          <h5 className="mb-0">Leave Request Details</h5>
        </div>

        <div className="card-body">
          {/* 🔹 Details Table */}
          <div className="table-responsive">
            <table className="table table-bordered align-middle mb-0" style={{ fontSize: '0.9rem' }}>
              <tbody>
                <tr><th>Employee Name</th><td>{item.EmployeeName}</td></tr>
                <tr><th>Leave Type</th><td>{item.LeaveType}</td></tr>
                <tr><th>Start Date</th><td>{formatDate(item.StartDate)}</td></tr>
                <tr><th>End Date</th><td>{formatDate(item.EndDate)}</td></tr>
                <tr><th>Days</th><td>{item.Days}</td></tr>
                <tr><th>Status</th><td>{item.Status}</td></tr>
                <tr><th>Reason</th><td className="text-wrap">{item.Reason}</td></tr>
                <tr><th>Manager</th><td>{item.Manager?.Title}</td></tr>
              </tbody>
            </table>
          </div>

          {/* 🔹 Manager Comment Input */}
          {showApproveReject && (
            <div className="mt-4">
              <label className="form-label fw-semibold">Add Comment <span style={{color:'red'}}>*</span></label>
              <textarea
                className="form-control"
                rows={2}
                placeholder="Type your comment here..."
                value={comments}
                onChange={(e) => setComments(e.target.value)}
              ></textarea>
              {commentError && (
                <small className="text-danger fw-semibold">{commentError}</small>
              )}
            </div>
          )}

          {/* 🔹 Action Buttons */}
          <div className="mt-4 d-flex flex-wrap justify-content-md-end">
            <button className="btn btn-secondary me-2 mb-2" onClick={() => onBack()}>
              Back
            </button>

            {showCancel && (
  <button
    className="btn btn-danger me-2 mb-2"
    onClick={(e) => {
      e.preventDefault();
      handleApproval("Cancelled");
    }}
  >
    Cancel
  </button>
)}

            {showApproveReject && (
              <>
                <button className="btn btn-success me-2 mb-2" onClick={() => handleApproval("Approved")}>
                  Approve
                </button>
                <button className="btn btn-danger mb-2" onClick={() => handleApproval("Rejected")}>
                  Reject
                </button>
              </>
            )}
          </div>

          {/* 🔹 Comments History Section */}
          <div className="mt-4">
            <h6 className="fw-semibold">Comments History:</h6>
            {commentsHistory.length > 0 ? (
              <div className="border rounded-3 p-2 bg-light" style={{ fontSize: '0.85rem' }}>
                {commentsHistory.map((cmt, index) => (
                  <div key={index} className="border-bottom py-1">{cmt}</div>
                ))}
              </div>
            ) : (
              <p className="text-muted" style={{ fontSize: '0.9rem' }}>
                No comments yet.
              </p>
            )}
          </div>
        </div>
      </div>

      {/* 🔹 Dialog */}
      {showDialog && (
        <div className="modal fade show" style={{ display: "block", backgroundColor: "rgba(0,0,0,0.4)" }}>
          <div className="modal-dialog modal-dialog-centered">
            <div className="modal-content">
              <div className="modal-header bg-primary text-white">
                <h5 className="modal-title">Message</h5>
                <button type="button" className="btn-close" onClick={handleDialogClose}></button>
              </div>
              <div className="modal-body text-center">
                <p>{dialogMessage}</p>
              </div>
              <div className="modal-footer justify-content-center">
                <button type="button" className="btn btn-primary" onClick={handleDialogClose}>OK</button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default ApplyLeave;