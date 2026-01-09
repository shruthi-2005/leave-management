import * as React from 'react';
import { useState, useEffect } from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import 'bootstrap/dist/css/bootstrap.min.css';
import { ILeaveRequestItem } from './LeaveRequests';


export interface ILeaveInfo {
  Id: number;
  EmployeeName: string;  
  LeaveType: string;
  RemainingLeaves: number;
}

export interface ILeavesProps {
  context: WebPartContext;
  siteUrl: string;
  onViewChange?: (view: any) => void;
  item?:ILeaveRequestItem; 
  spHttpClient:SPHttpClient;
  formatData?:(date:string)=>void;
}

const Leaves: React.FC<ILeavesProps> = ({ context, siteUrl, onViewChange }) => {

  // ✅ Hardcoded site URL for Teams environment
  siteUrl = "https://elevix.sharepoint.com/sites/Trainingportal";

  const [employeeName, setEmployeeName] = useState<string>('');
  const [Employeename0Id, setEmployeename0Id] = useState<string>('');
  const [leaveType, setLeaveType] = useState<string>('');
  const [startDate, setStartDate] = useState<string>('');
  const [endDate, setEndDate] = useState<string>('');
  const [days, setDays] = useState<number>(0);
  const [reason, setReason] = useState<string>('');
  const [managerId, setManagerId] = useState<number | null>(null);
  const [managerEmail, setManagerEmail] = useState<string>('');

  const [startHalfDay, setStartHalfDay] = useState<string>('');
  const [endHalfDay, setEndHalfDay] = useState<string>('');

  const [leaveBalances, setLeaveBalances] = useState<{ [key: string]: number }>({});
  const [remainingDays, setRemainingDays] = useState<number>(0);
  const [holidays, setHolidays] = useState<string[]>([]);

  const [warning, setWarning] = useState<string>('');
  const [submitDisabled, setSubmitDisabled] = useState<boolean>(true);
  const [existingLeaves, setExistingLeaves] = useState<{ StartDate: string; EndDate: string }[]>([]);
 

  // 🔹 Modal state
  const [showModal, setShowModal] = useState<boolean>(false);
  const [modalMessage, setModalMessage] = useState<string>('');

  useEffect(() => {
    context.spHttpClient.get(`${siteUrl}/_api/web/currentUser`, SPHttpClient.configurations.v1)
      .then(res => res.json())
      .then(user => {
        setEmployeeName(user.Title);
        setEmployeename0Id(user.Id);
        fetchLeaveBalances(user.Title);
        fetchHolidays();
      })
      .catch(() => setEmployeeName(''));
  }, []);

  const fetchLeaveBalances = (name: string) => {
    context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getByTitle('LeaveInformations')/items?$filter=EmployeeName eq '${name}'&$select=LeaveType,RemainingLeaves`,
      SPHttpClient.configurations.v1
    )
      .then(res => res.json())
      .then(data => {
        const balances: { [key: string]: number } = {};
        data.value.forEach((item: ILeaveInfo) => {
          balances[item.LeaveType] = item.RemainingLeaves;
        });
        setLeaveBalances(balances);
      })
      .catch(() => setLeaveBalances({}));
  };


  useEffect(() => {
  context.spHttpClient.get(`${siteUrl}/_api/web/currentUser`, SPHttpClient.configurations.v1)
    .then(res => res.json())
    .then(user => {
      setEmployeeName(user.Title);
      setEmployeename0Id(user.Id);
      fetchLeaveBalances(user.Title);
      fetchHolidays();
      fetchExistingLeaves(user.Title);  // ✅ add this line
    })
    .catch(() => setEmployeeName(''));
}, []);

  const fetchHolidays = () => {
    context.spHttpClient.get(
      `${siteUrl}/_api/web/lists/getByTitle('Holidays')/items?$select=Date`,
      SPHttpClient.configurations.v1
    )
      .then(res => res.json())
      .then(data => {
        const holidayDates = data.value.map((item: any) => {
          return new Date(item.Date).toISOString().split('T')[0];
        });
        setHolidays(holidayDates);
      })
      .catch(() => setHolidays([]));
  };

  const isWeekend = (date: Date): boolean => {
    const day = date.getDay();
    return day === 0 || day === 1;
  };

  const isHoliday = (date: Date): boolean => {
    const dateStr = date.toISOString().split('T')[0];
    return holidays.includes(dateStr);
  };
  const fetchExistingLeaves = (name: string) => {
  context.spHttpClient.get(
    `${siteUrl}/_api/web/lists/getByTitle('LeaveRequests')/items?$filter=EmployeeName eq '${name}' and Status ne 'Rejected'&$select=StartDate,EndDate`,
    SPHttpClient.configurations.v1
  )
    .then(res => res.json())
    .then(data => {
      const leaves = data.value.map((item: any) => ({
        StartDate: new Date(item.StartDate).toISOString().split('T')[0],
        EndDate: new Date(item.EndDate).toISOString().split('T')[0],
      }));
      setExistingLeaves(leaves);
    })
    .catch(() => setExistingLeaves([]));
};


const hasDateOverlap = (newStart: string, newEnd: string): boolean => {
  if (!newStart || !newEnd) return false;

  const start = new Date(newStart);
  const end = new Date(newEnd);

  return existingLeaves.some((leave) => {
    const oldStart = new Date(leave.StartDate);
    const oldEnd = new Date(leave.EndDate);

    // Overlap condition
    return (
      (start <= oldEnd && end >= oldStart)
    );
  });
};


  const calculateDays = (start: string, end: string, startHalf: string, endHalf: string) => {
    setWarning('');
    setDays(0);

    if (!start || !end) return;

    const sDate = new Date(start);
    const eDate = new Date(end);

    if (eDate < sDate) {
      setWarning('End date cannot be before start date.');
      return;
    }
    if (hasDateOverlap(start, end)) {
  setWarning('You already have a leave request in this date range.');
  setSubmitDisabled(true);
  return;
}

    let count = 0;
    let currentDate = new Date(sDate);

    while (currentDate <= eDate) {
      if (!isWeekend(currentDate) && !isHoliday(currentDate)) {
        count++;
      }
      currentDate.setDate(currentDate.getDate() + 1);
    }

    if (start === end) {
  if (startHalf === 'Morning' || startHalf === 'Afternoon') {
    count = 0.5;
  }
} else {
  if (startHalf === 'Morning' || startHalf === 'Afternoon') count -= 0.5;
  if (endHalf === 'Morning' || endHalf === 'Afternoon') count -= 0.5;
}

    setDays(count);

    // ✅ Leave balance validation logic
    if (leaveType && leaveBalances[leaveType] !== undefined) {
      const balance = leaveBalances[leaveType];
      setRemainingDays(balance);

      if (count > balance) {
        setWarning(`You only have ${balance} days remaining for ${leaveType}, but you selected ${count} days.`);
      } else if (balance - count === 0) {
        setWarning(`After this leave, your remaining balance for ${leaveType} will be 0 days.`);
      } else {
        setWarning('');
      }
    }
  };

  const validateForm = () => {
    const isValid =
      leaveType !== '' &&
      startDate !== '' &&
      endDate !== '' &&
      reason.trim() !== '' &&
      managerId !== null &&
      days > 0 &&
      warning === '';

    setSubmitDisabled(!isValid);
  };

  useEffect(() => {
    calculateDays(startDate, endDate, startHalfDay, endHalfDay);
    validateForm();
  }, [leaveType, startDate, endDate, startHalfDay, endHalfDay, reason, managerId, warning]);

  const getUserIdFromLoginName = async (loginName: string): Promise<number | null> => {
    try {
      const res = await context.spHttpClient.get(
        `${siteUrl}/_api/web/siteusers(@v)?@v='${encodeURIComponent(loginName)}'`,
        SPHttpClient.configurations.v1
      );
      const data = await res.json();
      return data.Id || null;
    } catch {
      return null;
    }
  };

  const getUserIdFromEmail = async (email: string): Promise<number | null> => {
    try {
      const res = await context.spHttpClient.get(
        `${siteUrl}/_api/web/siteusers/getByEmail('${email}')`,
        SPHttpClient.configurations.v1
      );
      if (res.ok) {
        const data = await res.json();
        return data.Id || null;
      }
      return null;
    } catch {
      return null;
    }
  };

  const getUserEmailById = async (id: number): Promise<string> => {
    try {
      const res = await context.spHttpClient.get(
        `${siteUrl}/_api/web/siteusers(${id})`,
        SPHttpClient.configurations.v1
      );
      if (res.ok) {
        const data = await res.json();
        return data.Email || '';
      }
      return '';
    } catch {
      return '';
    }
  };

  const sendEmailToManager = async (toEmail: string, subject: string, body: string) => {
    if (!toEmail) return;
    const emailProperties = {
      properties: {
        __metadata: { type: 'SP.Utilities.EmailProperties' },
        To: { results: [toEmail] },
        Subject: subject,
        Body: body
      }
    };

    try {
      await context.spHttpClient.post(
        `${siteUrl}/_api/SP.Utilities.Utility.SendEmail`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=verbose'
          },
          body: JSON.stringify(emailProperties)
        }
      );
    } catch (err) {
      console.error('Error sending email:', err);
    }
  };

  const handleSubmit = async () => {
    if (submitDisabled) return;

    const isoStartDate = new Date(startDate).toISOString();
    const isoEndDate = new Date(endDate).toISOString();

    const requestData = {
      __metadata: { type: "SP.Data.LeaveRequestsListItem" },
      Title: `Leave - ${employeeName} (${startDate} to ${endDate})`,
      EmployeeName: employeeName,
      Employeename0Id:Number(Employeename0Id),
      LeaveType: leaveType,
      StartDate: isoStartDate,
      EndDate: isoEndDate,
      Days: Number(days),
      Reason: reason,
      Status: 'Pending',
      ManagerId: Number(managerId),
      HalfDayLeave: startHalfDay !== '' || endHalfDay !== '',
      HalfDayLeaveStart: startHalfDay || '',
      HalfDayLeaveEnd: endHalfDay || ''
    };

    try {
      const res = await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getByTitle('LeaveRequests')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'odata-version': ''
          },
          body: JSON.stringify(requestData)
        }
      );

      if (res.ok) {
        let mgrEmail = managerEmail;
        if (!mgrEmail && managerId) {
          mgrEmail = await getUserEmailById(managerId);
          setManagerEmail(mgrEmail);
        }

        try {
          const infoResponse = await context.spHttpClient.get(
            `${siteUrl}/_api/web/lists/getbytitle('LeaveInformations')/items?$filter=EmployeeName eq '${employeeName}' and LeaveType eq '${leaveType}'`,
            SPHttpClient.configurations.v1
          );
          const infoData = await infoResponse.json();
          const leaveInfo = infoData.value && infoData.value[0];

          if (leaveInfo) {
            const updatedRemaining = (leaveInfo.RemainingLeaves || 0) - Number(days);
            const leaveTaken = (leaveInfo.TotalLeaves || 0) - updatedRemaining;

            await context.spHttpClient.post(
              `${siteUrl}/_api/web/lists/getbytitle('LeaveInformations')/items(${leaveInfo.Id})`,
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
                  LeaveTaken: leaveTaken
                })
              }
            );
          } else {
            const leaveTypes = ['Casual Leave', 'Sick Leave', 'Annual Leave'];
            for (const type of leaveTypes) {
              const createData = {
                __metadata: { type: 'SP.Data.LeaveInformationsListItem' },
                EmployeeName: employeeName,
                LeaveType: type,
                TotalLeaves: 10,
                LeaveTaken: type === leaveType ? Number(days) : 0,
                RemainingLeaves: type === leaveType ? 10 - Number(days) : 10
              };

              await context.spHttpClient.post(
                `${siteUrl}/_api/web/lists/getbytitle('LeaveInformations')/items`,
                SPHttpClient.configurations.v1,
                {
                  headers: {
                    'Accept': 'application/json;odata=verbose',
                    'Content-Type': 'application/json;odata=verbose'
                  },
                  body: JSON.stringify(createData)
                }
              );
            }
          }
        } catch (err) {
          console.error('Error updating LeaveInformations:', err);
        }

        try {
          const subject = `Leave Request: ${employeeName} (${startDate} to ${endDate})`;
          const body = `
            <p>Hello,</p>
            <p>${employeeName} has submitted a leave request.</p>
            <ul>
              <li><strong>Leave Type:</strong> ${leaveType}</li>
              <li><strong>Dates:</strong> ${startDate} to ${endDate}</li>
              <li><strong>Days:</strong> ${days}</li>
              <li><strong>Reason:</strong> ${reason}</li>
            </ul>
            <p>Please review the request in the Leave Requests list.</p>
          `;
          if (mgrEmail) {
            await sendEmailToManager(mgrEmail, subject, body);
          }
        } catch (err) {
          console.error('Error while attempting to send manager email:', err);
        }

        setModalMessage('Leave request submitted successfully.');
        setShowModal(true);
        resetForm();
      } else {
        const err = await res.json();
        setModalMessage('Error submitting leave request: ' + (err.error?.message?.value || 'Unknown error'));
        setShowModal(true);
      }
    } catch (e) {
      console.error(e);
      setModalMessage('Network error submitting leave request.');
      setShowModal(true);
    }
  };

  const resetForm = () => {
    setLeaveType('');
    setStartDate('');
    setEndDate('');
    setDays(0);
    setReason('');
    setManagerId(null);
    setManagerEmail('');
    setStartHalfDay('');
    setEndHalfDay('');
    setWarning('');
    setSubmitDisabled(true);
  };

  const PeoplePickerContext: IPeoplePickerContext = {
    spHttpClient: context.spHttpClient,
    msGraphClientFactory: context.msGraphClientFactory,
    absoluteUrl: context.pageContext.web.absoluteUrl
  };

  const handleCloseModal = () => {
    setShowModal(false);
    if (typeof onViewChange === 'function') {
        onViewChange('home');
    }
  };
  ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 return (
  <>
    <div className="container-fluid mt-3" style={{ fontFamily: "Segoe UI" }}>
      
      {/* ✅ Heading */}
      <h4 className="mb-3 text-center text-md-start">Apply for Leave</h4>

      <div
        className="card shadow-sm p-3"
        style={{ backgroundColor: "#f0f8ff", borderRadius: "12px" }}
      >
        <div className="row g-3">

          {/* 🔹 Employee Name */}
          <div className="col-12">
            <label className="form-label">Employee Name</label>
            <input
              type="text"
              className="form-control"
              value={employeeName}
              readOnly
            />
          </div>

          {/* 🔹 Leave Type */}
          <div className="col-12">
            <label className="form-label">Leave Type</label>
            <select
              className="form-select"
              value={leaveType}
              onChange={e => setLeaveType(e.target.value)}
            >
              <option value="">-- Select Leave Type --</option>
              <option value="casual Leave">Casual Leave</option>
              <option value="sick Leave">Sick Leave</option>
              <option value="annual Leave">Annual Leave</option>
            </select>

            {remainingDays > 0 && leaveType && (
              <small className="text-muted d-block mt-1">
                Remaining balance: {remainingDays} days
              </small>
            )}
          </div>

          {/* 🔹 Start Date */}
          <div className="col-12">
            <label className="form-label">Start Date</label>
            <input
              type="date"
              className="form-control"
              value={startDate}
              onChange={e => {
                setStartDate(e.target.value);
                setStartHalfDay('');
                setEndHalfDay('');
              }}
            />
          </div>

          {/* 🔹 End Date */}
          <div className="col-12">
            <label className="form-label">End Date</label>
            <input
              type="date"
              className="form-control"
              value={endDate}
              onChange={e => {
                setEndDate(e.target.value);
                setStartHalfDay('');
                setEndHalfDay('');
                calculateDays(startDate, e.target.value, startHalfDay, endHalfDay);
              }}
            />
          </div>

          {/* 🔹 Warning */}
          {warning && (
            <div className="col-12">
              <div className="alert alert-warning text-center mb-0">
                {warning}
              </div>
            </div>
          )}

          {/* 🔹 Half Day (Same Day) */}
          {startDate && endDate && startDate === endDate && (
            <div className="col-12">
              <label className="form-label">Half Day Leave</label>
              <div className="d-flex gap-3">
                <label>
                  <input
                    type="radio"
                    name="halfDaySame"
                    value="Morning"
                    checked={startHalfDay === 'Morning'}
                    onClick={() => {
                      const val = startHalfDay === 'Morning' ? '' : 'Morning'; // toggle deselect
                      setStartHalfDay(val);
                      calculateDays(startDate, endDate, val, '');
                    }}
                  /> Morning
                </label>
                <label>
                  <input
                    type="radio"
                    name="halfDaySame"
                    value="Afternoon"
                    checked={startHalfDay === 'Afternoon'}
                    onClick={() => {
                      const val = startHalfDay === 'Afternoon' ? '' : 'Afternoon'; // toggle deselect
                      setStartHalfDay(val);
                      calculateDays(startDate, endDate, val, '');
                    }}
                  /> Afternoon
                </label>
              </div>
            </div>
          )}

          {/* 🔹 Half Day (Different Dates) */}
          {startDate && endDate && startDate !== endDate && (
            <>
              <div className="col-12 col-md-6">
                <label className="form-label">Start Date Half Day</label>
                <div className="d-flex gap-3">
                  <label>
                    <input
                      type="radio"
                      name="startHalfDay"
                      value="Morning"
                      checked={startHalfDay === 'Morning'}
                      onClick={() => {
                        const val = startHalfDay === 'Morning' ? '' : 'Morning';
                        setStartHalfDay(val);
                        calculateDays(startDate, endDate, val, endHalfDay);
                      }}
                    /> Morning
                  </label>
                  <label>
                    <input
                      type="radio"
                      name="startHalfDay"
                      value="Afternoon"
                      checked={startHalfDay === 'Afternoon'}
                      onClick={() => {
                        const val = startHalfDay === 'Afternoon' ? '' : 'Afternoon';
                        setStartHalfDay(val);
                        calculateDays(startDate, endDate, val, endHalfDay);
                      }}
                    /> Afternoon
                  </label>
                </div>
              </div>

              <div className="col-12 col-md-6">
                <label className="form-label">End Date Half Day</label>
                <div className="d-flex gap-3">
                  <label>
                    <input
                      type="radio"
                      name="endHalfDay"
                      value="Morning"
                      checked={endHalfDay === 'Morning'}
                      onClick={() => {
                        const val = endHalfDay === 'Morning' ? '' : 'Morning';
                        setEndHalfDay(val);
                        calculateDays(startDate, endDate, startHalfDay, val);
                      }}
                    /> Morning
                  </label>
                  <label>
                    <input
                      type="radio"
                      name="endHalfDay"
                      value="Afternoon"
                      checked={endHalfDay === 'Afternoon'}
                      onClick={() => {
                        const val = endHalfDay === 'Afternoon' ? '' : 'Afternoon';
                        setEndHalfDay(val);
                        calculateDays(startDate, endDate, startHalfDay, val);
                      }}
                    /> Afternoon
                  </label>
                </div>
              </div>
            </>
          )}

          {/* 🔹 Total Days */}
          <div className="col-12">
            <label className="form-label">Total Days</label>
            <input
              type="number"
              className="form-control"
              value={days}
              readOnly
            />
          </div>

          {/* 🔹 Reason */}
          <div className="col-12">
            <label className="form-label">Reason</label>
            <textarea
              className="form-control"
              rows={3}
              value={reason}
              onChange={e => setReason(e.target.value)}
            />
          </div>

          {/* 🔹 Manager */}
          <div className="col-12">
            <label className="form-label">Manager</label>
            <div className="form-control p-0">
              <PeoplePicker
                context={PeoplePickerContext}
                titleText=""
                personSelectionLimit={1}
                showtooltip={true}
                required={true}
                ensureUser={true}
                
                onChange={async (items: any[]) => {
                  if (items.length > 0) {
                    const user = items[0];
                    if (user.secondaryText) {
                      const userId = await getUserIdFromEmail(user.secondaryText);
                      setManagerId(userId);
                      setManagerEmail(user.secondaryText);
                    } else if (user.loginName) {
                      const userId = await getUserIdFromLoginName(user.loginName);
                      setManagerId(userId);
                      if (userId) {
                        const email = await getUserEmailById(userId);
                        setManagerEmail(email);
                      } else {
                        setManagerEmail('');
                      }
                    }
                  } else {
                    setManagerId(null);
                    setManagerEmail('');
                  }
                }}
                principalTypes={[PrincipalType.User]}
                resolveDelay={200}
              />
            </div>
          </div>

        </div>

        {/* 🔘 Buttons */}
        <div className="d-flex justify-content-end mt-3 gap-2">
          <button
            className="btn btn-secondary"
            onClick={() => {
              if (typeof onViewChange === "function") onViewChange("home");
            }}
          >
            Back
          </button>
          <button
            className="btn btn-primary"
            onClick={handleSubmit}
            disabled={submitDisabled}
          >
            Submit
          </button>
        </div>
      </div>

      {/* ✅ Modal */}
      {showModal && (
        <div className="modal fade show" style={{ display: 'block', backgroundColor: 'rgba(0,0,0,0.4)' }}>
          <div className="modal-dialog modal-dialog-centered">
            <div className="modal-content">
              <div className="modal-header bg-primary text-white">
                <h5 className="modal-title">Notification</h5>
                <button className="btn-close" onClick={handleCloseModal}></button>
              </div>
              <div className="modal-body text-center">
                <p>{modalMessage}</p>
              </div>
              <div className="modal-footer justify-content-center">
                <button className="btn btn-primary" onClick={handleCloseModal}>OK</button>
              </div>
            </div>
          </div>
        </div>
      )}

    </div>
  </>
);
}
export default Leaves;  
