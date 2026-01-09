import * as React from 'react';
import { useState } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import 'bootstrap/dist/css/bootstrap.min.css';

export interface IAddTripProps {
  context: WebPartContext;
  siteUrl: string;
  onSave: (trip: any) => void;
  onCancel?: () => void;
  onBack?: () => void;
  onTripAdded?: (trip: { Id: number; TripName: string }) => void;
}

const AddTrip: React.FC<IAddTripProps> = ({ context, onCancel, onSave, onBack }) => {
  const [tripName, setTripName] = useState<string>('');
  const [startDate, setStartDate] = useState<string>('');
  const [endDate, setEndDate] = useState<string>('');
  const [status, setStatus] = useState<string>('Not Started');
  const [participants, setParticipants] = useState<string[]>([]);
  const [saving, setSaving] = useState(false);

  // 🔹 Error state only for EndDate
  const [endDateError, setEndDateError] = useState<string>('');

  const validateEndDate = (start: string, end: string) => {
    if (start && end && new Date(end) < new Date(start)) {
      setEndDateError("❌ End Date cannot be before Start Date.");
      return false;
    }
    setEndDateError('');
    return true;
  };

  const handleSave = async () => {
    if (!tripName || !startDate || !endDate || participants.length === 0) {
      alert("Please fill all required fields.");
      return;
    }

    if (!validateEndDate(startDate, endDate)) {
      return;
    }

    setSaving(true);

    try {
      const userIds: number[] = [];
      for (const email of participants) {
        const res = await context.spHttpClient.get(
          `${context.pageContext.web.absoluteUrl}/_api/web/siteusers/getbyemail('${email}')`,
          SPHttpClient.configurations.v1
        );
        const data = await res.json();
        if (data && data.Id) userIds.push(data.Id);
      }

      const item = {
        __metadata: { type: "SP.Data.TripsDataListItem" },
        TripName: tripName,
        StartDate: startDate ? new Date(startDate).toISOString().split('T')[0] : null,
        EndDate: endDate ? new Date(endDate).toISOString().split('T')[0] : null,
        Status: status,
        ParticipantsId: { results: userIds }
      };

      const saveRes = await context.spHttpClient.post(
        `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('TripsData')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "odata-version": ""
          },
          body: JSON.stringify(item)
        }
      );

      if (saveRes.ok) {
        const data = await saveRes.json();
        alert("Trip saved successfully");
        onSave?.({ Id: data.d.Id, TripName: tripName });
        resetForm();
        if (onBack) onBack();
      } else {
        const err = await saveRes.json();
        alert("Error: " + (err.error?.message?.value || "Unknown error"));
      }
    } catch (error) {
      console.error("Network error:", error);
      alert("Network error: " + error);
    }

    setSaving(false);
  };

  const resetForm = () => {
    setTripName('');
    setStartDate('');
    setEndDate('');
    setStatus('Not Started');
    setParticipants([]);
    setEndDateError('');
  };

  const peoplePickerContext: IPeoplePickerContext = {
    spHttpClient: context.spHttpClient,
    msGraphClientFactory: context.msGraphClientFactory,
    absoluteUrl: context.pageContext.web.absoluteUrl,
  };

  return (
    <div className="container mt-5">
      <div className="card shadow-lg border-0 rounded-3">
        <div className="card-header bg-primary text-white text-center fw-bold fs-5">
          Add New Trip
        </div>
        <div className="card-body p-4">
          <table className="table table-borderless align-middle">
            <tbody>
              <tr>
                <th style={{ width: "30%" }}>Trip Name <span className="text-danger">*</span></th>
                <td>
                  <input
                    type="text"
                    className="form-control"
                    value={tripName}
                    placeholder="Enter trip name"
                    onChange={e => setTripName(e.target.value)}
                  />
                </td>
              </tr>
              <tr>
                <th>Start Date <span className="text-danger">*</span></th>
                <td>
                  <input
                    type="date"
                    className="form-control"
                    value={startDate}
                    onChange={e => {
                      setStartDate(e.target.value);
                      validateEndDate(e.target.value, endDate);
                    }}
                  />
                </td>
              </tr>
              <tr>
                <th>End Date <span className="text-danger">*</span></th>
                <td>
                  <input
                    type="date"
                    className={`form-control ${endDateError ? 'is-invalid' : ''}`}
                    value={endDate}
                    onChange={e => {
                      setEndDate(e.target.value);
                      validateEndDate(startDate, e.target.value);
                    }}
                  />
                  {endDateError && (
                    <div className="alert alert-danger mt-2 p-2">{endDateError}</div>
                  )}
                </td>
              </tr>
              <tr>
                <th>Status</th>
                <td>
                  <select
                    className="form-select"
                    value={status}
                    onChange={e => setStatus(e.target.value)}
                  >
                    <option value="Not Started">Not Started</option>
                    <option value="Started">Started</option>
                    <option value="Trip Closed">Trip Closed</option>
                  </select>
                </td>
              </tr>
              <tr>
                <th>Participants <span className="text-danger">*</span></th>
                <td>
                  <PeoplePicker
                    context={peoplePickerContext}
                    titleText="Select Participants"
                    personSelectionLimit={5}
                    showtooltip={true}
                    required={false}
                    onChange={(items: any[]) => {
                      const emails = items.map(u => u.secondaryText);
                      setParticipants(emails);
                    }}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={500}
                  />
                </td>
              </tr>
            </tbody>
          </table>

          <div className="text-center mt-3">
            <button className="btn btn-success px-4 me-2" onClick={handleSave} disabled={saving}>
              💾 {saving ? "Saving..." : "Save Trip"}
            </button>
            <button className="btn btn-secondary px-4 me-2" onClick={onCancel}>
              Cancel
            </button>
            <button className="btn btn-secondary px-4" onClick={resetForm}>
              Reset
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

export default AddTrip;