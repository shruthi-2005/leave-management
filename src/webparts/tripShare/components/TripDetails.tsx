import * as React from 'react';
import { useState } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import 'bootstrap/dist/css/bootstrap.min.css';

interface Participant {                        
  Title: string;
  Id: number;
  Email?: string;
}

interface Trip {
  Id: number;
  Title: string;
  TripName: string;
  StartDate: string;
  EndDate: string;
  Status: string;
  Participants: Participant[];
}

interface TripDetailsProps {
  context: WebPartContext;
  trip: Trip;
  onAddExpense: () => void;
  onViewSummary: () => void;
  onBack: () => void;
}

const TripDetails: React.FC<TripDetailsProps> = ({ context, trip, onAddExpense, onViewSummary, onBack }) => {

  const [editableTrip, setEditableTrip] = useState<Trip>({ ...trip });
  const [isEditing, setIsEditing] = useState(false);
  const [saving, setSaving] = useState(false);

  // ✅ Safe formatter for date fields
  const formatDate = (isoDate: string | null | undefined) => {
    if (!isoDate) return "";
    const date = new Date(isoDate);
    if (isNaN(date.getTime())) return "";
    return date.toISOString().split('T')[0];
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) => {
    const { name, value } = e.target;
    setEditableTrip(prev => ({ ...prev, [name]: value }));
  };

  const handleSave = async () => {
    setSaving(true);
    try {
      const response: SPHttpClientResponse = await context.spHttpClient.post(
        `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('TripsData')/items(${editableTrip.Id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            "Accept": "application/json;odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "odata-version": "",
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE"
          },
          body: JSON.stringify({
            TripName: editableTrip.TripName,
            StartDate: editableTrip.StartDate
              ? new Date(editableTrip.StartDate).toISOString().split('T')[0]
              : null,
            EndDate: editableTrip.EndDate
              ? new Date(editableTrip.EndDate).toISOString().split('T')[0]
              : null,
            Status: editableTrip.Status,
          })
        }
      );

      if (response.ok) {
        alert("Trip updated successfully!");
        setIsEditing(false);
      } else {
        const error = await response.json();
        console.error("Update failed:", error);
        alert("Error updating trip: " + (error.error?.message?.value || "Unknown error"));
      }
    } catch (err) {
      console.error("Error:", err);
      alert("Error updating trip.");
    }
    setSaving(false);
  };

  return (
    <div className="container mt-4">
      <h2 className="mb-4 text-center text-primary fw-bold border-bottom pb-2">
        ✈️ Trip Details
      </h2>

      <div className="card shadow-sm border-0 mb-4">
        <div className="card-body">
          {/* Trip Name */}
          <div className="row mb-2">
            <div className="col-md-6"><strong>Trip Name:</strong></div>
            <div className="col-md-6">
              {isEditing ? (
                <input
                  type="text"
                  className="form-control"
                  name="TripName"
                  value={editableTrip.TripName}
                  onChange={handleChange}
                />
              ) : (
                <span className="text-muted">{editableTrip.TripName}</span>
              )}
            </div>
          </div>

          {/* Start Date */}
          <div className="row mb-2">
            <div className="col-md-6"><strong>Start Date:</strong></div>
            <div className="col-md-6">
              {isEditing ? (
                <input
                  type="date"
                  className="form-control"
                  name="StartDate"
                  value={formatDate(editableTrip.StartDate)}
                  onChange={handleChange}
                />
              ) : (
                <span className="text-muted">
                  {editableTrip.StartDate
                    ? new Date(editableTrip.StartDate).toLocaleDateString()
                    : "—"}
                </span>
              )}
            </div>
          </div>

          {/* End Date */}
          <div className="row mb-2">
            <div className="col-md-6"><strong>End Date:</strong></div>
            <div className="col-md-6">
              {isEditing ? (
                <input
                  type="date"
                  className="form-control"
                  name="EndDate"
                  value={formatDate(editableTrip.EndDate)}
                  onChange={handleChange}
                />
              ) : (
                <span className="text-muted">
                  {editableTrip.EndDate
                    ? new Date(editableTrip.EndDate).toLocaleDateString()
                    : "—"}
                </span>
              )}
            </div>
          </div>

          {/* Status */}
          <div className="row mb-2">
            <div className="col-md-6"><strong>Status:</strong></div>
            <div className="col-md-6">
              {isEditing ? (
                <select
                  className="form-select"
                  name="Status"
                  value={editableTrip.Status}
                  onChange={handleChange}
                >
                  <option value="Not Started">Not Started</option>
                  <option value="Started">Started</option>
                  <option value="Trip Closed">Trip Closed</option>
                </select>
              ) : (
                <span
                  className={`badge ${
                    editableTrip.Status === "Completed"
                      ? "bg-success"
                      : editableTrip.Status === "Started"
                      ? "bg-warning text-dark"
                      : editableTrip.Status === "Trip Closed"
                      ? "bg-danger"
                      : "bg-secondary"
                  }`}
                >
                  {editableTrip.Status}
                </span>
              )}
            </div>
          </div>
        </div>
      </div>

      {/* Participants */}
      <h4 className="mb-3 text-secondary">👥 Participants</h4>
      <ul className="list-group mb-4 shadow-sm">
        {editableTrip.Participants && editableTrip.Participants.length > 0 ? (
          editableTrip.Participants.map(participant => (
            <li key={participant.Id} className="list-group-item d-flex justify-content-between align-items-center">
              <span>{participant.Title}</span>
              {participant.Email && <span className="text-muted small">{participant.Email}</span>}
            </li>
          ))
        ) : (
          <li className="list-group-item text-center text-muted">No participants added</li>
        )}
      </ul>

      {/* Action buttons */}
      <div className="d-flex gap-3 justify-content-center">
        {isEditing ? (
          <button className="btn btn-success px-4" onClick={handleSave} disabled={saving}>
            💾 {saving ? 'Saving...' : 'Save'}
          </button>
        ) : (
          <button className="btn btn-warning px-4" onClick={() => setIsEditing(true)}>
            ✏️ Edit
          </button>
        )}
        {editableTrip.Status !== 'Trip Closed' && !isEditing && (
          <button className="btn btn-primary px-4" onClick={onAddExpense}>
            ➕ Add Expense
          </button>
        )}
        <button className="btn btn-info text-white px-4" onClick={onViewSummary}>
          📊 View Summary
        </button>
        <button className="btn btn-outline-secondary px-4" onClick={onBack}>
          ← Back
        </button>
      </div>
    </div>
  );
};

export default TripDetails;