import * as React from 'react';
import { useEffect, useState } from 'react';
import { SPHttpClientResponse, SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import 'bootstrap/dist/css/bootstrap.min.css';

interface IParticipants {
  Id: number;
  Title: string;
  Email?: string;
}

interface ITrip {
  Id: number;
  TripName: string;
  StartDate: string;
  EndDate: string;
  Status: string;
  Participants?: IParticipants[];
}

interface TripDashboardProps {
  context: WebPartContext;
  onTripSelect: (trip: ITrip) => void;
  onAddTrip: () => void;
}

const TripDashboard: React.FC<TripDashboardProps> = ({ context, onTripSelect, onAddTrip }) => {
  const [trips, setTrips] = useState<ITrip[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);

  const [searchText, setSearchText] = useState<string>('');
  const [filter, setFilter] = useState<'all' | 'notstarted' | 'started' | 'closed'>('all');

  useEffect(() => {
    fetchTrips();
  }, [context]);

  const fetchTrips = async () => {
    setLoading(true);
    setError(null);

    const url = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('TripsData')/items?$select=Id,TripName,StartDate,EndDate,Status,Participants/Id,Participants/Title&$expand=Participants`;
    try {
      const response: SPHttpClientResponse = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
      }
      const data = await response.json();
      setTrips(data.value);
    } catch (err: any) {
      setError(err.message || 'Unknown error');
      setTrips([]);
    } finally {
      setLoading(false);
    }
  };

  // 🗑️ Delete Trip
  const deleteTrip = async (tripId: number) => {
    if (!confirm("Are you sure you want to delete this trip?")) return;

    const siteUrl = context.pageContext.web.absoluteUrl;
    try {
      await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('TripsData')/items(${tripId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE'
          }
        }
      );
      alert("Trip deleted successfully ✅");
      fetchTrips(); // refresh
    } catch (err: any) {
      alert("Error deleting trip: " + err.message);
    }
  };

  // 🔎 Search + Status Filter
  const filteredTrips = trips.filter((trip) => {
    const matchesText = trip.TripName.toLowerCase().includes(searchText.toLowerCase());

    let matchesFilter = true;
    if (filter === 'notstarted') {
      matchesFilter = trip.Status.toLowerCase() === 'not started';
    } else if (filter === 'started') {
      matchesFilter = trip.Status.toLowerCase() === 'started';
    } else if (filter === 'closed') {
      matchesFilter = trip.Status.toLowerCase() === 'trip closed';
    }

    return matchesText && matchesFilter;
  });

  return (
    <div className="container mt-4">
      <div className="d-flex justify-content-between align-items-center mb-4">
        <h2 className="fw-bold text-primary">🌍 Trips Dashboard</h2>
        <button className="btn btn-success shadow-sm px-4" onClick={onAddTrip}>
          + Add Trip
        </button>
      </div>

      {/* Search & Filter */}
      <div className="row mb-4">
        <div className="col-md-6">
          <input
            type="text"
            className="form-control shadow-sm"
            placeholder="🔎 Search by Trip Name..."
            value={searchText}
            onChange={(e) => setSearchText(e.target.value)}
          />
        </div>
        <div className="col-md-6">
          <select
            className="form-select shadow-sm"
            value={filter}
            onChange={(e) =>
              setFilter(e.target.value as 'all' | 'notstarted' | 'started' | 'closed')
            }
          >
            <option value="all">All Trips</option>
            <option value="notstarted">Not Started Trips</option>
            <option value="started">Started Trips</option>
            <option value="closed">Closed Trips</option>
          </select>
        </div>
      </div>

      {/* Loading, Error, No trips */}
      {loading && <div className="alert alert-info shadow-sm">⏳ Loading trips...</div>}
      {error && <div className="alert alert-danger shadow-sm">⚠️ Error: {error}</div>}
      {!loading && !error && filteredTrips.length === 0 && (
        <div className="alert alert-warning shadow-sm">🚫 No trips found.</div>
      )}

      {/* Trips list */}
      <div className="row">
        {filteredTrips.map((trip) => (
          <div className="col-md-4 mb-4" key={trip.Id}>
            <div className="card h-100 shadow-lg border-0 trip-card" style={{ borderRadius: "15px" }}>
              <div className="card-body">
                <h5 className="card-title fw-bold text-dark">{trip.TripName}</h5>
                <p className="card-text text-muted">
                  <i className="bi bi-calendar-event"></i>{" "}
                  {new Date(trip.StartDate).toLocaleDateString()} -{" "}
                  {new Date(trip.EndDate).toLocaleDateString()}
                  <br />
                  <span className="badge bg-info text-dark mt-2">{trip.Status}</span>
                  <br />
                  <strong>Participants:</strong>{" "}
                  <span className="text-dark">
                    {trip.Participants?.map((p) => p.Title).join(", ") || "None"}
                  </span>
                </p>
              </div>
              <div className="card-footer d-flex justify-content-between bg-white border-0">
                <button className="btn btn-primary btn-sm" onClick={() => onTripSelect(trip)}>
                  View
                </button>
                <button className="btn btn-danger btn-sm" onClick={() => deleteTrip(trip.Id)}>
                  Delete
                </button>
              </div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};

export default TripDashboard;