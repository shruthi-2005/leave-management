import * as React from 'react';
import { useState, useEffect } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import 'bootstrap/dist/css/bootstrap.min.css';

interface Participant {
  Id: number;
  Title: string;
  Share: number;
  TotalPaid: number;
  NetBalance: number;
}

interface Expense {
  Id: number;
  Title: string;
  Amount: number;
  PaidById: number;
  PaidByName: string;
}

interface Trip {
  Id: number;
  Title: string;
  Status: string;
}

interface SummaryProps {
  context: WebPartContext;
  trip: Trip;
  onBack: () => void;
  readOnly?: boolean;
   onCloseTrip?: () => void;
}

interface ISettlement {
  from: string;
  to: string;
  fromUserId?: number;
  toUserId?: number;
  amount: number;
  status: string;
}

const Summary: React.FC<SummaryProps> = ({ context, trip, onBack, readOnly = false }) => {
  const [participants, setParticipants] = useState<Participant[]>([]);
  const [expenses, setExpenses] = useState<Expense[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [selectedParticipant, setSelectedParticipant] = useState<Participant | null>(null);
  const [settlements, setSettlements] = useState<ISettlement[]>([]);
  const [editingRow, setEditingRow] = useState<number | null>(null);
  const [settlementsCalculated, setSettlementsCalculated] = useState(false);

  useEffect(() => {
    fetchData();
  }, [trip]);

  const fetchData = async () => {
    setLoading(true);
    setError('');
    try {
      const siteUrl = context.pageContext.web.absoluteUrl;

      // Fetch Participants
      const participantsResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('TripsData')/items(${trip.Id})?$select=Participants/Id,Participants/Title&$expand=Participants`,
        SPHttpClient.configurations.v1
      );
      if (!participantsResponse.ok) throw new Error('Failed to fetch participants.');
      const participantsData = await participantsResponse.json();
      const baseParticipants: Participant[] = (participantsData.Participants || []).map((p: any) => ({
        Id: p.Id,
        Title: p.Title,
        Share: 0,
        TotalPaid: 0,
        NetBalance: 0
      }));

      // Fetch Expenses
      const expensesResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('TripExpenses')/items?$select=Id,Title,Amount,SpentBy/Id,SpentBy/Title,TripId&$expand=SpentBy&$filter=TripId eq ${trip.Id}`,
        SPHttpClient.configurations.v1
      );
      if (!expensesResponse.ok) throw new Error('Failed to fetch expenses.');
      const expensesData = await expensesResponse.json();
      const expensesList: Expense[] = (expensesData.value || []).map((e: any) => ({
        Id: e.Id,
        Title: e.Title,
        Amount: Number(e.Amount),
        PaidById: e.SpentBy?.Id || 0,
        PaidByName: e.SpentBy?.Title || ''
      }));

      // Fetch Settlements
      const settlementsResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('TripSettlements')/items?$filter=TripId eq ${trip.Id}`,
        SPHttpClient.configurations.v1
      );
      const settlementsData = await settlementsResponse.json();
      const settledList: ISettlement[] = (settlementsData.value || []).map((s: any) => ({
        from: baseParticipants.find(p => p.Id === s.FromUserId)?.Title || '',
        to: baseParticipants.find(p => p.Id === s.ToUserId)?.Title || '',
        fromUserId: s.FromUserId,
        toUserId: s.ToUserId,
        amount: Number(s.SettlementAmount),
        status: s.Status
      }));

      // Calculate per-head
      const totalExpense = expensesList.reduce((sum, e) => sum + (Number.isFinite(e.Amount) ? e.Amount : 0), 0);
      const count = baseParticipants.length;
      const perHead = count > 0 ? totalExpense / count : 0;

      const updatedParticipants = baseParticipants.map(p => {
        const totalPaid = expensesList.filter(exp => exp.PaidById === p.Id).reduce((sum, exp) => sum + exp.Amount, 0);
        const totalSettledPaid = settledList.filter(st => st.fromUserId === p.Id && st.status === 'Settled').reduce((sum, st) => sum + st.amount, 0);
        const totalSettledReceived = settledList.filter(st => st.toUserId === p.Id && st.status === 'Settled').reduce((sum, st) => sum + st.amount, 0);
        const net = (totalPaid - perHead) - totalSettledPaid + totalSettledReceived;
        return { ...p, Share: perHead, TotalPaid: totalPaid, NetBalance: net };
      });

      setParticipants(updatedParticipants);
      setExpenses(expensesList);
      setSettlements(settledList);

      // Auto calculate settlements if trip closed
      if (trip.Status === 'Trip Closed' && !settlementsCalculated) {
        calculateSettlements(updatedParticipants);
        setSettlementsCalculated(true);
      }
    } catch (e: any) {
      setError(e.message);
    }
    setLoading(false);
  };

  const participantClick = (p: Participant) => {
    if (readOnly || trip.Status === 'Trip Closed') return;
    setSelectedParticipant(p);
  };

  const closeDetail = () => setSelectedParticipant(null);

  const closeTrip = async () => {
    if (!window.confirm('Are you sure you want to close this trip?')) return;
    try {
      const siteUrl = context.pageContext.web.absoluteUrl;
      const body = { Status: 'Trip Closed' };

      const response = await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('TripsData')/items(${trip.Id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            "odata-version":"",
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          },
          body: JSON.stringify(body)
        }
      );

      if (response.ok) {
        alert('Trip closed successfully!');
        await fetchData();
      } else {
        const err = await response.json();
        alert('Error closing trip: ' + (err.error?.message || response.statusText));
      }
    } catch (e: any) {
      alert('Error closing trip: ' + e.message);
    }
  };

  const calculateSettlements = (currentParticipants: Participant[] = participants) => {
    if (currentParticipants.length === 0) return;
    const lenders: any[] = [];
    const borrowers: any[] = [];

    currentParticipants.forEach(p => {
      const diff = p.TotalPaid - p.Share;
      if (diff > 0) lenders.push({ name: p.Title, id: p.Id, amount: diff });
      else if (diff < 0) borrowers.push({ name: p.Title, id: p.Id, amount: -diff });
    });

    const results: ISettlement[] = [];
    let i = 0, j = 0;
    while (i < lenders.length && j < borrowers.length) {
      const lend = lenders[i];
      const borrow = borrowers[j];
      const settleAmount = Math.min(lend.amount, borrow.amount);

      results.push({
        from: borrow.name,
        to: lend.name,
        fromUserId: borrow.id,
        toUserId: lend.id,
        amount: settleAmount,
        status: "Pending"
      });

      lend.amount -= settleAmount;
      borrow.amount -= settleAmount;

      if (lend.amount === 0) i++;
      if (borrow.amount === 0) j++;
    }

    setSettlements(results);
    setSettlementsCalculated(true);
  };

  const saveSettlement = async (settlement: ISettlement, index: number) => {
    try {
      const siteUrl = context.pageContext.web.absoluteUrl;
      const body = {
        FromUserId: settlement.fromUserId,
        ToUserId: settlement.toUserId,
        SettlementAmount: settlement.amount,
        Status: "Settled",
        TripId: trip.Id.toString()
      };

      const response = await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('TripSettlements')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'odata-version':'3.0'
          },
          body: JSON.stringify(body)
        }
      );

      if (response.ok) {
        const updatedSettlements = [...settlements];
        updatedSettlements[index].status = "Settled";
        setSettlements(updatedSettlements);

        const updatedParticipants = participants.map(p => {
          if (p.Id === settlement.fromUserId) return { ...p, NetBalance: p.NetBalance + settlement.amount };
          if (p.Id === settlement.toUserId) return { ...p, NetBalance: p.NetBalance - settlement.amount };
          return p;
        });
        setParticipants(updatedParticipants);
        alert("Settlement saved successfully!");
      } else {
        const err = await response.json();
        alert("Error saving settlement: " + (err.error?.message || response.statusText));
      }
    } catch (e: any) {
      alert("Error saving settlement: " + e.message);
    }
  };

  const saveExpenseUpdate = async (expense: Expense) => {
    try {
      const siteUrl = context.pageContext.web.absoluteUrl;
      const body = { Amount: expense.Amount };

      const response = await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('TripExpenses')/items(${expense.Id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            "odata-version":"",
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          },
          body: JSON.stringify(body)
        }
      );

      if (response.ok) {
        alert(`Expense "${expense.Title}" updated successfully!`);
        setEditingRow(null);
        fetchData();
      } else {
        const err = await response.json();
        alert('Error updating expense: ' + (err.error?.message || response.statusText));
      }
    } catch (e: any) {
      alert('Error updating expense: ' + e.message);
    }
  };

  const deleteExpense = async (exp: Expense) => {
    if (!window.confirm(`Are you sure you want to delete "${exp.Title}"?`)) return;
    try {
      const siteUrl = context.pageContext.web.absoluteUrl;
      const response = await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('TripExpenses')/items(${exp.Id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE',
            'odata-version':""
          }
        }
      );
      if (response.ok) {
        alert(`Expense "${exp.Title}" deleted successfully!`);
        fetchData();
      } else {
        const err = await response.json();
        alert('Error deleting expense: ' + (err.error?.message || response.statusText));
      }
    } catch (e: any) {
      alert('Error deleting expense: ' + e.message);
    }
  };

  return (
    <div className="container mt-3">
      <button className="btn btn-secondary mb-3" onClick={onBack}>← Back to Trips</button>
      <h2 className="mb-3">Summary for <span className="text-primary">{trip.Title}</span></h2>

      {trip.Status !== 'Trip Closed' && !readOnly && (
        <button className="btn btn-danger mb-3" onClick={closeTrip}>Close Trip</button>
      )}

      {loading && <div className="alert alert-info">Loading summary...</div>}
      {error && <div className="alert alert-danger">{error}</div>}

      {!loading && !selectedParticipant && (
        <table className="table table-bordered table-striped table-hover">
          <thead className="table-dark">
            <tr>
              <th>Participant</th>
              <th>Total Package</th>
              <th>Total Paid</th>
              <th>Net Balance</th>
              <th>Owes / Gets</th>
            </tr>
          </thead>
          <tbody>
            {participants.map(p => (
              <tr key={p.Id}>
                <td
                  style={{ cursor: !readOnly && trip.Status !== 'Trip Closed' ? 'pointer' : 'default' }}
                  className="text-primary"
                  onClick={() => participantClick(p)}
                >
                  {p.Title}
                </td>
                <td>₹{p.Share.toFixed(2)}</td>
                <td>₹{p.TotalPaid.toFixed(2)}</td>
                <td className={p.NetBalance < 0 ? 'text-danger' : 'text-success'}>₹{p.NetBalance.toFixed(2)}</td>
                <td className={p.NetBalance < 0 ? 'text-danger' : 'text-success'}>
                  {p.NetBalance < 0 ? `Owes ₹${Math.abs(p.NetBalance).toFixed(2)}` : `Gets ₹${p.NetBalance.toFixed(2)}`}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}

      {selectedParticipant && (
        <div className="mt-4">
          <button className="btn btn-outline-secondary mb-3" onClick={closeDetail}>← Back to Summary</button>
          <h3 className="mb-3">Expenses by <span className="text-info">{selectedParticipant.Title}</span></h3>
          <table className="table table-bordered table-striped table-hover">
            <thead className="table-dark">
              <tr>
                <th>Expense</th>
                <th>Amount</th>
              </tr>
            </thead>
            <tbody>
              {expenses.filter(exp => exp.PaidById === selectedParticipant.Id).map(exp => (
                <tr key={exp.Id}>
                  <td>{exp.Title}</td>
                  <td>₹{exp.Amount.toFixed(2)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {!loading && expenses.length > 0 && (
        <div className="mt-4">
          <h4>All Expenses for this Trip</h4>
          <table className="table table-bordered table-striped table-hover">
            <thead className="table-dark">
              <tr>
                <th>Participant</th>
                <th>Expense</th>
                <th>Amount</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody>
              {expenses.map((exp, idx) => (
                <tr key={exp.Id}>
                  <td>{exp.PaidByName}</td>
                  <td>{exp.Title}</td>
                  <td>
                    {editingRow === exp.Id ? (
                      <input
                        type="number"
                        value={exp.Amount}
                        className="form-control"
                        onChange={(e) => {
                          const newAmount = Number(e.target.value);
                          const updatedExpenses = [...expenses];
                          updatedExpenses[idx] = { ...updatedExpenses[idx], Amount: newAmount };
                          setExpenses(updatedExpenses);
                        }}
                      />
                    ) : (
                      `₹${exp.Amount.toFixed(2)}`
                    )}
                  </td>
                  <td>
                    {editingRow === exp.Id ? (
                      <>
                        <button className="btn btn-success btn-sm me-2" onClick={() => saveExpenseUpdate(exp)}>Save</button>
                        <button className="btn btn-danger btn-sm" onClick={() => deleteExpense(exp)}>Delete</button>
                      </>
                    ) : (
                      <button className="btn btn-primary btn-sm" onClick={() => setEditingRow(exp.Id)}>Edit</button>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {trip.Status === 'Trip Closed' && (
        <div className="mt-4">
          <h4>💰 Settlements</h4>
          {settlements.length > 0 ? (
            <table className="table table-bordered table-striped">
              <thead className="table-dark">
                <tr>
                  <th>From</th>
                  <th>To</th>
                  <th>Amount</th>
                  <th>Status</th>
                  <th>Settlements
</th>
                </tr>
              </thead>
              <tbody>
                {settlements.map((s, idx) => (
                  <tr key={idx}>
                    <td>{s.from}</td>
                    <td>{s.to}</td>
                    <td>₹{s.amount.toFixed(2)}</td>
                    <td>{s.status}</td>
                    <td>
                      {s.status === "Pending" ? (
                        <input type="checkbox" onChange={() => saveSettlement(s, idx)} />
                      ) : (
                        <span className="text-success">✔ Settled</span>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          ) : (
            <div className="alert alert-info">No settlements calculated yet.</div>
          )}
        </div>
      )}

      {trip.Status === 'Trip Closed' && (
        <div className="alert alert-secondary mt-4">
          Trip Closed — summary is read-only.
        </div>
      )}
    </div>
  );
};

export default Summary;