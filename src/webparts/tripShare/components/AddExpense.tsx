import * as React from 'react';
import { useState } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PeoplePicker, PrincipalType, IPeoplePickerContext } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import 'bootstrap/dist/css/bootstrap.min.css';

interface Trip {
  Id: number;
  Title: string;
  Participants: { Id: number; Title: string }[];
}

interface AddExpenseProps {
  context: WebPartContext;
  trip: Trip;
  tripId:number;
  siteUrl:string;
  onSave: () => void;
  onCancel: () => void;
  onTripCreated:(trip:{Id:number; Title:string; TotalPackage:number})=> void;
  onExpenseSaved:()=>void;
} 

interface IPickedUser {
  loginName: string;
  text: string;
}

const AddExpense: React.FC<AddExpenseProps> = ({ context, trip, onSave, onCancel,onExpenseSaved,tripId}) => {
  const [expenseTitle, setExpenseTitle] = useState('');
  const [expenseAmount, setExpenseAmount] = useState('');
  const [paidBy, setPaidBy] = useState<IPickedUser | null>(null);
  const [expenseDate, setExpenseDate] = useState('');
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState('');

  const handleSubmit = async () => {
    if (!expenseTitle || !expenseAmount || !paidBy || !expenseDate) {
      setError('Please fill all required fields.');
      return;
    }

    if (isNaN(Number(expenseAmount)) || Number(expenseAmount) <= 0) {
      setError('Please enter a valid expense amount.');  
      return; 
    }

    setSaving(true);
    setError('');

    let spUserId = 0;
    try {
      const spUserResponse = await context.spHttpClient.get(
        `${context.pageContext.web.absoluteUrl}/_api/web/siteusers(@v)?@v='${encodeURIComponent(paidBy.loginName)}'`,
        SPHttpClient.configurations.v1
      );
      const spUserData = await spUserResponse.json();
      spUserId = spUserData.Id;
    } catch (err: any) {
      setError('Error resolving user: ' + err.message);
      setSaving(false);
      return;
    }

    const isoDate = new Date(expenseDate).toISOString();

    const item: any = {
      Title: expenseTitle,
      TripId: trip.Id.toString(),
      TripName: trip.Title,
      Amount: Number(expenseAmount),
      Date: isoDate,
      SpentById: spUserId
    };

    try { 
      const response: SPHttpClientResponse = await context.spHttpClient.post(
        `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('TripExpenses')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': '',
          },
          body: JSON.stringify(item),
        }
      );

      if (response.ok) {
        onSave();
      } else {
        const err = await response.json();
        setError('Error saving expense: ' + (err.error?.message || response.statusText));
      }
    } catch (e: any) {
      setError('Error saving expense: ' + e.message);
    }

    setSaving(false);
  };

  const peoplePickerContext: IPeoplePickerContext = {
    spHttpClient: context.spHttpClient,
    msGraphClientFactory: context.msGraphClientFactory,
    absoluteUrl: context.pageContext.web.absoluteUrl
  };

  return (
    <div className="container mt-4">
      <div className="card shadow-lg border-0 rounded-3">
        <div className="card-header bg-primary text-white">
          <h4 className="mb-0">Add Expense to "{trip.Title}"</h4>
        </div>
        <div className="card-body p-4">

          {error && <div className="alert alert-danger">{error}</div>}

          <div className="mb-3">
            <label className="form-label fw-bold">Expense Details</label>
            <input
              type="text"
              className="form-control"
              placeholder="Enter expense details"
              value={expenseTitle}
              onChange={e => setExpenseTitle(e.target.value)}
            />
          </div>

          <div className="mb-3">
            <label className="form-label fw-bold">Expense Amount</label>
            <input
              type="number"
              className="form-control"
              placeholder="Enter amount"
              value={expenseAmount}
              onChange={e => setExpenseAmount(e.target.value)}
            />
          </div>

          <div className="mb-3">
            <label className="form-label fw-bold">Paid By</label>
            <div className="border rounded p-2">
              <PeoplePicker
                context={peoplePickerContext}
                personSelectionLimit={1}
                showtooltip={true}
                defaultSelectedUsers={paidBy ? [paidBy.loginName] : []}
                onChange={(items: any[]) =>
                  setPaidBy(items.length > 0 ? { loginName: items[0].loginName, text: items[0].text } : null)
                }
                principalTypes={[PrincipalType.User]}
                resolveDelay={500}
              />
            </div>
          </div>

          <div className="mb-4">
            <label className="form-label fw-bold">Expense Date</label>
            <input
              type="date"
              className="form-control"
              value={expenseDate}
              onChange={e => setExpenseDate(e.target.value)}
            />
          </div>

          <div className="d-flex justify-content-end gap-2">
            <button className="btn btn-secondary" onClick={onCancel} disabled={saving}>
              Cancel
            </button>
            <button className="btn btn-success px-4" onClick={handleSubmit} disabled={saving}>
              {saving ? 'Saving...' : 'Save'}
            </button>
          </div>

        </div>
      </div>
    </div>
  );
};

export default AddExpense;