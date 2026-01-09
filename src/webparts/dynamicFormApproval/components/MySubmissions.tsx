import * as React from 'react';
import { useEffect, useState } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';

interface IMySubmissionsProps {
  context: WebPartContext;
  onBack: () => void;
  refreshTrigger?: number;
  onSelect?: (item: ISubmissionItem) => void;
}

interface ISubmissionItem {
  Id: number;
  Title: string;
  FormType: string;
  RelatedItemId: number;
  Status: string;
  CurrentApprovalLevel: number;
}

const MySubmissions: React.FC<IMySubmissionsProps> = ({ context, onBack }) => {
  const [items, setItems] = useState<ISubmissionItem[]>([]);
  const [selectedItem, setSelectedItem] = useState<ISubmissionItem | null>(null);
  const [formData, setFormData] = useState<any>(null);
  const [loading, setLoading] = useState<boolean>(true);
  const [loadingDetails, setLoadingDetails] = useState<boolean>(false);
  const [editMode, setEditMode] = useState<boolean>(false);

  // Helper: get list info by form type
  const getListInfoForFormType = (formType: string) => {
    switch (formType) {
      case 'Invoice':
        return { listName: 'InvoiceList', select: 'Id,InvoiceNumber,Customer,Amount,Date,Status,CurrentApprovalLevel' };
      case 'PurchaseOrder':
        return { listName: 'purchaseorderlist', select: 'Id,Title,Vender,Amount,Date,Status,CurrentApprovalLevel' };
      case 'EmployeeInfo':
        return { listName: 'EmployeeInfoList', select: 'Id,Title,Name,Department,JoiningDate,Status,CurrentApprovalLevel' };
      default:
        return { listName: 'InvoiceList', select: 'Id,InvoiceNumber,Customer,Amount,Date,Status,CurrentApprovalLevel' };
    }
  };

  // Fetch FormSubmissions and merge Master list Approval Level
  useEffect(() => {
    const fetchData = async () => {
      try {
        setLoading(true);
        const url = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('FormSubmissions')/items?$select=Id,Title,FormType,RelatedItemId,Status,CurrentApprovalLevel`;
        const response = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        const data = await response.json();
        const submissions: ISubmissionItem[] = data.value || [];

        // Fetch Master list Approval Level for each item
        const updatedItems = await Promise.all(
          submissions.map(async (item) => {
            const info = getListInfoForFormType(item.FormType);
            try {
              const masterRes = await context.spHttpClient.get(
                `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${info.listName}')/items(${item.RelatedItemId})?$select=CurrentApprovalLevel`,
                SPHttpClient.configurations.v1
              );
              const masterData = await masterRes.json();
              return { ...item, CurrentApprovalLevel: masterData.CurrentApprovalLevel };
            } catch {
              return item;
            }
          })
        );

        setItems(updatedItems);
      } catch (err) {
        console.error('Error fetching submissions:', err);
      } finally {
        setLoading(false);
      }
    };
    fetchData();
  }, [context]);

  // Fetch full details for selected item
  const fetchOriginalItem = async (formType: string, relatedId: number) => {
    if (!relatedId) return null;
    try {
      setLoadingDetails(true);
      const info = getListInfoForFormType(formType);
      const url = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${info.listName}')/items(${relatedId})?$select=${info.select}`;
      const response = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await response.json();
      return data;
    } catch (err) {
      console.error(`Error fetching ${formType} details:`, err);
      return null;
    } finally {
      setLoadingDetails(false);
    }
  };

  const saveChanges = async () => {
    if (!selectedItem || !formData) return;
    try {
      const info = getListInfoForFormType(selectedItem.FormType);

      // Update Master list
      const updateUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${info.listName}')/items(${selectedItem.RelatedItemId})`;
      await context.spHttpClient.post(updateUrl, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        },
        body: JSON.stringify(formData)
      });

      // Update FormSubmissions
      const formSubUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('FormSubmissions')/items(${selectedItem.Id})`;
      await context.spHttpClient.post(formSubUrl, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        },
        body: JSON.stringify({
          Title: formData.Title || selectedItem.Title,
          Status: formData.Status || selectedItem.Status,
          CurrentApprovalLevel: formData.CurrentApprovalLevel
        })
      });

      alert("Changes saved successfully ✅");
      setEditMode(false);
    } catch (err) {
      console.error("Error saving changes:", err);
      alert("Error saving changes ❌");
    }
  };

  const deleteItem = async () => {
    if (!selectedItem) return;
    if (!window.confirm("Are you sure you want to delete this item?")) return;

    try {
      const info = getListInfoForFormType(selectedItem.FormType);

      // Delete from Master list
      const deleteUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${info.listName}')/items(${selectedItem.RelatedItemId})`;
      await context.spHttpClient.post(deleteUrl, SPHttpClient.configurations.v1, {
        headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE' }
      });

      // Delete from FormSubmissions
      const formSubUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('FormSubmissions')/items(${selectedItem.Id})`;
      await context.spHttpClient.post(formSubUrl, SPHttpClient.configurations.v1, {
        headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE' }
      });

      alert("Item deleted successfully 🗑️");
      setSelectedItem(null);
      setItems(items.filter(i => i.Id !== selectedItem.Id));
    } catch (err) {
      console.error("Error deleting item:", err);
      alert("Error deleting item ❌");
    }
  };

  const handleChange = (field: string, value: any) => {
    setFormData({ ...formData, [field]: value });
  };

  useEffect(() => {
    if (selectedItem?.RelatedItemId) {
      fetchOriginalItem(selectedItem.FormType, selectedItem.RelatedItemId)
        .then(data => setFormData(data))
        .catch(err => console.error(err));
      setEditMode(false);
    }
  }, [selectedItem]);

  if (loading) return <div>Loading submissions...</div>;

  // Details page
  if (selectedItem && formData) {
    return (
      <div className="container mt-3">
        <h4>{selectedItem.FormType} Details</h4>
        {loadingDetails && <div>Loading details...</div>}

        {/* Invoice */}
        {selectedItem.FormType === 'Invoice' && (
          <>
            <div className="mb-2">
              <label>Invoice Number</label>
              <input className="form-control" value={formData.InvoiceNumber || ''} disabled={!editMode} onChange={e => handleChange('InvoiceNumber', e.target.value)} />
            </div>
            <div className="mb-2">
              <label>Customer</label>
              <input className="form-control" value={formData.Customer || ''} disabled={!editMode} onChange={e => handleChange('Customer', e.target.value)} />
            </div>
            <div className="mb-2">
              <label>Amount</label>
              <input type="number" className="form-control" value={formData.Amount || ''} disabled={!editMode} onChange={e => handleChange('Amount', e.target.value)} />
            </div>
            <div className="mb-2">
              <label>Date</label>
              <input type="date" className="form-control" value={formData.Date?.split('T')[0] || ''} disabled={!editMode} onChange={e => handleChange('Date',e.target.value)} />
            </div>
          </>
        )}

        {/* PurchaseOrder */}
        {selectedItem.FormType === 'PurchaseOrder' && (
          <>
            <div className="mb-2">
              <label>Title</label>
              <input className="form-control" value={formData.Title || ''} disabled={!editMode} onChange={e => handleChange('Title', e.target.value)} />
            </div>
            <div className="mb-2">
              <label>Vendor</label>
              <input className="form-control" value={formData.Vender || ''} disabled={!editMode} onChange={e => handleChange('Vender', e.target.value)} />
            </div>
            <div className="mb-2">
              <label>Amount</label>
              <input type="number" className="form-control" value={formData.Amount || ''} disabled={!editMode} onChange={e => handleChange('Amount', e.target.value)} />
            </div>
            <div className="mb-2">
              <label>Date</label>
              <input type="date" className="form-control" value={formData.Date?.split('T')[0] || ''} disabled={!editMode} onChange={e => handleChange('Date', e.target.value)} />
            </div>
          </>
        )}

        {/* EmployeeInfo */}
        {selectedItem.FormType === 'EmployeeInfo' && (
          <>
            <div className="mb-2">
              <label>Employee Name</label>
              <input className="form-control" value={formData.Name || ''} disabled={!editMode} onChange={e => handleChange('Name', e.target.value)} />
            </div>
            <div className="mb-2">
              <label>Department</label>
              <input className="form-control" value={formData.Department || ''} disabled={!editMode} onChange={e => handleChange('Department', e.target.value)} />
            </div>
            <div className="mb-2">
              <label>Joining Date</label>
              <input type="date" className="form-control" value={formData.JoiningDate?.split('T')[0] || ''} disabled={!editMode} onChange={e => handleChange('JoiningDate', e.target.value)} />
            </div>
          </>
        )}

        {/* Action buttons */}
        <div className="mt-3">
          {!editMode && <button className="btn btn-warning me-2" onClick={() => setEditMode(true)}>Edit</button>}
          {editMode && <button className="btn btn-success me-2" onClick={saveChanges}>Save</button>}
          <button className="btn btn-danger me-2" onClick={deleteItem}>Delete</button>
          <button className="btn btn-secondary" onClick={() => setSelectedItem(null)}>Back</button>
        </div>
      </div>
    );
  }

  // Submissions list
  return (
    <div className="container mt-3">
      <h4>My Submissions</h4>
      <table className="table table-bordered mt-3">
        <thead>
          <tr>
            <th>ID</th>
            <th>Title</th>
            <th>Form Type</th>
            <th>Status</th>
            <th>Approval Level</th>
          </tr>
        </thead>
        <tbody>
          {items.map(item => (
            <tr key={item.Id} style={{ cursor: "pointer" }} onClick={() => setSelectedItem(item)}>
              <td>{item.Id}</td>
              <td>{item.Title}</td>
              <td>{item.FormType}</td>
              <td>{item.Status}</td>
              <td>{item.CurrentApprovalLevel}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export default MySubmissions;