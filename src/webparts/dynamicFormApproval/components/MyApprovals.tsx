import * as React from 'react';
import { useState, useEffect } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { IApprovalItem } from './DynamicFormApproval'; // shared interface

interface IProps {
  context: WebPartContext;
  onSelect: (item: IApprovalItem) => void;
  onBack: () => void; 
  onActionComplete: () => void;
  onApproveSelect: (item: IApprovalItem) => void;
}

const MyApprovals: React.FC<IProps> = ({ context, onSelect, onBack }) => {
  const [items, setItems] = useState<IApprovalItem[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [filter, setFilter] = useState<string>("All"); 

  useEffect(() => {
    loadApprovals();
  }, []);

  // 🔹 Master list mapping
  const getListInfoForFormType = (formType: string) => {
    switch (formType) {
      case 'Invoice':
        return { listName: 'InvoiceList', select: 'Id,CurrentApprovalLevel' };
      case 'PurchaseOrder':
        return { listName: 'PurchaseOrderList', select: 'Id,CurrentApprovalLevel' };
      case 'EmployeeInfo':
        return { listName: 'EmployeeInfoList', select: 'Id,CurrentApprovalLevel' };
      default:
        return null;
    }
  };

  const loadApprovals = async () => {
    setLoading(true);
    try {
      const me = context.pageContext.user.email.toLowerCase();

      const url = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('FormSubmissions')/items?` +
        `$select=Id,FormType,Status,ReferenceName,Created,RelatedItemId,Author/Title,Author/EMail&$expand=Author`;

      const r = await context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: { 'Accept': 'application/json;odata=nometadata','odata-version': '' }
      });
      if (!r.ok) throw new Error(`Failed to fetch items. Status: ${r.status}`);

      const data = await r.json();
      const approvals: IApprovalItem[] = [];

      for (const s of data.value) {
        if (!["Pending", "Approved", "Rejected"].includes(s.Status)) continue;

        // 🔹 Approval Matrix check
        const matrixUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ApprovalMatrix')/items?` +
          `$select=Id,FormType,Level,Approver/EMail,ManagerLevel1/EMail,ManagerLevel2/EMail,ManagerLevel3/EMail&$expand=Approver,ManagerLevel1,ManagerLevel2,ManagerLevel3` +
          `&$filter=FormType eq '${s.FormType}'`;

        const mRes = await context.spHttpClient.get(matrixUrl, SPHttpClient.configurations.v1, {
          headers: { 'Accept': 'application/json;odata=nometadata','odata-version': '' }
        });
        const mData = await mRes.json();
        const matrix = mData.value?.[0];
        if (!matrix) continue;

        const approversEmails: string[] = [];
        if (matrix.Approver?.EMail) approversEmails.push(matrix.Approver.EMail.toLowerCase());
        if (matrix.ManagerLevel1?.EMail) approversEmails.push(matrix.ManagerLevel1.EMail.toLowerCase());
        if (matrix.ManagerLevel2?.EMail) approversEmails.push(matrix.ManagerLevel2.EMail.toLowerCase());
        if (matrix.ManagerLevel3?.EMail) approversEmails.push(matrix.ManagerLevel3.EMail.toLowerCase());

        if (!approversEmails.includes(me)) continue;

        // 🔹 fetch CurrentApprovalLevel from master list
        const info = getListInfoForFormType(s.FormType);
        let masterLevel: number | null = null;

        if (info && s.RelatedItemId) {
          const masterUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${info.listName}')/items(${s.RelatedItemId})?$select=${info.select}`;
          const masterRes = await context.spHttpClient.get(masterUrl, SPHttpClient.configurations.v1, {
            headers: { 'Accept': 'application/json;odata=nometadata','odata-version': '' }
          });

          if (masterRes.ok) {
            const masterData = await masterRes.json();
            masterLevel = masterData.CurrentApprovalLevel || null;
          }
        }

        approvals.push({
          Id: s.Id,
          FormType: s.FormType,
          Status: s.Status,
          Created: s.Created,
          Title: s.ReferenceName || "",
          CurrentApprovalLevel: masterLevel ??0, // 🔹 master list value used here
          RelatedItemId: s.RelatedItemId,
          CreatedBy: { Title: s.Author?.Title || "", Email: s.Author?.EMail || "" },
          Level1Approver: '',
          Level1Action: '',
          Level1Comments: '',
          Level1ActionDate: undefined,
          Level2Approver: '',
          Level2Action: '',
          Level2Comments: '',
          Level2ActionDate: undefined,
          Level3Approver: '',
          Level3Action: '',
          Level3Comments: '',
          Level3ActionDate: undefined,
          Level1ApproverId: undefined,
          Level2ApproverId: undefined,
          Level3ApproverId: undefined
        });
      }

      setItems(approvals);
    } catch (err: any) {
      console.error("Load MyApprovals error:", err);
      alert("Load MyApprovals error: " + (err.message || err));
    } finally {
      setLoading(false);
    }
  };

  const filteredItems = filter === "All" ? items : items.filter(i => i.Status === filter);

  return (
    <div className="container mt-3">
      <button className="btn btn-secondary mb-3" onClick={onBack}>
        ← Back
      </button>

      <h4>My Approvals</h4>

      <div className="mb-3">
        <label>Status Filter: </label>
        <select className="form-select w-auto d-inline-block ms-2" value={filter} onChange={e => setFilter(e.target.value)}>
          <option value="All">All</option>
          <option value="Pending">Pending</option>
          <option value="Approved">Approved</option>
          <option value="Rejected">Rejected</option>
        </select>
      </div>

      {loading && <p>Loading...</p>}
      {!loading && filteredItems.length === 0 && <p>No approvals found.</p>}
      {!loading && filteredItems.length > 0 && (
        <table className="table table-bordered">
          <thead>
            <tr>
              <th>Form Type</th>
              <th>Level</th>
              <th>Date</th>
              <th>Status</th>
              <th>Reference Name</th>
              <th>Submitted By</th>
            </tr>
          </thead>
          <tbody>
            {filteredItems.map(i => (
              <tr key={i.Id} style={{ cursor: 'pointer' }} onClick={() => onSelect(i)}>
                <td>{i.FormType}</td>
                <td>{i.CurrentApprovalLevel}</td> {/* 🔹 Master list value */}
                <td>{new Date(i.Created).toLocaleDateString()}</td>
                <td>
                  <span className={`badge ${i.Status === "Pending" ? "bg-warning" : i.Status === "Approved" ? "bg-success" : "bg-danger"}`}>
                    {i.Status}
                  </span>
                </td>
                <td>{i.Title}</td>
                <td>{i.CreatedBy?.Title}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
};

export default MyApprovals;