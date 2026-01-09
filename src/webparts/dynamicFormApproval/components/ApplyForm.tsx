import * as React from 'react';
import { useEffect, useState } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { IApprovalItem } from './DynamicFormApproval';

interface IApplyFormProps {
  context: WebPartContext;
  item: IApprovalItem;
  onActionComplete: () => void;
  onBack: () => void;
  onSubmitSuccess:()=>void;
}

interface IRelatedItem {
  Id: number;
  Title?: string;
  InvoiceNumber?: string;
  Customer?: string;
  Vender?: string;
  Name?: string;
  Department?: string;
  Amount?: number;
  Date?: string;
  Status?: string;
  CurrentApprovalLevel?: number;
}

const ApplyForm: React.FC<IApplyFormProps> = ({ context, item, onActionComplete, onBack }) => {
  const [related, setRelated] = useState<IRelatedItem | null>(null);
  const [formSubmission, setFormSubmission] = useState<IApprovalItem | null>(null);
  const [comments, setComments] = useState<string>('');
  const siteUrl = context.pageContext.web.absoluteUrl;

  const getCurrentUserId = async (): Promise<number | null> => {
    try {
      const res = await context.spHttpClient.get(`${siteUrl}/_api/web/currentuser`, SPHttpClient.configurations.v1);
      const data = await res.json();
      return data.Id || null;
    } catch {
      return null;
    }
  };

  const getUserTitleById = async (id: number): Promise<string> => {
    try {
      const res = await context.spHttpClient.get(`${siteUrl}/_api/web/getuserbyid(${id})?$select=Title`, SPHttpClient.configurations.v1);
      const data = await res.json();
      return data?.Title || '-';
    } catch {
      return '-';
    }
  };

  const fetchApproverNames = async (formItem: IApprovalItem) => {
    const level1 = formItem.Level1ApproverId ? await getUserTitleById(formItem.Level1ApproverId) : '-';
    const level2 = formItem.Level2ApproverId ? await getUserTitleById(formItem.Level2ApproverId) : '-';
    const level3 = formItem.Level3ApproverId ? await getUserTitleById(formItem.Level3ApproverId) : '-';
    setFormSubmission({ ...formItem, Level1Approver: level1, Level2Approver: level2, Level3Approver: level3 });
  };

  const getListInfoForFormType = (formType: string) => {
    switch (formType) {
      case 'Invoice':
        return { listName: 'InvoiceList', select: 'Id,InvoiceNumber,Customer,Amount,Date,Status,CurrentApprovalLevel' };
      case 'PurchaseOrder':
        return { listName: 'purchaseorderlist', select: 'Id,Title,Vender,Amount,Date,Status,CurrentApprovalLevel,ReferenceName' };
      case 'EmployeeInfo':
        return { listName: 'EmployeeInfoList', select: 'Id,Title,Name,Department,JoiningDate,Status,CurrentApprovalLevel,ReferenceName' };
      default:
        return { listName: 'InvoiceList', select: 'Id,InvoiceNumber,Customer,Amount,Date,Status,CurrentApprovalLevel' };
    }
  };

  useEffect(() => {
    if (!item) return;

    const fetchFormSubmission = async () => {
      const resp = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('FormSubmissions')/items(${item.Id})`,
        SPHttpClient.configurations.v1
      );
      if (resp.ok) {
        const d = await resp.json();
        await fetchApproverNames(d);
      }
    };

    const fetchRelated = async () => {
      const info = getListInfoForFormType(item.FormType);
      if (!item.RelatedItemId) return;
      const url = `${siteUrl}/_api/web/lists/getbytitle('${info.listName}')/items(${item.RelatedItemId})?$select=${info.select}`;
      const response = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (response.ok) setRelated(await response.json());
    };

    fetchFormSubmission();
    fetchRelated();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [item]);

  const refreshAll = async () => {
    if (item?.RelatedItemId) {
      const info = getListInfoForFormType(item.FormType);
      const response = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${info.listName}')/items(${item.RelatedItemId})?$select=${info.select}`,
        SPHttpClient.configurations.v1
      );
      if (response.ok) setRelated(await response.json());
    }
    if (item?.Id) {
      const resp = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('FormSubmissions')/items(${item.Id})`,
        SPHttpClient.configurations.v1
      );
      if (resp.ok) {
        const d = await resp.json();
        await fetchApproverNames(d);
      }
    }
  };

  const sendEmail = async (to: string[], subject: string, body: string, relatedItemId?: number) => {
    let emailBody = body;
    if (relatedItemId) {
      const approvalLink = `${siteUrl}/SitePages/DynamicFormApproval.aspx?formType=${item.FormType}&itemId=${relatedItemId}`;
      emailBody += `<p><a href="${approvalLink}">Open approval page</a></p>`;
    }
    await context.spHttpClient.post(`${siteUrl}/_api/SP.Utilities.Utility.SendEmail`, SPHttpClient.configurations.v1, {
      headers: { Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', 'odata-version': '' },
      body: JSON.stringify({
        properties: {
          __metadata: { type: 'SP.Utilities.EmailProperties' },
          To: { results: to },
          Subject: subject,
          Body: emailBody
        }
      })
    });
  };

  const updateRelatedList = async (nextLevel: number | null, status: string | null = null) => {
    if (!item?.RelatedItemId) return;
    const info = getListInfoForFormType(item.FormType);
    const payload: any = {};
    if (nextLevel !== null) payload.CurrentApprovalLevel = nextLevel;
    if (status !== null) payload.Status = status;

    await context.spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${info.listName}')/items(${item.RelatedItemId})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE',
          'odata-version': ''
        },
        body: JSON.stringify(payload)
      }
    );
  };

  const handleApprove = async () => {
    try {
      const currentLevel = (related?.CurrentApprovalLevel ?? item.CurrentApprovalLevel ?? 1) as number;
      const now = new Date().toISOString();
      const userId = await getCurrentUserId();
      if (!userId) return alert('Unable to get current user ID.');

      const updateObj: any = {};
      if (currentLevel === 1) {
        updateObj.Level1ApproverId = userId;
        updateObj.Level1Action = 'Approved';
        updateObj.Level1Comments = comments;
        updateObj.Level1ActionDate = now;
      } else if (currentLevel === 2) {
        updateObj.Level2ApproverId = userId;
        updateObj.Level2Action = 'Approved';
        updateObj.Level2Comments = comments;
        updateObj.Level2ActionDate = now;
      } else {
        updateObj.Level3ApproverId = userId;
        updateObj.Level3Action = 'Approved';
        updateObj.Level3Comments = comments;
        updateObj.Level3ActionDate = now;
      }
      if (currentLevel === 3) updateObj.Status = 'Approved';

      await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('FormSubmissions')/items(${item.Id})`,
        SPHttpClient.configurations.v1,
        {
          headers: { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata', 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE', 'odata-version': '' },
          body: JSON.stringify(updateObj)
        }
      );

      if (currentLevel < 3) {
        const nextLevel = currentLevel + 1;

        // next approver email
        const matrixRes = await context.spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('ApprovalMatrix')/items?$select=ManagerLevel${nextLevel}/EMail&$expand=ManagerLevel${nextLevel}&$filter=FormType eq '${item.FormType}' and IsActive eq 1`,
          SPHttpClient.configurations.v1
        );
        const matrixData = await matrixRes.json();
        const mgrField = matrixData.value?.[0]?.[`ManagerLevel${nextLevel}`];
        const nextApproverEmail = Array.isArray(mgrField) ? mgrField[0]?.EMail : mgrField?.EMail;

        if (nextApproverEmail) {
          await sendEmail([nextApproverEmail], `Approval required - ${item.FormType}`, `<p>Please review ${item.FormType} (${related?.Title || item.Title || ''}).</p>`, item.RelatedItemId);
        }
        await updateRelatedList(nextLevel, 'Pending');
      } else {
        await updateRelatedList(3, 'Approved');
        const submitterEmail = item.CreatedBy?.Email || (formSubmission as any)?.CreatedBy?.Email;
        if (submitterEmail) {
          await sendEmail([submitterEmail], `${item.FormType} Approved`, `<p>Your ${item.FormType} (${related?.Title || item.Title || ''}) has been approved.</p>`, item.RelatedItemId);
        }
      }

      await refreshAll();
      alert('Approved successfully');
      onActionComplete();
    } catch (err: any) {
      alert('Error during approval: ' + err.message);
    }
  };

  const handleReject = async () => {
    try {
      const currentLevel = (related?.CurrentApprovalLevel ?? item.CurrentApprovalLevel ?? 1) as number;
      const now = new Date().toISOString();
      const userId = await getCurrentUserId();
      if (!userId) return alert('Unable to get current user ID.');

      const updateObj: any = {};
      if (currentLevel === 1) {
        updateObj.Level1ApproverId = userId;
        updateObj.Level1Action = 'Rejected';
        updateObj.Level1Comments = comments;
        updateObj.Level1ActionDate = now;
      } else if (currentLevel === 2) {
        updateObj.Level2ApproverId = userId;
        updateObj.Level2Action = 'Rejected';
        updateObj.Level2Comments = comments;
        updateObj.Level2ActionDate = now;
      } else {
        updateObj.Level3ApproverId = userId;
        updateObj.Level3Action = 'Rejected';
        updateObj.Level3Comments = comments;
        updateObj.Level3ActionDate = now;
      }
      updateObj.Status = 'Rejected';

      await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('FormSubmissions')/items(${item.Id})`,
        SPHttpClient.configurations.v1,
        {
          headers: { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata', 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE', 'odata-version': '' },
          body: JSON.stringify(updateObj)
        }
      );

      await updateRelatedList(null, 'Rejected');
      const submitterEmail = item.CreatedBy?.Email || (formSubmission as any)?.CreatedBy?.Email;
      if (submitterEmail) {
        await sendEmail([submitterEmail], `${item.FormType} Rejected`, `<p>Your ${item.FormType} (${related?.Title || item.Title || ''}) has been rejected.</p>`, item.RelatedItemId);
      }

      await refreshAll();
      alert('Rejected successfully');
      onActionComplete();
    } catch (err: any) {
      alert('Error during rejection: ' + err.message);
    }
  };

  if (!related || !formSubmission) return <div>Loading details...</div>;

  const renderRelatedSummary = () => {
    switch (item.FormType) {
      case 'Invoice':
        return (
          <>
            <p><b>Invoice Number:</b> {related.InvoiceNumber}</p>
            <p><b>Customer:</b> {related.Customer}</p>
            <p><b>Amount:</b> {related.Amount}</p>
            <p><b>Date:</b> {related.Date ? new Date(related.Date).toLocaleDateString() : '-'}</p>
          </>
        );
      case 'PurchaseOrder':
        return (
          <>
            <p><b>Vendor:</b> {related.Vender}</p>
            <p><b>Amount:</b> {related.Amount}</p>
            <p><b>Reference:</b> {related.Title}</p>
            <p><b>Status:</b> {related.Status}</p>
          </>
        );
      case 'EmployeeInfo':
        return (
          <>
            <p><b>Name:</b> {related.Name}</p>
            <p><b>Department:</b> {related.Department}</p>
            <p><b>Joining:</b> {related.Date ? new Date(related.Date).toLocaleDateString() : '-'}</p>
          </>
        );
      default:
        return null;
    }
  };

  return (
    <div className="card">
      <div className="card-body">
        <h4>{item.FormType} Details</h4>
        {renderRelatedSummary()}

        <p><b>Submission Status:</b> {formSubmission.Status}</p>
        <p><b>Approval Level (related):</b> {related.CurrentApprovalLevel}</p>

        {(formSubmission.Level1Action || formSubmission.Level2Action || formSubmission.Level3Action) && (
          <div className="mt-3">
            <h5>Approval History</h5>
            <table className="table table-bordered">
              <thead>
                <tr>
                  <th>Level</th>
                  <th>Approver</th>
                  <th>Action</th>
                  <th>Comments</th>
                  <th>Date</th>
                </tr>
              </thead>
              <tbody>
                {formSubmission.Level1Action && (
                  <tr>
                    <td>1</td>
                    <td>{formSubmission.Level1Approver || '-'}</td>
                    <td>{formSubmission.Level1Action}</td>
                    <td>{formSubmission.Level1Comments}</td>
                    <td>{formSubmission.Level1ActionDate ? new Date(formSubmission.Level1ActionDate).toLocaleString() : ''}</td>
                  </tr>
                )}
                {formSubmission.Level2Action && (
                  <tr>
                    <td>2</td>
                    <td>{formSubmission.Level2Approver || '-'}</td>
                    <td>{formSubmission.Level2Action}</td>
                    <td>{formSubmission.Level2Comments}</td>
                    <td>{formSubmission.Level2ActionDate ? new Date(formSubmission.Level2ActionDate).toLocaleString() : ''}</td>
                  </tr>
                )}
                {formSubmission.Level3Action && (
                  <tr>
                    <td>3</td>
                    <td>{formSubmission.Level3Approver || '-'}</td>
                    <td>{formSubmission.Level3Action}</td>
                    <td>{formSubmission.Level3Comments}</td>
                    <td>{formSubmission.Level3ActionDate ? new Date(formSubmission.Level3ActionDate).toLocaleString() : ''}</td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        )}

        {related.Status === 'Pending' && (
          <div className="mt-3">
            <textarea
              className="form-control mb-2"
              placeholder="Enter comments..."
              value={comments}
              onChange={e => setComments(e.target.value)}
            />
            <button className="btn btn-success me-2" onClick={handleApprove}>Approve</button>
            <button className="btn btn-danger me-2" onClick={handleReject}>Reject</button>
          </div>
        )}

        <button className="btn btn-secondary mt-2" onClick={onBack}>Back</button>
      </div>
    </div>
  );
};

export default ApplyForm; 