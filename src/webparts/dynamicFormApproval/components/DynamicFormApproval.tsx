import * as React from 'react';
import { useState } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import Dashboard from './Dashboard';
import DynamicForm from './DynamicForm';
import MySubmissions from './MySubmissions';
import MyApprovals from './MyApprovals';
import ApplyForm from './ApplyForm';
import { Screen } from './Screen';
import { SPHttpClient } from '@microsoft/sp-http';

interface IDynamicFormApprovalProps {
  context: WebPartContext;
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}


export interface ISubmissionItem {
  Id: number;
  Title: string;
  FormType: string;
  RelatedItemId: number;
  Status: string;
  CurrentApprovalLevel: number;
  ReferenceName?: string;
  Created?: string;
  CreatedBy?: {
    Title: string;
    Email: string;
  };
}
export interface IApprovalItem {
  Level1ApproverId: any;
  Level2ApproverId: any;
  Level3ApproverId: any;
  Level1Approver: string;
  Level1Action: string;
  Level1Comments: string;
  Level1ActionDate: any;
  Level2Approver: string;
  Level2Action: string;
  Level2Comments: string;
  Level2ActionDate: any;
  Level3Approver: string;
  Level3Action: string;
  Level3Comments: string;
  Level3ActionDate: any;
  Created: string | number | Date;
  Id: number;
  Title: string;
  FormType: string;
  CurrentApprovalLevel: number;
  Status: string;
  RelatedItemId: number;
  CreatedBy: {
    Title: string;
    Email: string;
  };
}

type FormType = 'PurchaseOrder' | 'Invoice' | 'EmployeeInfo';

const DynamicFormApproval: React.FC<IDynamicFormApprovalProps> = ({ context }) => {
  const [currentScreen, setCurrentScreen] = useState<Screen>('dashboard');
  const [selectedFormType, setSelectedFormType] = useState<FormType | null>(null);
  const [selectedApproval, setSelectedApproval] = useState<IApprovalItem | null>(null);

  const handleViewChange = (view: Screen, formType?: FormType) => {
    setCurrentScreen(view);
    if (formType) setSelectedFormType(formType);
  };

  function setSelectedSubmission(item: ISubmissionItem) {
    console.log("Selected Submission:", item);
  }

  // 🔹 Send mail helper
  const sendMail = async (to: string[], subject: string, body: string) => {
    const siteUrl = context.pageContext.web.absoluteUrl;
    const url = `${siteUrl}/_api/SP.Utilities.Utility.SendEmail`;
    const response = await context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        "Accept": "application/json;odata=nometadata",
        "Content-type": "application/json;odata=nometadata"
      },
      body: JSON.stringify({
        properties: {
          To: { results: to },
          Subject: subject,
          Body: body
        }
      })
    });
    if (!response.ok) {
      console.error("Mail send failed", await response.text());
    }
  };

  // 🔹 Handle approval action complete (after ApplyForm updates list)
  const handleActionComplete = async () => {
    if (!selectedApproval) return;

    const siteUrl = context.pageContext.web.absoluteUrl;

    // Get latest item from FormSubmissions
    const url = `${siteUrl}/_api/web/lists/getbytitle('FormSubmissions')/items(${selectedApproval.Id})?$expand=Author`;
    const res = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    const updatedItem: any = await res.json();

    const status = updatedItem.Status;
    const formType = updatedItem.FormType;

    // Case 1: Approved & next approver exists
    if (status === "Pending" && updatedItem.CurrentApprovalLevel <= 3) {
      const nextLevel = updatedItem.CurrentApprovalLevel; // Use incremented level

      // Fetch ApprovalMatrix for the next approver
      const matrixUrl = `${siteUrl}/_api/web/lists/getbytitle('ApprovalMatrix')/items?$filter=FormType eq '${formType}' and Level eq ${nextLevel}&$expand=ManagerLevel1,ManagerLevel2,ManagerLevel3`;
      const matrixRes = await context.spHttpClient.get(matrixUrl, SPHttpClient.configurations.v1);
      const matrixJson = await matrixRes.json();

      let approverEmail = '';

      if (nextLevel === 1) approverEmail = matrixJson.value[0]?.ManagerLevel1?.EMail;
      else if (nextLevel === 2) approverEmail = matrixJson.value[0]?.ManagerLevel2?.EMail;
      else if (nextLevel === 3) approverEmail = matrixJson.value[0]?.ManagerLevel3?.EMail;

      if (approverEmail) {
        await sendMail(
          [approverEmail],
          `Approval Required - ${formType}`,
          `Please review and approve the submission: ${updatedItem.Title}`
        );
      }
    }

    // Case 2: Final Approval/Rejection → Notify submitter
    if (status === "Approved" || status === "Rejected") {
      const submitterEmail = updatedItem.Author?.Email;
      if (submitterEmail) {
        await sendMail(
          [submitterEmail],
          `Your submission has been ${status}`,
          `Your ${formType} (${updatedItem.Title}) has been ${status}.`
        );
      }
    }

    // Refresh approvals screen
    handleViewChange("myApprovals");
  };

  return (
    <div style={{ display: 'flex' }}>
      {currentScreen === 'dashboard' && (
        <Dashboard
          context={context}
          onViewChange={handleViewChange}
          onSelectSubmission={(item: ISubmissionItem) => {
            setSelectedSubmission(item);
            handleViewChange('mySubmissions');
          }}
          onSelectApproval={(item: IApprovalItem) => {
            setSelectedApproval(item);
            handleViewChange('applyForm');
          }}
          onScreenChange={() => {}}
          onSubmitSuccess={() => {}}
          onActionComplete={handleActionComplete}
          onApproveSelect={(item: IApprovalItem) => {
            setSelectedApproval(item);
            handleViewChange('applyForm');
          }}
        />
      )}

      <div style={{ flex: 1, padding: '1rem' }}>
        {currentScreen === 'dynamicForm' && selectedFormType && (
          <DynamicForm
            context={context}
            formType={selectedFormType}
            onSubmitSuccess={() => handleViewChange('mySubmissions')}
            onBack={() => handleViewChange('dashboard')}
          />
        )}

        {currentScreen === 'mySubmissions' && (
          <MySubmissions
            context={context}
            onBack={() => handleViewChange('dashboard')}
            onSelect={(item: ISubmissionItem) => setSelectedSubmission(item)}
          />
        )}

        {currentScreen === 'myApprovals' && (
          <MyApprovals
            context={context}
            onSelect={(item: IApprovalItem) => {
              setSelectedApproval(item);
              handleViewChange('applyForm');
            } }
            onActionComplete={handleActionComplete}
            onApproveSelect={(item: IApprovalItem) => {
              setSelectedApproval(item);
              handleViewChange('applyForm');
            } } onBack={()=>handleViewChange('dashboard') }   />
        )}

        {currentScreen === 'applyForm' && selectedApproval && (
          <ApplyForm
            context={context}
            item={selectedApproval}
            onActionComplete={handleActionComplete}
            onBack={() => handleViewChange('myApprovals')} onSubmitSuccess={function (): void {
              throw new Error('Function not implemented.');
            } }          />
        )}
      </div>
    </div>
  );
};

export default DynamicFormApproval;