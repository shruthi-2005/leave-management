import * as React from 'react';
import { useState } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import MySubmissions from './MySubmissions';
import DynamicForm from './DynamicForm';
import MyApprovals from './MyApprovals';
import { Screen } from './Screen';
import 'bootstrap/dist/css/bootstrap.min.css';
import Dropdown from 'react-bootstrap/Dropdown';
import { ISubmissionItem, IApprovalItem } from './DynamicFormApproval';

type FormType = 'PurchaseOrder' | 'Invoice' | 'EmployeeInfo';

interface IDashboardProps {
  context: WebPartContext;
  onViewChange: (view: Screen, formType?: FormType) => void;
  onScreenChange: (screen: Screen, data?: any) => void;
  onSubmitSuccess: () => void;
  onActionComplete: () => void;
  onApproveSelect: (item: IApprovalItem) => void;
  onSelectSubmission: (item: ISubmissionItem) => void;
  onSelectApproval: (item: IApprovalItem) => void;
}

type TabType = 'submissions' | 'approvals' | 'pending' | 'dynamicForm' | 'submissionDetail' | 'approvalDetail';

const Dashboard: React.FC<IDashboardProps> = ({
  context,
  onSelectSubmission,
  onSelectApproval
}) => {
  const [activeTab, setActiveTab] = useState<TabType>('submissions'); // ✅ Default highlight
  const [formType, setFormType] = useState<FormType | null>(null);
  const [refreshCount, setRefreshCount] = useState<number>(0);

  const [selectedSubmission, setSelectedSubmission] = useState<ISubmissionItem | null>(null);
  const [selectedApproval, setSelectedApproval] = useState<IApprovalItem | null>(null);

  const handleAddNew = (selectedForm: FormType) => {
    setFormType(selectedForm);
    setActiveTab('dynamicForm');
  };

  const handleTabClick = (tab: TabType) => {
    setActiveTab(tab);
    setSelectedSubmission(null);
    setSelectedApproval(null);
  };

  return (
    <div className="d-flex">
      {/* Left Drawer / Menu */}
      <div className="bg-light border-end p-3" style={{ width: '220px', minHeight: '100vh' }}>
        <h5 className="mb-4">Menu</h5>
        <ul className="nav flex-column">
          <li className="nav-item">
            <button
              className={`nav-link btn btn-link text-start ${activeTab === 'submissions' ? 'fw-bold text-primary bg-white shadow-sm rounded' : 'text-dark'}`}
              onClick={() => handleTabClick('submissions')}
            >
              My Submissions
            </button>
          </li>
          <li className="nav-item">
            <button
              className={`nav-link btn btn-link text-start ${activeTab === 'approvals' ? 'fw-bold text-primary bg-white shadow-sm rounded' : 'text-dark'}`}
              onClick={() => handleTabClick('approvals')}
            >
              My Approvals
            </button>
          </li>
          
        </ul>

        {/* Add New Button Dropdown */}
        <Dropdown className="mt-4">
          <Dropdown.Toggle variant="primary" className="w-100">Add New</Dropdown.Toggle>
          <Dropdown.Menu className="w-100">
            <Dropdown.Item onClick={() => handleAddNew('PurchaseOrder')}>Purchase Order</Dropdown.Item>
            <Dropdown.Item onClick={() => handleAddNew('Invoice')}>Invoice</Dropdown.Item>
            <Dropdown.Item onClick={() => handleAddNew('EmployeeInfo')}>Employee Info</Dropdown.Item>
          </Dropdown.Menu>
        </Dropdown>
      </div>

      {/* Main Content */}
      <div className="flex-grow-1 p-4">
        {/* My Submissions */}
        {activeTab === 'submissions' && !selectedSubmission && (
          <MySubmissions
            context={context}
            onBack={() => {}}
            refreshTrigger={refreshCount}
            onSelect={(item: ISubmissionItem) => {
              setSelectedSubmission(item);
              setActiveTab('submissionDetail');
              onSelectSubmission(item);
            }}
          />
        )}

        {activeTab === 'submissionDetail' && selectedSubmission && (
          <div>
            <button className="btn btn-secondary mb-3" onClick={() => handleTabClick('submissions')}>
              ← Back
            </button>
            <div className="card p-3 shadow-sm">
              <h5>Submission Details</h5>
              <pre>{JSON.stringify(selectedSubmission, null, 2)}</pre>
            </div>
          </div>
        )}

        {/* Add New Form */}
        {activeTab === 'dynamicForm' && formType && (
          <DynamicForm
            context={context}
            formType={formType}
            onSubmitSuccess={() => {
              setRefreshCount(prev => prev + 1);
              handleTabClick('submissions');
            }}
            onBack={() => handleTabClick('submissions')}
          />
        )}

        {/* My Approvals */}
        {activeTab === 'approvals' && !selectedApproval && (
          <MyApprovals
            context={context}
            onActionComplete={() => setRefreshCount(prev => prev + 1)}
            onSelect={(item: IApprovalItem) => {
              setSelectedApproval(item);
              setActiveTab('approvalDetail');
              onSelectApproval(item);
            }}
            onApproveSelect={() => {}}
            onBack={() => handleTabClick('submissions')}
          />
        )}

        {activeTab === 'approvalDetail' && selectedApproval && (
          <div>
            <button className="btn btn-secondary mb-3" onClick={() => handleTabClick('approvals')}>
              ← Back
            </button>
            <div className="card p-3 shadow-sm">
              <h5>Approval Details</h5>
              <pre>{JSON.stringify(selectedApproval, null, 2)}</pre>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default Dashboard;