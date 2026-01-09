import * as React from 'react';
import { useState } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import 'bootstrap/dist/css/bootstrap.min.css';

interface IDynamicFormProps {
  context: WebPartContext;
  formType: 'PurchaseOrder' | 'Invoice' | 'EmployeeInfo';
  onBack: () => void;
  onSubmitSuccess: () => void;
}

const DynamicForm: React.FC<IDynamicFormProps> = ({ context, formType, onBack, onSubmitSuccess }) => {
  const [formData, setFormData] = useState<any>({});
  const [errors, setErrors] = useState<any>({});
  const [loading, setLoading] = useState(false);

  // Today's date in yyyy-mm-dd format
  const today = new Date().toISOString().split('T')[0];

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const { name, value } = e.target;
    setFormData({ ...formData, [name]: value });

    // Clear error for the field as user types
    if (errors[name]) {
      setErrors({ ...errors, [name]: '' });
    }
  };

  const validate = (): boolean => {
    const newErrors: any = {};
    if (formType === 'PurchaseOrder') {
      if (!formData.Vender) newErrors.Vender = 'Vendor is required';
      if (!formData.Amount) newErrors.Amount = 'Amount is required';
      if (!formData.Date) newErrors.Date = 'Date is required';
      else if (formData.Date < today) newErrors.Date = 'Date cannot be earlier than today';
    } else if (formType === 'Invoice') {
      if (!formData.InvoiceNumber) newErrors.InvoiceNumber = 'Invoice Number is required';
      if (!formData.Customer) newErrors.Customer = 'Customer is required';
      if (!formData.Amount) newErrors.Amount = 'Amount is required';
      if (!formData.Date) newErrors.Date = 'Date is required';
      else if (formData.Date < today) newErrors.Date = 'Date cannot be earlier than today';
    } else if (formType === 'EmployeeInfo') {
      if (!formData.Name) newErrors.Name = 'Employee Name is required';
      if (!formData.Department) newErrors.Department = 'Department is required';
      if (!formData.JoiningDate) newErrors.JoiningDate = 'Joining Date is required';
      else if (formData.JoiningDate < today) newErrors.JoiningDate = 'Joining Date cannot be earlier than today';
    }

    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const handleSubmit = async () => {
    if (!validate()) return;
    setLoading(true);

    // === KEEP ALL YOUR LOGIC BELOW UNCHANGED ===
    let listName = '';
    if (formType === 'PurchaseOrder') listName = 'purchaseorderlist';
    else if (formType === 'Invoice') listName = 'InvoiceList';
    else if (formType === 'EmployeeInfo') listName = 'EmployeeInfoList';

    const currentUserName = context.pageContext.user.displayName;

    const body: any = {
      ReferenceName: currentUserName,
      Title:
        formType +
        ' - ' +
        (formData.InvoiceNumber ||
          formData.Customer ||
          formData.Name ||
          new Date().toISOString()),
    };

    if (formType === 'PurchaseOrder') {
      body.Vender = formData.Vender;
      body.Amount = formData.Amount;
      body.Date = formData.Date;
      body.Status = 'Pending';
      body.CurrentApprovalLevel = 1;
    } else if (formType === 'Invoice') {
      body.InvoiceNumber = formData.InvoiceNumber;
      body.Customer = formData.Customer;
      body.Amount = formData.Amount;
      body.Date = formData.Date;
      body.Status = 'Pending';
      body.CurrentApprovalLevel = 1;
    } else if (formType === 'EmployeeInfo') {
      body.Title = formData.Name;
      body.Name = formData.Name;
      body.Department = formData.Department;
      body.JoiningDate = formData.JoiningDate;
      body.Status = 'Pending';
      body.CurrentApprovalLevel = 1;
    }

    try {
      const res: SPHttpClientResponse = await context.spHttpClient.post(
        `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': '',
          },
          body: JSON.stringify(body),
        }
      );

      if (!res.ok) {
        const err = await res.json();
        throw new Error(JSON.stringify(err));
      }

      const savedItem = await res.json();

      const submissionBody: any = {
        Title: formType + ' Submission - ' + savedItem.Id,
        FormType: formType,
        ReferenceName: currentUserName,
        Status: 'Pending',
        CurrentApprovalLevel: 1,
        RelatedItemId: savedItem.Id,
      };

      const subRes = await context.spHttpClient.post(
        `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('FormSubmissions')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': '',
          },
          body: JSON.stringify(submissionBody),
        }
      );

      if (!subRes.ok) {
        const err = await subRes.json();
        throw new Error(
          'FormSubmissions save failed: ' + JSON.stringify(err)
        );
      }

      const sendFirstApproverEmail = async () => {
        try {
          const matrixUrl =
            `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ApprovalMatrix')/items?` +
            `$select=Id,FormType,Level,ManagerLevel1/EMail` +
            `&$expand=ManagerLevel1` +
            `&$filter=FormType eq '${formType}' and Level eq 1`;

          const resMatrix = await context.spHttpClient.get(matrixUrl, SPHttpClient.configurations.v1, {
            headers: { Accept: 'application/json;odata=verbose', 'odata-version': '' }
          });

          if (!resMatrix.ok) throw new Error('ApprovalMatrix fetch failed');

          const data = await resMatrix.json();
          const matrix = data.d?.results?.[0] || data.value?.[0];

          if (matrix?.ManagerLevel1?.EMail) {
            await context.spHttpClient.post(
              `${context.pageContext.web.absoluteUrl}/_api/SP.Utilities.Utility.SendEmail`,
              SPHttpClient.configurations.v1,
              {
                headers: {
                  Accept: 'application/json;odata=verbose',
                  'Content-Type': 'application/json;odata=verbose',
                  'odata-version': ''
                },
                body: JSON.stringify({
                  properties: {
                    __metadata: { type: 'SP.Utilities.EmailProperties' },
                    To: { results: [matrix.ManagerLevel1.EMail] },
                    Subject: `${formType} ${savedItem.Title} awaiting your approval`,
                    Body: `<p>Please review ${formType} "${savedItem.Title}".</p>` +
                          `<p><a href="${context.pageContext.web.absoluteUrl}/SitePages/DynamicFormApproval.aspx?formType=${formType}&itemId=${savedItem.Id}">Click here to approve</a></p>`
                  }
                })
              }
            );
          }
        } catch (err) {
          console.error('First approver email error:', err);
        }
      };

      await sendFirstApproverEmail();

      alert('✅ Form saved Successfully.');
      onSubmitSuccess();
      onBack();
    } catch (error: any) {
      alert('❌ Error: ' + error.message);
      console.error(error);
    }

    setLoading(false);
  };

  // === Helper to render input fields consistently with label on right ===
  const renderField = (label: string, name: string, type: string = 'text') => (
    <div className="row mb-3 align-items-center">
      <div className="col-sm-9 order-1">
        <input
          type={type}
          className={`form-control ${errors[name] ? 'is-invalid' : ''}`}
          name={name}
          value={formData[name] || ''}
          onChange={handleChange}
          min={type === 'date' ? today : undefined}
        />
        {errors[name] && <div className="invalid-feedback">{errors[name]}</div>}
      </div>
      <label className="col-sm-3 col-form-label text-end fw-bold order-0">{label}</label>
    </div>
  );

  return (
    <div className="container mt-3">
      <h3 className="mb-4">{formType} Form</h3>

      {formType === 'PurchaseOrder' && (
        <>
          {renderField('Vendor', 'Vender')}
          {renderField('Amount', 'Amount', 'number')}
          {renderField('Date', 'Date', 'date')}
        </>
      )}

      {formType === 'Invoice' && (
        <>
          {renderField('Invoice Number', 'InvoiceNumber')}
          {renderField('Customer', 'Customer')}
          {renderField('Amount', 'Amount', 'number')}
          {renderField('Date', 'Date', 'date')}
        </>
      )}

      {formType === 'EmployeeInfo' && (
        <>
          {renderField('Employee Name', 'Name')}
          {renderField('Department', 'Department')}
          {renderField('Joining Date', 'JoiningDate', 'date')}
        </>
      )}

      <div className="d-flex gap-2 mt-3">
        <button className="btn btn-primary" onClick={handleSubmit} disabled={loading}>
          {loading ? 'Submitting...' : 'Submit'}
        </button>
        <button className="btn btn-secondary" onClick={onBack}>
          Back
        </button>
      </div>
    </div>
  );
};

export default DynamicForm;