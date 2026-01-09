import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IFormSubmission {
  Id?: number;
  Title: string;
  FormType: string;
  CurrentApprovalLevel: number;
  Status: string;
  CreatedById?: number;
  ReferenceName: string;
  RelatedItemId?: number;
}

export interface IApprovalMatrix {
  Id: number;
  Title: string;
  FormType: string;
  Level: number;
  ApproverId?: number;
  IsActive: boolean;
}

export default class FormService {
  private context: WebPartContext;
  private formSubmissionsList: string = "FormSubmissions";
  private approvalMatrixList: string = "ApprovalMatrix";

  constructor(context: WebPartContext) {
    this.context = context;
  }

  // Create FormSubmission
  public async createFormSubmission(item: IFormSubmission): Promise<any> {
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${this.formSubmissionsList}')/items`;
    const res = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata'
      },
      body: JSON.stringify(item)
    });
    return res.json();
  }

  // Get ApprovalMatrix approver by FormType + Level
  public async getApprover(formType: string, level: number): Promise<IApprovalMatrix | null> {
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${this.approvalMatrixList}')/items?$filter=FormType eq '${formType}' and Level eq ${level} and IsActive eq 1`;
    const res = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    const data = await res.json();
    return data.value.length > 0 ? data.value[0] : null;
  }

  // Update FormSubmission (approval/reject)
  public async updateFormSubmission(id: number, updateObj: any): Promise<any> {
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${this.formSubmissionsList}')/items(${id})`;
    const res = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify(updateObj)
    });
    return res.ok;
  }
}