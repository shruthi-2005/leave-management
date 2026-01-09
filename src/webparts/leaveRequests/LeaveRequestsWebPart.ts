import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import LeaveDashboard from './components/LeaveDashboard';
import ApplyLeave from './components/ApplyLeave';
import LeaveRequests from './components/LeaveRequests';
import Leaves from './components/Leaves';
import { ViewType } from './components/ViewType';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ILeaveRequestsWebPartProps {
  context:WebPartContext;
}

export default class LeaveRequestsWebPart extends BaseClientSideWebPart<ILeaveRequestsWebPartProps> {
  private currentView: ViewType = ViewType.home;
  private selectedItem: any = null;

  private renderComponent(): void {
    let element: React.ReactElement<any> | undefined;

    switch (this.currentView) {
      case ViewType.home:
       element = React.createElement(LeaveDashboard, {
  context: this.context,
  spHttpClient: this.context.spHttpClient,
  siteUrl: this.context.pageContext.web.absoluteUrl,
  onViewChange: (view: ViewType) => {
    this.currentView = view;
    this.renderComponent();
  }
});
        break;

      case ViewType.apply:
        if (this.selectedItem) {
          // Manager clicked a request → view details
          element = React.createElement(ApplyLeave, {
            context: this.context,
            item: this.selectedItem,
            viewType: this.currentView,
            onBack: () => {
              this.selectedItem = null;
              this.currentView = ViewType.myApproval;
              this.renderComponent();
            },
            spHttpClient: this.context.spHttpClient,
            siteUrl: this.context.pageContext.web.absoluteUrl,
            onViewChange: (view: ViewType) => {
              this.currentView = view;
              this.renderComponent();
            },
            
          });
        } else {
          // Normal user applying for leave
          element = React.createElement(Leaves, {
            context: this.context,
            spHttpClient:this.context.spHttpClient,
            siteUrl:this.context.pageContext.web.absoluteUrl,
            onViewChange: (view: ViewType) => {
              this.currentView = view;
              this.renderComponent();
            }
          });
        }
        break;

      case ViewType.myLeaves:
      case ViewType.myApproval:
        element = React.createElement(LeaveRequests, {
          context: this.context,
          spHttpClient: this.context.spHttpClient,
          siteUrl: this.context.pageContext.web.absoluteUrl,
          onViewChange: (view: ViewType) => {
            this.currentView = view;
            this.renderComponent();
          },
          onSelectItem: (item: any) => {
            this.selectedItem = item;
            this.currentView = ViewType.apply;
            this.renderComponent();
          },
          viewType: this.currentView
        });
        break;
    }

    if (element) {
      ReactDom.render(element, this.domElement);
    }
  }

  public render(): void {
    this.renderComponent();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}