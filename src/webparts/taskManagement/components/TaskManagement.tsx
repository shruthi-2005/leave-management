import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import FormDetails, { ITaskItem } from './FormDetails';
import TaskForm from './TaskForm';
import TaskList from './TaskList';
import './customStyles.css';
import { LIST_NAMES, getSiteUrl } from '../../../constants';

interface ITaskManagementProps {
  context: WebPartContext;
  taskId?: string | null; // 🆕 Added
}

const TaskManagement: React.FC<ITaskManagementProps> = ({ context, taskId }) => {
  const [view, setView] = useState<'list' | 'details' | 'form'>('list');
  const [selectedTask, setSelectedTask] = useState<ITaskItem | undefined>(undefined);
  const [activeTab, setActiveTab] = useState<'MyRequests' | 'Pending' | 'Completed' | 'Overdue'>('MyRequests');
  const previousTabRef = useRef(activeTab);

  // 🔹 Add New Task → open TaskForm
  const handleAddNew = () => {
    const newTask: ITaskItem = {
      Id: 0,
      Title: '',
      TaskName: '',
      TaskDescription: '',
      AssignedBy: context.pageContext.user.email || '',
      AssignedByDisplayName: context.pageContext.user.displayName || '',
      AssignedTo: '',
      AssignedToDisplayName: '',
      NotifyUsers: [],
      NotifyUsersDisplayName: [],
      DueDate: '',
      Status: 'Open',
      Priority: 'Low',
      Comments: '',
      CommentsHistory: '',
      LastUpdatedBy: '',
      LastUpdatedDate: ''
    };
    setSelectedTask(newTask);
    setView('form');
  };

  // 🔹 Edit Task → open details page
  const handleEditTask = (task: ITaskItem) => {
    previousTabRef.current = activeTab;
    setSelectedTask(task);
    setView('details');

    // update URL param for deep link
    if (window.history.replaceState) {
      const url = new URL(window.location.href);
      url.searchParams.set('taskId', task.Id.toString());
      window.history.replaceState({}, '', url.toString());
    }
  };

  // 🔹 Back to list/home
  const handleBack = () => {
    setView('list');
    setSelectedTask(undefined);
    setActiveTab(previousTabRef.current);

    if (window.history.replaceState) {
      const url = new URL(window.location.href);
      url.searchParams.delete('taskId');
      window.history.replaceState({}, '', url.toString());
    }
  };

  const handleRefresh = () => {};

  // 🧭 🆕 When taskId prop exists (from deeplink), open directly
  useEffect(() => {
    if (taskId) {
      console.log('🧭 Opening deep-linked task directly for ID:', taskId);
      fetchAndOpenTask(taskId);
    }
  }, [taskId]);

  // 🔹 Fetch & open the selected task
  const fetchAndOpenTask = async (id: string) => {
    try {
      const siteUrl = getSiteUrl(context);
      const response: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${LIST_NAMES.Tasks}')/items(${id})?` +
          `$select=Id,Title,TaskName,TaskDescription,AssignedBy/Id,AssignedBy/Title,AssignedBy/EMail,` +
          `AssignedTo/Id,AssignedTo/Title,AssignedTo/EMail,NotifyUsers/Id,NotifyUsers/Title,NotifyUsers/EMail,` +
          `DueDate,Status,Comments,CommentsHistory,Priority,LastUpdatedBy/Id,LastUpdatedBy/Title,LastUpdatedDate&` +
          `$expand=AssignedBy,AssignedTo,NotifyUsers,LastUpdatedBy`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: 'application/json;odata=verbose' } }
      );

      const item = (await response.json()).d;
      if (!item) return;

      const task: ITaskItem = {
        Id: item.Id,
        Title: item.Title,
        TaskName: item.TaskName,
        TaskDescription: item.TaskDescription,
        AssignedBy: item.AssignedBy?.EMail || '',
        AssignedByDisplayName: item.AssignedBy?.Title || '',
        AssignedTo: item.AssignedTo?.EMail || '',
        AssignedToDisplayName: item.AssignedTo?.Title || '',
        NotifyUsers: item.NotifyUsers
          ? item.NotifyUsers.map((u: any) => ({ Id: u.Id, Title: u.Title, Email: u.EMail }))
          : [],
        NotifyUsersDisplayName: item.NotifyUsers ? item.NotifyUsers.map((u: any) => u.Title) : [],
        DueDate: item.DueDate,
        Status: item.Status,
        Comments: item.Comments,
        CommentsHistory: item.CommentsHistory || [],
        Priority: item.Priority,
        LastUpdatedBy: item.LastUpdatedBy?.EMail || '',
        LastUpdatedDate: item.LastUpdatedDate || ''
      };

      setSelectedTask(task);
      setView('details');
    } catch (err) {
      console.error('❌ Error fetching deep-linked task:', err);
    }
  };

  // 🔹 UI
  return (
    <div>
      {view === 'list' && (
        <TaskList
          context={context}
          activeTab={activeTab}
          onTabChange={(tab) => setActiveTab(tab)}
          onAddNew={handleAddNew}
          onEditTask={handleEditTask}
          onSave={() => {}}
          onCancel={() => {}}
        />
      )}

      {view === 'details' && selectedTask && (
        <FormDetails
          context={context}
          task={selectedTask}
          onBack={handleBack}
          onRefresh={handleRefresh}
          onSave={() => {
            handleBack();
            handleRefresh();
          }}
          onCancel={handleBack}
          onTabChange={(tab) => {
            console.log('🔁 Tab change requested:', tab);
            setActiveTab(tab);
            setView('list');
          }}
        />
      )}

      {view === 'form' && selectedTask && (
        <TaskForm
          context={context}
          task={selectedTask}
          onSave={() => {
            handleBack();
            handleRefresh();
          }}
          onCancel={handleBack}
        />
      )}
    </div>
  );
};

export default TaskManagement;