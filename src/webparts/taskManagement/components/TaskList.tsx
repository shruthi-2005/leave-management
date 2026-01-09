import * as React from 'react';
import { useEffect, useState, useRef } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import 'bootstrap/dist/css/bootstrap.min.css';
import FormDetails, { ITaskItem } from './FormDetails';
import { LIST_NAMES, getSiteUrl } from '../../../constants';

interface ITaskListProps {
  context: WebPartContext;
  onAddNew: () => void;
  onEditTask: (task: ITaskItem) => void;
  onSave: () => void;       // parent save if needed
  onCancel: () => void;
  activeTab: 'MyRequests' | 'Pending' | 'Completed' | 'Overdue';
  onTabChange: (tab: 'MyRequests' | 'Pending' | 'Completed' | 'Overdue') => void;
}

const TaskList: React.FC<ITaskListProps> = ({
  context,
  onAddNew,
  onEditTask,
  onSave,
  onCancel,
  activeTab,
  onTabChange
}) => {
  const [tasks, setTasks] = useState<ITaskItem[]>([]);
  const [selectedTask, setSelectedTask] = useState<ITaskItem | null>(null);
  const [refreshFlag, setRefreshFlag] = useState(false);
  const previousTabRef = useRef(activeTab);

  const currentUserEmail = context.pageContext.user.email.toLowerCase();

  const loadTasks = async () => {
    try {
      const today = new Date().toISOString().split('T')[0];
      let filter = '';
      switch (activeTab) {
        case 'MyRequests':
          filter = `AssignedBy/EMail eq '${currentUserEmail}'`;
          break;
        case 'Pending':
          filter = `AssignedTo/EMail eq '${currentUserEmail}' and (Status eq 'Open' or Status eq 'In Progress')`;
          break;
        case 'Completed':
          filter = `AssignedTo/EMail eq '${currentUserEmail}' and Status eq 'Completed'`;
          break;
        case 'Overdue':
          filter = `AssignedTo/EMail eq '${currentUserEmail}' and Status ne 'Completed' and DueDate lt datetime'${today}T00:00:00Z'`;
          break;
      }

      const siteUrl = getSiteUrl(context);
      let url =
        `${siteUrl}/_api/web/lists/getByTitle('${LIST_NAMES.Tasks}')/items?` +
        `$select=Id,Title,TaskName,TaskDescription,AssignedBy/Title,AssignedBy/EMail,` +
        `AssignedTo/Title,AssignedTo/EMail,DueDate,Status,Comments,Priority,NotifyUsers/Title,NotifyUsers/EMail,` +
        `CommentsHistory,LastUpdatedBy/Title,LastUpdatedBy/EMail,LastUpdatedDate&` +
        `$expand=AssignedBy,AssignedTo,NotifyUsers,LastUpdatedBy&` +
        `${filter ? `$filter=${encodeURIComponent(filter)}&` : ''}$orderby=Id desc&$top=5000`;

      let allItems: any[] = [];
      while (url) {
        const response: SPHttpClientResponse = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        const data = await response.json().catch(() => null);
        if (!data || !data.value) break;

        allItems = allItems.concat(data.value);
        url = data['@odata.nextLink'] || null;
      }

      const formattedTasks: ITaskItem[] = allItems.map((t: any) => ({
        Id: t.Id,
        Title: t.Title,
        TaskName: t.TaskName,
        TaskDescription: t.TaskDescription,
        AssignedBy: t.AssignedBy?.EMail?.toLowerCase() || '',
        AssignedByDisplayName: t.AssignedBy?.Title || '',
        AssignedTo: t.AssignedTo?.EMail?.toLowerCase() || '',
        AssignedToDisplayName: t.AssignedTo?.Title || '',
        DueDate: t.DueDate,
        Status: t.Status,
        Comments: t.Comments || '',
        Priority: t.Priority || '',
        NotifyUsers: Array.isArray(t.NotifyUsers) ? t.NotifyUsers : [],
        NotifyUsersDisplayName: Array.isArray(t.NotifyUsers) ? t.NotifyUsers.map((u: any) => u.Title) : [],
        CommentsHistory: t.CommentsHistory || '',
        LastUpdatedBy: t.LastUpdatedBy?.EMail?.toLowerCase() || '',
        LastUpdatedByDisplayName: t.LastUpdatedBy?.Title || '',
        LastUpdatedDate: t.LastUpdatedDate || ''
      }));

      setTasks(formattedTasks);
    } catch (error) {
      console.error('❌ Error fetching tasks:', error);
    }
  };

  useEffect(() => {
    loadTasks();
  }, [activeTab, refreshFlag]);

  // ✅ Fixed navigation block
  if (selectedTask) {
    return (
      <FormDetails
        context={context}
        task={selectedTask}
        onBack={() => {
          setSelectedTask(null);
          onTabChange(previousTabRef.current);
        }}
        onRefresh={() => setRefreshFlag(!refreshFlag)}
        onSave={(updatedTask: ITaskItem) => {
          setSelectedTask(null);

          if (updatedTask.Status === 'In Progress') {
            onTabChange('Pending');
          } else if (updatedTask.Status === 'Completed') {
            onTabChange('Completed');
          } else {
            onTabChange(previousTabRef.current);
          }

          onSave();
        }}
        onCancel={onCancel}
        onTabChange={(tab) => {
          console.log("🔁 Tab change requested from FormDetails:", tab);
          onTabChange(tab);  // direct parent handler
          setSelectedTask(null);  // back to list view
        }}
      />
    );
  }

  return (
    <div className="container mt-3">
      <div className="d-flex flex-wrap justify-content-between align-items-center mb-3">
        <h4 className="mb-2">Tasks</h4>
        <button className="btn btn-primary mb-2" onClick={onAddNew}>
          + Add New Task
        </button>
      </div>

      <ul className="nav nav-tabs flex-wrap mb-3">
        {['MyRequests', 'Pending', 'Completed', 'Overdue'].map((tab) => (
          <li className="nav-item" key={tab}>
            <button
              className={`nav-link ${activeTab === tab ? 'active' : ''}`}
              onClick={() => onTabChange(tab as any)}
            >
              {tab === 'MyRequests'
                ? 'My Requests'
                : tab === 'Pending'
                ? 'My Pending Tasks'
                : tab === 'Completed'
                ? 'My Completed Tasks'
                : 'Overdue Tasks'}
            </button>
          </li>
        ))}
      </ul>

      <div className="table-responsive">
        <table className="table table-striped table-hover align-middle">
          <thead className="table-light">
            <tr>
              <th>Task Name</th>
              <th>Assigned By</th>
              <th>Assigned To</th>
              <th>Status</th>
              <th>Priority</th>
            </tr>
          </thead>
          <tbody>
            {tasks.length === 0 && (
              <tr>
                <td colSpan={5} className="text-center">
                  No tasks found.
                </td>
              </tr>
            )}
            {tasks.map((task) => (
              <tr
                key={task.Id}
                onClick={() => {
                  previousTabRef.current = activeTab;
                  setSelectedTask(task);
                  onEditTask(task);
                }}
                style={{ cursor: 'pointer' }}
              >
                <td>{task.TaskName || task.Title}</td>
                <td>{task.AssignedByDisplayName || task.AssignedBy}</td>
                <td>{task.AssignedToDisplayName || task.AssignedTo}</td>
                <td>
                  <span
                    className={`badge ${
                      task.Status === 'Completed'
                        ? 'bg-success'
                        : task.Status === 'In Progress'
                        ? 'bg-warning text-dark'
                        : 'bg-secondary'
                    }`}
                  >
                    {task.Status}
                  </span>
                </td>
                <td>
                  <span
                    className={`badge ${
                      task.Priority === 'High'
                        ? 'bg-danger'
                        : task.Priority === 'Medium'
                        ? 'bg-info text-dark'
                        : 'bg-light text-dark'
                    }`}
                  >
                    {task.Priority || '-'}
                  </span>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default TaskList;