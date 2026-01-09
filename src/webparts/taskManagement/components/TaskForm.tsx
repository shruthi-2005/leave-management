// ✅ TaskForm.tsx
import * as React from 'react';
import { useState, useEffect } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { PeoplePicker, PrincipalType, IPeoplePickerContext } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import 'bootstrap/dist/css/bootstrap.min.css';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton } from '@fluentui/react/lib/Button';
export interface ITaskItem {
  Title: string;
  Id: number;
  TaskName: string;
  TaskDescription: string;
  AssignedBy: string;
  AssignedByDisplayName?: string;
  AssignedTo: string;
  AssignedToDisplayName?: string;
  DueDate: string;
  Status: string;
  Comments?: string;
  CommentsHistory?: string;
  Priority?: string;
  LastUpdatedBy?: string;
  LastUpdatedDate?: string;
  NotifyUsers?: (string | { Id?: number; Title?: string; Email?: string; EMail?: string })[];
  NotifyUsersDisplayName?: string[];
}

interface IUser {
  id?: number;
  text: string;
  email: string;
}

interface ITaskFormProps {
  context: WebPartContext;
  task?: ITaskItem;
  onSave: () => void;
  onCancel: () => void;
}

// ✅ Hardcoded subsite (where the Tasks list exists)
const SITE_URL = "https://elevix.sharepoint.com/sites/Trainingportal";

const TaskForm: React.FC<ITaskFormProps> = ({ context, task, onSave, onCancel }) => {
  
  const [taskName, setTaskName] = useState(task?.TaskName || '');
  const [taskDescription, setTaskDescription] = useState(task?.TaskDescription || '');
  const [assignedTo, setAssignedTo] = useState<string>(task?.AssignedTo || '');
  const [dueDate, setDueDate] = useState(task?.DueDate || '');
  const [status, setStatus] = useState(task?.Status || 'Open');
  const [priority, setPriority] = useState(task?.Priority || 'Medium');
  const [comments, setComments] = useState(task?.Comments || '');
  const [notifyUsers, setNotifyUsers] = useState<IUser[]>([]);
  const [currentUserId, setCurrentUserId] = useState<number | null>(null);
  const [assignedByName, setAssignedByName] = useState('');
  const [saving, setSaving] = useState(false);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);

  const [taskNameError, setTaskNameError] = useState('');
  const [taskDescError, setTaskDescError] = useState('');
  const [assignedToError, setAssignedToError] = useState('');
  const [dueDateError, setDueDateError] = useState('');
  const [notifyUsersError, setNotifyUsersError] = useState('');
    const [priorityError, setPriorityError] = useState('');
    const [commentsError, setCommentsError] = useState('');
    const [statusError,setStatusError] = useState('');


  const [dialogVisible, setDialogVisible] = useState(false);
  const [dialogTitle, setDialogTitle] = useState('');
  const [dialogMessage, setDialogMessage] = useState('');

  




  const peoplePickerContext: IPeoplePickerContext = {
    spHttpClient: context.spHttpClient,
    msGraphClientFactory: context.msGraphClientFactory,
    absoluteUrl: SITE_URL,
  };

  // ✅ Fetch current user
  useEffect(() => {
    const fetchCurrentUser = async () => {
      try {
        const res = await context.spHttpClient.get(
          `${SITE_URL}/_api/web/currentuser`,
          SPHttpClient.configurations.v1
        );
        const data = await res.json();
        setCurrentUserId(data.Id);
        setAssignedByName(data.Title);
      } catch (error) {
        console.error("⚠️ Failed to fetch current user:", error);
        setCurrentUserId(null);
      }
    };
    fetchCurrentUser();
  }, []);

  const getUserIdFromEmail = async (email: string): Promise<number | null> => {
    try {
      const res = await context.spHttpClient.get(
        `${SITE_URL}/_api/web/siteusers/getbyemail('${email}')`,
        SPHttpClient.configurations.v1
      );
      const data = await res.json();
      return data.Id || null;
    } catch {
      return null;
    }
  };

  const getRequestDigest = async (): Promise<string> => {
    const res = await fetch(`${SITE_URL}/_api/contextinfo`, {
      method: "POST",
      headers: { Accept: "application/json;odata=verbose" }
    });
    const data = await res.json();
    return data.d.GetContextWebInformation.FormDigestValue;
  };

  const sendEmail = async (to: string, subject: string, body: string) => {
    try {
      const digest = await getRequestDigest();
      await fetch(
        `${SITE_URL}/_api/SP.Utilities.Utility.SendEmail`,
        {
          method: "POST",
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": digest
          },
          body: JSON.stringify({
            properties: {
              __metadata: { type: "SP.Utilities.EmailProperties" },
              To: { results: [to] },
              Subject: subject,
              Body: body
            }
          })
        }
      );
    } catch (err) {
      console.error("❌ Email error:", err);
    }
  };

  const validateFields = (): boolean => {
    let valid = true;
    
    setTaskNameError('');
    setTaskDescError('');
    setAssignedToError('');
    setDueDateError('');
    setNotifyUsersError('');
    setCommentsError('');
    setPriorityError('');
    setStatusError('');

const priorityOptions =['Low','Medium','High'];
const statusOptions = ['Open'];
    
    if (!taskName.trim()) { setTaskNameError('Task Name is required'); valid = false; }
   // if (!taskDescription.trim()) { setTaskDescError('Description is required'); valid = false; }
    if (!assignedTo) { setAssignedToError('Please assign to a user'); valid = false; }
    if (!notifyUsers || notifyUsers.length ===0) { setNotifyUsersError('Please assign notifyusers'); valid = false; }
   
    if (!dueDate) {
      setDueDateError('Due Date is required'); valid = false;
    } else {
      const selectedDate = new Date(dueDate);
      const today = new Date(); today.setHours(0, 0, 0, 0);
      if (selectedDate < today) { setDueDateError('Due Date cannot be in the past'); valid = false; }
    }
        if (!comments) { setCommentsError('Comments is required'); valid = false; }
// Priority validation
    if (!priority) { 
        setPriorityError('Priority is required'); 
        valid = false; 
    } else if (!priorityOptions.includes(priority)) { 
        setPriorityError('Invalid priority selected'); 
        valid = false; 
    }

    // Status validation (always Open)
    if (!status) { 
        setStatusError('Status is required'); 
        valid = false; 
    } else if (!statusOptions.includes(status)) { 
        setStatusError('Status must be Open'); 
        valid = false; 
    }
    return valid;
  };

  const removeFile = (index: number) => {
    setSelectedFiles(prev => prev.filter((_, i) => i !== index));
  };

  const handleSave = async () => {
    if (!validateFields()) return;
    if (!currentUserId) { 
      setDialogTitle("Warning");
      setDialogMessage("User info not loaded.");
      setDialogVisible(true);
      return; 
    }

    setSaving(true);
    try {
      const assignedToId = await getUserIdFromEmail(assignedTo);
      if (!assignedToId) { 
        setDialogTitle("Warning");
        setDialogMessage("Assigned To user not found.");
        setDialogVisible(true);
        setSaving(false); 
        return; 
      }

      const notifyUsersIds: number[] = [];
      for (const u of notifyUsers) {
        if (typeof u.id === "number") {
          notifyUsersIds.push(u.id);
        } else {
          const userId = await getUserIdFromEmail(u.email);
          if (userId) notifyUsersIds.push(userId);
        }
      }

      const updatedDate = new Date().toISOString();

      const item: any = {
        __metadata: { type: "SP.Data.TasksListItem" },
        
        TaskName: taskName,
        TaskDescription: taskDescription,
        AssignedById: currentUserId,
        AssignedToId: assignedToId,
        DueDate: new Date(dueDate).toISOString(),
        Status: status,
        Priority: priority,
        Comments: comments,
        LastUpdatedById: currentUserId,
        LastUpdatedDate: updatedDate,
        CommentsHistory: task?.CommentsHistory
          ? `${task.CommentsHistory}\n[${new Date().toLocaleString()} - ${assignedByName}]: ${comments}`
          : `[${new Date().toLocaleString()} - ${assignedByName}]: ${comments}`
      };

      if (notifyUsersIds.length > 0) {
        item.NotifyUsersId = { results: notifyUsersIds };
      }

      const url = task?.Id
        ? `${SITE_URL}/_api/web/lists/getbytitle('Tasks')/items(${task.Id})`
        : `${SITE_URL}/_api/web/lists/getbytitle('Tasks')/items`;

      const headers: any = {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        "odata-version": ""
      };

      if (task?.Id) { headers['IF-MATCH'] = '*'; headers['X-HTTP-Method'] = 'MERGE'; }

      const response = await context.spHttpClient.post(url, SPHttpClient.configurations.v1, { headers, body: JSON.stringify(item) });
      let taskId = task?.Id;

      if (!taskId && response.ok) {
        const resData = await response.json();
        taskId = resData.d.Id;
      }

      if (response.ok && taskId) {
        // ✅ Upload attachments
        for (const file of selectedFiles) {
          const arrayBuffer = await file.arrayBuffer();
          await context.spHttpClient.post(
            `${SITE_URL}/_api/web/lists/getByTitle('Tasks')/items(${taskId})/AttachmentFiles/add(FileName='${file.name}')`,
            SPHttpClient.configurations.v1,
            { headers: { "accept": "application/json;odata=verbose" }, body: arrayBuffer }
          );
        }

        setDialogTitle("Success");
        setDialogMessage("Task saved successfully!");
        setDialogVisible(true);

        //  Dynamic Teams deeplink to open Task details page
// 🔹 Dynamic link to open task detail in Teams (or fallback browser)
const taskLink = (() => {
  const contextObj = {
    subEntityId: `TASK_${taskId}`,
    webUrl: `${SITE_URL}/SitePages/TaskManagement.aspx?taskId=${taskId}`
  };

  // Encode context for Teams link
  const contextParam = encodeURIComponent(JSON.stringify(contextObj));

  return `https://teams.microsoft.com/l/app/c179a8fd-4c6d-4db6-82f2-40a8672087c0?context=${contextParam}`;
})();

        //  Email to Assigned To
        await sendEmail(
          assignedTo,
          task?.Id ? `Task Updated – ${taskName}` : `New Task Assigned – ${taskName}`,
          `
          Hi ${assignedTo},<br/><br/>
          ${task?.Id ? 'Task updated' : 'You have been assigned a new task'}:<br/>
          <b>Task:</b> ${taskName}<br/>
          <b>Description:</b> ${taskDescription}<br/>
          <b>Priority:</b> ${priority}<br/>
          <b>Due Date:</b> ${dueDate}<br/>
          <b>Assigned By:</b> ${assignedByName}<br/><br/>
          👉 <a href="${taskLink}">Open in Teams microsoft</a><br/><br/>
          Thanks,<br/>Task Management System
          `
        );

        // ✅ Email to Notify Users
        for (const user of notifyUsers) {
          await sendEmail(
            user.email,
            `Notification: Task "${taskName}"`,
            `
            Hi ${user.text},<br/><br/>
            You are notified about the task:<br/>
            <b>Task:</b> ${taskName}<br/>
            <b>Description:</b> ${taskDescription}<br/>
            <b>Priority:</b> ${priority}<br/>
            <b>Due Date:</b> ${dueDate}<br/>
            <b>Assigned By:</b> ${assignedByName}<br/><br/>
            👉 <a href="${taskLink}">Open in Teams microsoft</a><br/><br/>
            Thanks,<br/>Task Management System
            `
          );
        }

      } else {
        console.error("❌ Error saving task:", await response.text());
        setDialogTitle("Error");
        setDialogMessage("Error saving task");
        setDialogVisible(true);
      }
    } catch (error) {
      console.error("🚨 Error:", error);
      setDialogTitle("Error");
      setDialogMessage("Error saving task: " + error);
      setDialogVisible(true);
    }
    setSaving(false);
  };

  const todayStr = new Date().toISOString().split('T')[0];
  if (currentUserId === null) return <div>Loading user info...</div>;

  
  return (
  <>
    <div className="container-fluid mt-3">
      <div
        className="card shadow-sm p-3"
        style={{ backgroundColor: "#f0f8ff", borderRadius: "12px" }}
      >
        

        <div className="row">
          <div className="col-12">
            <label className="form-label mt-2">Task Name <span className="text-danger">*</span></label>
            <input
              className="form-control"
              value={taskName}
              onChange={(e) => {
                setTaskName(e.target.value);
                if (e.target.value) setTaskNameError("");
              }}
            />
            {taskNameError && (
              <div className="text-danger">{taskNameError}</div>
            )}
          </div>
        </div>

        <div className="row">
          <div className="col-12">
            <label className="form-label mt-2">Description </label>
            <textarea
              className="form-control"
              value={taskDescription}
              onChange={(e) => {
                setTaskDescription(e.target.value);
                if (e.target.value) setTaskDescError("");
              }}
            />
            {taskDescError && (
              <div className="text-danger">{taskDescError}</div>
            )}
          </div>
        </div>

        <div className="row">
          <div className="col-12">
            <label className="form-label mt-2">Assign To <span className="text-danger">*</span></label>
            <div className="form-control p-0">
            <PeoplePicker
              context={peoplePickerContext}
              titleText=""
              personSelectionLimit={1}
              showtooltip={true}
              onChange={(items: any[]) => {
                setAssignedTo(
                  items.length > 0 ? items[0].mail || items[0].secondaryText : ""
                );
                if (items.length > 0) setAssignedToError("");
              }}
              principalTypes={[PrincipalType.User]}
              defaultSelectedUsers={assignedTo ? [assignedTo] : []}
            />
            </div>
            {assignedToError && (
              <div className="text-danger">{assignedToError}</div>
            )}
          </div>
        </div>

        <div className="row">
          <div className="col-12">
            <label className="form-label mt-2">Notify Users <span className="text-danger">*</span></label>
            <div className="form-control p-0">
            <PeoplePicker
              context={peoplePickerContext}
              titleText=""
              personSelectionLimit={10}
              showtooltip={true}
              onChange={(items: any[]) => {
                const users = items.map((i) => ({
                  id: i.id,
                  text: i.text,
                  email: i.mail || i.secondaryText,
                }));
                setNotifyUsers(users);
                setNotifyUsersError("");
              }}
              principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup]}
              defaultSelectedUsers={notifyUsers.map((u) => u.email)}
            />
            </div>
            {notifyUsersError && (
              <div className="text-danger">{notifyUsersError}</div>
            )}
          </div>
        </div>

        <div className="row">
          <div className="col-12">
            <label className="form-label mt-2">Due Date <span className="text-danger">*</span></label>
            <input
              type="date"
              className="form-control"
              value={dueDate}
              min={todayStr}
              onChange={(e) => {
                setDueDate(e.target.value);
                if (e.target.value) setDueDateError("");
              }}
            />
            {dueDateError && <div className="text-danger">{dueDateError}</div>}
          </div>
        </div>

        <div className="row">
          <div className="col-12">
            <label className="form-label mt-2">Status</label>
            <select
              className="form-select"
              value={status}
              onChange={(e) => setStatus(e.target.value)}
            >
              <option value="Open">Open</option>
              <option value="In Progress">In Progress</option>
              <option value="Completed">Completed</option>
            </select>
            {statusError&& <div className="text-danger">{statusError}</div>}            
          </div>
        </div>

        <div className="row">
          <div className="col-12">
            <label className="form-label mt-2">Priority <span className="text-danger">*</span></label>
            <select
              className="form-select"
              value={priority}
              onChange={(e) => setPriority(e.target.value)}
            >
              <option value="Low">Low</option>
              <option value="Medium">Medium</option>
              <option value="High">High</option>
            </select>
            {priorityError&& <div className="text-danger">{priorityError}</div>}
          </div>
        </div>

        <div className="row">
          <div className="col-12">
            <label className="form-label mt-2">Comments <span className="text-danger">*</span></label>
            <textarea
              className="form-control"
              value={comments}
              onChange={(e) => setComments(e.target.value)}
            />
            {commentsError&& <div className="text-danger">{commentsError}</div>}
          </div>
        </div>

        <div className="row">
          <div className="col-12">
            <label className="form-label mt-2">Attachments</label>
            <input
              type="file"
              className="form-control"
              multiple
              onChange={(e) => {
                if (e.target.files) setSelectedFiles(Array.from(e.target.files));
              }}
            />
            <ul className="list-group mt-2">
              {selectedFiles.map((file, idx) => (
                <li
                  key={idx}
                  className="list-group-item d-flex justify-content-between align-items-center flex-wrap"
                >
                  <span
                    className="text-truncate"
                    style={{ maxWidth: "70%" }}
                  >
                    {file.name}
                  </span>
                  <button
                    className="btn btn-sm btn-danger mt-1"
                    onClick={() => removeFile(idx)}
                  >
                    Remove
                  </button>
                </li>
              ))}
            </ul>
          </div>
        </div>

        <div className="row mt-3">
          <div className="col-12 d-flex flex-column flex-md-row justify-content-end">
            <button
              className="btn btn-secondary me-md-2 mb-2 mb-md-0"
              onClick={onCancel}
              disabled={saving}
            >
              Cancel
            </button>
            <button
              className="btn btn-primary"
              onClick={handleSave}
              disabled={saving}
            >
              {saving ? "Saving..." : "Save"}
            </button>
          </div>
        </div>
      </div>
    </div>

    {/* ✅ Dialog */}
    <Dialog
      hidden={!dialogVisible}
      onDismiss={() => {}}
      dialogContentProps={{
        type: DialogType.normal,
        title: dialogTitle,
        subText: dialogMessage,
      }}
    >
      <DialogFooter>
        <PrimaryButton
          onClick={() => {
            setDialogVisible(false);
            if (dialogTitle === "Success") {
              onSave();
            }
          }}
          text="OK"
        />
      </DialogFooter>
    </Dialog>
  </>
);
};

export default TaskForm;  