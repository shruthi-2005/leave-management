// FormDetails.tsx
import * as React from 'react';
import { useState, useEffect } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { PeoplePicker, PrincipalType, IPeoplePickerContext } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { LIST_NAMES } from '../../../constants';
import { __makeTemplateObject, __metadata } from 'tslib';
import * as microsoftTeams from "@microsoft/teams-js"

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
  NotifyUsers?: (string | { Id: number; Title: string; Email: string; EMail?: string })[];
  NotifyUsersDisplayName?: string[];
}

interface IPeoplePickerUser {
  Email?: string;
  id?: string | number;
  imageUrl?: string;
  loginName?: string;      // Use this for SPUser login
  secondaryText?: string;  // Usually the email
  text?: string;           // Display name
}

interface IFormDetailsProps {
  context: WebPartContext;
  task: ITaskItem;
  fromTab?: 'MyRequests' | 'Pending' | 'Completed' | 'Overdue'; // new
  onBack: (tabToGoBack?: 'MyRequests' | 'Pending' | 'Completed' | 'Overdue') => void;
  onRefresh: () => void;
  onSave: (updatedTask:ITaskItem) => void;
  onCancel: () => void;
  onTabChange:(tab: 'MyRequests' | 'Pending' | 'Completed' | 'Overdue')=>void
}

// === HARD-CODED SITE URL FOR TEAMS / IFRAME ===
const siteUrl = 'https://elevix.sharepoint.com/sites/Trainingportal';
const listName = LIST_NAMES?.Tasks || 'Tasks';

const FormDetails: React.FC<IFormDetailsProps> = ({ context, task, onBack, onRefresh, onCancel,onSave,onTabChange }) => {
  const [status, setStatus] = useState<string>(task?.Status || '');
  const [comments, setComments] = useState<string>('');
  const [currentUserEmail, setCurrentUserEmail] = useState<string>('');
  const [attachments, setAttachments] = useState<{ FileName: string; ServerRelativeUrl: string }[]>([]);
  const [uploadingFiles, setUploadingFiles] = useState<File[]>([]);
  const [showDialog, setShowDialog] = useState<boolean>(false);
  const [dialogMessage, setDialogMessage] = useState<string>('');

  const [editableTaskName, setEditableTaskName] = useState<string>(task?.TaskName || '');
  const [editableTaskDescription, setEditableTaskDescription] = useState<string>(task?.TaskDescription || '');
  const [editableDueDate, setEditableDueDate] = useState<string>(task?.DueDate ? task.DueDate.split('T')[0] : '');
  const [editablePriority, setEditablePriority] = useState<string>(task?.Priority || '');
  const [assignedToSelected, setAssignedToSelected] = useState<IPeoplePickerUser[] | undefined>(undefined);
  const [notifyUsersSelected, setNotifyUsersSelected] = useState<IPeoplePickerUser[] | undefined>(undefined);
//const [savedStatus, setSavedStatus]=useState<string>('')
//const [saveSuccess, setSaveSuccess] = useState(false);

useEffect(() => {
  const initTeams = async () => {
    try {
      await microsoftTeams.app.initialize();
      const context = await microsoftTeams.app.getContext();

      // ✅ Updated for latest SDK version
      const subEntityId =
        (context as any).subEntityId || context.page?.subPageId;

      console.log("📩 Teams Deep Link ID:", subEntityId);

      if (subEntityId) {
        window.location.href = `${siteUrl}/SitePages/TaskManagement.aspx?taskId=${subEntityId}`;
      }
    } catch (error) {
      console.error("❌ Teams initialization error:", error);
    }
  };

  initTeams();
}, []);


  const peoplePickerContext: IPeoplePickerContext = {
    spHttpClient: context.spHttpClient,
    msGraphClientFactory: (context as any).msGraphClientFactory,
    absoluteUrl: siteUrl,
  };

  useEffect(() => {
    const email = (context?.pageContext?.user?.email || '').toLowerCase();
    setCurrentUserEmail(email);
  }, [context]);

  useEffect(() => {
    setEditableTaskName(task?.TaskName || '');
    setEditableTaskDescription(task?.TaskDescription || '');
    setEditableDueDate(task?.DueDate ? task.DueDate.split('T')[0] : '');
    setEditablePriority(task?.Priority || '');
    setAssignedToSelected(undefined);
    setNotifyUsersSelected(undefined);
    setStatus(task?.Status || '');
    if (task?.Id) {
      loadAttachments().catch(err => console.error('Attachment load error:', err));
    } else {
      setAttachments([]);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [task]);

  // Determine who can edit
 const isAssignedTo = currentUserEmail === (task?.AssignedTo || '').toLowerCase();
const isAssignedBy = currentUserEmail === (task?.AssignedBy || '').toLowerCase();
const isSameUser = isAssignedTo && isAssignedBy;

// Assigned To editable if Open or In Progress
const canEditAssignedTo = isAssignedTo && ['Open', 'In Progress'].includes(task?.Status || '');

// Assigned By can edit all fields if Status = Open
const canEditAssignedByFull = isAssignedBy && task?.Status === 'Open';

// Assigned By can edit Status & Comments only if Status = In Progress
const canEditAssignedByStatusOnly = isAssignedBy && ['In Progress','Completed'].includes( task?.Status ||'');

// Show Status & Comments section if AssignedBy = AssignedTo
const showStatusComments = canEditAssignedTo || canEditAssignedByStatusOnly || isSameUser;

// Show Save button if editable
const showSaveButton = canEditAssignedTo || canEditAssignedByStatusOnly || isSameUser;

  const statusOptions: string[] = canEditAssignedTo
    ? ['In Progress', 'Completed']
    : canEditAssignedByFull
      ? ['In Progress', 'Completed']
      : canEditAssignedByStatusOnly
        ? ['Approved', 'Rejected']
        : [];

  // ----------------- Helpers -----------------
  const getRequestDigest = async (): Promise<string> => {
    const response = await fetch(`${siteUrl}/_api/contextinfo`, {
      method: "POST",
      headers: { Accept: "application/json;odata=verbose" }
    });
    const data = await response.json().catch(() => null);
    return data?.d?.GetContextWebInformation?.FormDigestValue || data?.GetContextWebInformation?.FormDigestValue || '';
  };

  const getItemTypeForList = async (listTitle: string): Promise<string> => {
    const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${listTitle}')?$select=ListItemEntityTypeFullName`;
    const response = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
    if (!response.ok) throw new Error(`Failed to get list item type for ${listTitle}`);
    const data = await response.json().catch(() => null);
    return data?.ListItemEntityTypeFullName || data?.d?.ListItemEntityTypeFullName || 'SP.Data.TasksListItem';
  };

  const getUserIdFromEmail = async (email: string): Promise<number> => {
    if (!email) return 0;
    try {
      const endpoint = `${siteUrl}/_api/web/siteusers/getbyemail('${encodeURIComponent(email)}')?$select=Id`;
      const response = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
      if (!response.ok) {
        // fallback to ensureuser
        const ensure = await context.spHttpClient.post(
          `${siteUrl}/_api/web/siteusers/ensureuser`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: 'application/json;odata=verbose',
              'Content-Type': 'application/json;odata=verbose'
            },
            body: JSON.stringify({ logonName: email })
          }
        );
        if (!ensure.ok) return 0;
        const en = await ensure.json().catch(() => null);
        return en?.d?.Id || en?.Id || 0;
      }
      const data = await response.json().catch(() => null);
      return data?.d?.Id || data?.Id || 0;
    } catch (err) {
      console.error("❌ getUserIdFromEmail error:", email, err);
      return 0;
    }
  };

  // ----------------- Attachments -----------------
  const loadAttachments = async (): Promise<void> => {
    if (!task?.Id) {
      setAttachments([]);
      return;
    }
    try {
      const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${task.Id})/AttachmentFiles`;
      const res = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
      if (!res.ok) {
        console.warn('Attachments fetch returned not ok:', res.status);
        setAttachments([]);
        return;
      }
      const data = await res.json().catch(() => null);
      const arr = data?.value || data?.d?.results || [];
      if (Array.isArray(arr)) {
        setAttachments(arr.map((a: any) => ({ FileName: a.FileName, ServerRelativeUrl: a.ServerRelativeUrl })));
      } else {
        setAttachments([]);
      }
    } catch (err) {
      console.error("❌ Error loading attachments:", err);
      setAttachments([]);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) setUploadingFiles(Array.from(e.target.files));
  };

  const handleUploadAttachments = async () => {
    if (!uploadingFiles || uploadingFiles.length === 0) return;
    try {
      const digest = await getRequestDigest();
      for (const file of uploadingFiles) {
        const buffer = await file.arrayBuffer();
        const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${task.Id})/AttachmentFiles/add(FileName='${encodeURIComponent(file.name)}')`;
        await fetch(endpoint, {
          method: "POST",
          headers: {
            "Accept": "application/json;odata=verbose",
            "X-RequestDigest": digest
          },
          body: buffer
        });
      }
      await loadAttachments();
      setUploadingFiles([]);
    } catch (err) {
      console.error("❌ Upload failed:", err);
    }
  };

  const handleDeleteAttachment = async (fileName: string) => {
    try {
      const digest = await getRequestDigest();
      const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${task.Id})/AttachmentFiles/getbyfilename('${encodeURIComponent(fileName)}')`;
      await fetch(endpoint, {
        method: "DELETE",
        headers: {
          "Accept": "application/json;odata=verbose",
          "X-RequestDigest": digest,
          "IF-MATCH": "*"
        }
      });
      await loadAttachments();
    } catch (err) {
      console.error("❌ Delete failed:", err);
    }
  };

  // ----------------- Notify Emails -----------------
  const getNotifyUsersEmailsFromTask = (): string[] => {
    if (!task?.NotifyUsers || task.NotifyUsers.length === 0) return [];
    return task.NotifyUsers.map(u => {
      if (typeof u === 'string') return u;
      return (u as any).Email || (u as any).EMail || (u as any).Title || '';
    }).filter(Boolean) as string[];
  };

  const sendNotifyUsersEmails = async (updatedStatus?: string) => {
    if (!task?.NotifyUsers || task.NotifyUsers.length === 0) return;
    // ✅ Dynamic Teams deeplink to open Task details page
const taskLink = `https://teams.microsoft.com/l/app/c179a8fd-4c6d-4db6-82f2-40a8672087c0?context=${encodeURIComponent(JSON.stringify({
  subEntityId: `${task}`,
  webUrl: `${siteUrl}/SitePages/TaskManagement.aspx?taskId=${task}`
}))}`;
    for (const user of task.NotifyUsers) {
      let email: string | undefined;
      let displayName: string | undefined;
      if (typeof user === 'string') {
        email = user;
        displayName = user;
      } else {
        email = (user as any).Email || (user as any).EMail;
        displayName = (user as any).Title || (user as any).Email || (user as any).EMail;
      }
      if (!email) continue;
      try {
        const digest = await getRequestDigest();
        await fetch(`${siteUrl}/_api/SP.Utilities.Utility.SendEmail`, {
          method: "POST",
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": digest
          },
          body: JSON.stringify({
            properties: {
              __metadata: { type: "SP.Utilities.EmailProperties" },
              To: { results: [email.toLowerCase()] },
              Subject: `Notification: Task "${task.TaskName}"`,
              Body: `
                Hi ${displayName || ''},<br/><br/>
                You are notified about the task:<br/>
                <b>Task:</b> ${task.TaskName}<br/>
                <b>Description:</b> ${task.TaskDescription || ''}<br/>
                <b>Status:</b> ${updatedStatus || status || ''}<br/>
                <b>Updated On:</b> ${new Date().toLocaleString()}<br/><br/>
                👉 <a href="${taskLink}">Open in Teams</a><br/><br/>
                Thanks,<br/>Task Management System
              `
            }
          })
        });
      } catch (err) {
        console.error(`❌ Failed to send NotifyUsers email to ${email}:`, err);
      }
    }
  };

  // ----------------- Save -----------------
const handleSave = async () => {
  console.log(' handleSave triggered');
  if (!comments.trim()) {
    setDialogMessage('⚠️ Comments required!');
    setShowDialog(true);
    return;
  }

  try {
    const listItemType = await getItemTypeForList(listName);
    console.log('list item type:',listItemType)
    const updatedDate = new Date().toISOString();

    const updatedCommentsHistory = task?.CommentsHistory
      ? `${task.CommentsHistory}\n[${new Date().toLocaleString()} - ${currentUserEmail}]: ${comments}`
      : `[${new Date().toLocaleString()} - ${currentUserEmail}]: ${comments}`;

    const lastUpdatedById = await getUserIdFromEmail(currentUserEmail);

    const body: any = {
      __metadata: { type: listItemType || 'SP.Data.TasksListItem' },
      Status: status,
      Comments: comments,
      CommentsHistory: updatedCommentsHistory,
      LastUpdatedById: lastUpdatedById,
      LastUpdatedDate: updatedDate
    };

    // PeoplePicker save logic when AssignedBy can edit everything
    if (canEditAssignedByFull && task?.Status === 'Open') {
      body.TaskName = editableTaskName;
      body.TaskDescription = editableTaskDescription;
      body.DueDate = editableDueDate ? new Date(editableDueDate).toISOString() : null;
      body.Priority = editablePriority || '';

      // resolve AssignedToId
      let assignedToId = 0;
      if (assignedToSelected && assignedToSelected.length > 0) {
        const sel = assignedToSelected[0];
        if (sel && (sel as any).id && !isNaN(Number((sel as any).id))) {
          assignedToId = Number((sel as any).id);
        } else {
          const emailToResolve = sel?.secondaryText || sel?.Email || sel?.loginName || '';
          if (emailToResolve) assignedToId = await getUserIdFromEmail(emailToResolve);
        }
      } else if (task?.AssignedTo) {
        assignedToId = await getUserIdFromEmail(task.AssignedTo);
      }
      if (assignedToId && assignedToId > 0) body.AssignedToId = assignedToId;

      // resolve NotifyUsers
      const notifyUsersIds: number[] = [];
      if (notifyUsersSelected && notifyUsersSelected.length > 0) {
        for (const u of notifyUsersSelected) {
          if (u && (u as any).id && !isNaN(Number((u as any).id))) {
            notifyUsersIds.push(Number((u as any).id));
          } else {
            const emailToResolve = u?.secondaryText || u?.Email || u?.loginName || '';
            if (emailToResolve) {
              const uid = await getUserIdFromEmail(emailToResolve);
              if (uid && uid > 0) notifyUsersIds.push(uid);
            }
          }
        }
      } else if (task?.NotifyUsers && task.NotifyUsers.length > 0) {
        for (const nu of task.NotifyUsers) {
          if (typeof nu === 'object') {
            const idVal = (nu as any).Id || (nu as any).id;
            if (idVal && !isNaN(Number(idVal))) notifyUsersIds.push(Number(idVal));
            else {
              const em = (nu as any).Email || (nu as any).EMail || (nu as any).Title || '';
              if (em) {
                const uid = await getUserIdFromEmail(em.toString());
                if (uid && uid > 0) notifyUsersIds.push(uid);
              }
            }
          } else if (typeof nu === 'string') {
            const uid = await getUserIdFromEmail(nu);
            if (uid && uid > 0) notifyUsersIds.push(uid);
          }
        }
      }

      if (notifyUsersIds.length > 0) body.NotifyUsersId = { results: notifyUsersIds };
      else body.NotifyUsersId = null;
    }

    // Update item
    const updateEndpoint = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${task.Id})`;
    const updateRes = await context.spHttpClient.post(updateEndpoint, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE',
        'odata-version': ''
      },
      body: JSON.stringify(body)
    });

    if (!updateRes.ok) {
      const errTxt = await updateRes.text().catch(() => 'no text');
      console.error("❌ Update failed:", errTxt);
      setDialogMessage('❌ Update failed, check console');
      setShowDialog(true);
      return;
    }
    
    
let tabToGo: 'Pending'|'Completed'|null=null;
    // ----------------- Navigate tab based on status -----------------
    //
    if (status === 'In Progress') 
      tabToGo=('Pending');
     else if (status === 'Completed') 
      tabToGo=('Completed');
   console.log('status after save:',status)
   console.log('tabToGo:',tabToGo)

   sessionStorage.setItem('nextTab',tabToGo||'');

     console.log("✅ Task updated successfully");
    setDialogMessage("✅ Task updated successfully!");
    setShowDialog(true); // show message
  
    //console.log('onTabChange function exists?',typeof onTabChange);
// if (tabToGo && typeof onTabChange === 'function'){
 // console.log('calling onTabChange now...');
 // onTabChange(tabToGo);
// }

    
    // who to email
    const mainRecipients: string[] = [];
    if (['In Progress', 'Completed'].includes(status || '')) {
      if (task?.AssignedBy) mainRecipients.push(task.AssignedBy);
    } else if (['Approved', 'Rejected'].includes(status || '')) {
      if (task?.AssignedTo) mainRecipients.push(task.AssignedTo);
    }

    // ✅ Dynamic Teams deeplink to open Task details page
const taskLink = `https://teams.microsoft.com/l/app/c179a8fd-4c6d-4db6-82f2-40a8672087c0?context=${encodeURIComponent(JSON.stringify({
  subEntityId: `${task}`,
  webUrl: `${siteUrl}/SitePages/TaskManagement.aspx?taskId=${task}`
}))}`;
    const subject = `Task Update – ${task.TaskName}`;
    const bodyEmail = `
      Hi,<br/><br/>
      The task <b>${task.TaskName}</b> has been updated.<br/>
      <b>Status:</b> ${status}<br/>
      <b>Comments:</b> ${comments || 'No comments'}<br/>
      <b>Updated On:</b> ${new Date(updatedDate).toLocaleDateString()}<br/><br/>
      👉 <a href="${taskLink}">Open in Teams</a><br/><br/>
      Thanks,<br/>Task Management System
    `;

    if (mainRecipients.length > 0) {
      const digest = await getRequestDigest();
      try {
        await fetch(`${siteUrl}/_api/SP.Utilities.Utility.SendEmail`, {
          method: "POST",
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": digest
          },
          body: JSON.stringify({
            properties: {
              __metadata: { type: "SP.Utilities.EmailProperties" },
              To: { results: mainRecipients.map(e => e.toLowerCase()) },
              Subject: subject,
              Body: bodyEmail
            }
          })
        });
      } catch (err) {
        console.error('Email send failed:', err);
      }
    }

    // notify other users
    await sendNotifyUsersEmails(status);


  } catch (err) {
    console.error('❌ Save error:', err);
    setDialogMessage('❌ Error updating task');
    setShowDialog(true);
  }
};

  // ----------------- Navigation -----------------
  const handleBack = () => {
    try {
      const url = new URL(window.location.href);
      url.searchParams.delete('taskId');
      window.history.replaceState({}, document.title, url.toString());
    } catch (err) {
      // ignore
    }
    onBack();
  };

  const onAssignedToChange = (items?: any[]) => {
    if (items && items.length > 0) setAssignedToSelected(items as IPeoplePickerUser[]);
    else setAssignedToSelected(undefined);
  };

  const onNotifyUsersChange = (items?: any[]) => {
    if (items && items.length > 0) setNotifyUsersSelected(items as IPeoplePickerUser[]);
    else setNotifyUsersSelected(undefined);
  };

  const assignedToDefault: string[] = task?.AssignedTo ? [task.AssignedTo] : [];
  const notifyUsersDefault: string[] = getNotifyUsersEmailsFromTask();

  // ----------------- Render ----------------
  return (
    <div
      style={{
        maxWidth: '100%',
        width: '100%',
        margin: '0 auto',
        padding: '20px',
        backgroundColor: '#f0f8ff',
        borderRadius: '12px',
        boxShadow: '0 4px 12px rgba(0,0,0,0.15)'
      }}
    >
      {/* Dialog */}
      <Dialog
        hidden={!showDialog}
        onDismiss={() => setShowDialog(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Task Update',
          subText: dialogMessage
        }}
      >
        <DialogFooter>
  <PrimaryButton
  text="OK"
  onClick={() => {
    setShowDialog(false);
    const nextTab = sessionStorage.getItem('nextTab');
    if (nextTab && typeof onTabChange === 'function') {
      onTabChange(nextTab as 'Pending' | 'Completed');
      sessionStorage.removeItem('nextTab');
    }
  }}
/>
</DialogFooter>
  
      </Dialog>

      {/* Back Button */}
      <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: '10px' }}>
        <button
          style={{ padding: '5px 10px', backgroundColor: '#6c757d', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer' }}
          onClick={handleBack}
        >
          Back
        </button>
      </div>

      {/* Task Name */}
      <div style={{ marginBottom: '15px' }}>
        <b>Task Name:</b>
        {canEditAssignedByFull ? (
          <input
            type="text"
            value={editableTaskName || task.TaskName}
            onChange={(e) => setEditableTaskName(e.target.value)}
            style={{ width: '100%', padding: '8px', borderRadius: '4px', border: '1px solid #ccc', marginTop: '5px' }}
          />
        ) : (
          <div style={{ padding: '8px', borderRadius: '4px', backgroundColor: '#e0e0e0', color: '#555' }}>
            {task.TaskName}
          </div>
        )}
      </div>

      {/* Description */}
      <div style={{ marginBottom: '15px' }}>
        <b>Description:</b>
        {canEditAssignedByFull ? (
          <textarea
            value={editableTaskDescription || task.TaskDescription}
            onChange={(e) => setEditableTaskDescription(e.target.value)}
            rows={4}
            style={{ width: '100%', padding: '8px', borderRadius: '4px', border: '1px solid #ccc', marginTop: '5px' }}
          />
        ) : (
          <div style={{ padding: '8px', borderRadius: '4px', backgroundColor: '#e0e0e0', color: '#555', wordBreak: 'break-word' }}>
            {task.TaskDescription}
          </div>
        )}
      </div>

      {/* Assigned By */}
      {!canEditAssignedByFull && (
      <div style={{ marginBottom: '15px' }}>
        <b>Assigned By:</b>
        <div style={{ padding: '8px', borderRadius: '4px', backgroundColor: '#e0e0e0', color: '#555', wordBreak: 'break-word' }}>
          {task.AssignedByDisplayName || task.AssignedBy}
        </div>
      </div>
      )}

      {/* Assigned To */}
      <div style={{ marginBottom: '15px' }}>
        <b>Assigned To:</b>
        { canEditAssignedByFull ? (
          <PeoplePicker
            context={peoplePickerContext}
            personSelectionLimit={1}
            showtooltip={true}
            required={true}
            principalTypes={[PrincipalType.User]}
            defaultSelectedUsers={assignedToDefault}
            onChange={onAssignedToChange}
          />
        ) : (
          <div style={{ padding: '8px', borderRadius: '4px', backgroundColor: '#e0e0e0', color: '#555', wordBreak: 'break-word' }}>
            {task.AssignedToDisplayName || task.AssignedTo}
          </div>
        )}
      </div>

      {/* Due Date */}
      <div style={{ marginBottom: '15px' }}>
        <b>Due Date:</b>
        {canEditAssignedByFull ? (
          <input
            type="date"
            value={editableDueDate || task.DueDate || ''}
            onChange={(e) => setEditableDueDate(e.target.value)}
            style={{ width: '100%', padding: '8px', borderRadius: '4px', border: '1px solid #ccc', marginTop: '5px' }}
          />
        ) : (
          <div style={{ padding: '8px', borderRadius: '4px', backgroundColor: '#e0e0e0', color: '#555' }}>
            {task.DueDate ? new Date(task.DueDate).toLocaleDateString() : '-'}
          </div>
        )}
      </div>

      {/* Priority */}
      <div style={{ marginBottom: '15px' }}>
        <b>Priority:</b>
        {canEditAssignedByFull ? (
          <select
            value={editablePriority || task.Priority || ''}
            onChange={(e) => setEditablePriority(e.target.value)}
            style={{ width: '100%', padding: '8px', borderRadius: '4px', border: '1px solid #ccc', marginTop: '5px' }}
          >
            <option value="">Select Priority</option>
            <option value="Low">Low</option>
            <option value="Normal">Normal</option>
            <option value="High">High</option>
          </select>
        ) : (
          <div style={{ padding: '8px', borderRadius: '4px', backgroundColor: '#e0e0e0', color: '#555' }}>
            {task.Priority || '-'}
          </div>
        )}
      </div>

      {/* Notify Users */}
      <div style={{ marginBottom: '15px' }}>
        <b>Notify Users:</b>
        {canEditAssignedByFull ? (
          <PeoplePicker
            context={peoplePickerContext}
            personSelectionLimit={10}
            showtooltip={true}
            principalTypes={[PrincipalType.User]}
            defaultSelectedUsers={notifyUsersDefault}
            onChange={onNotifyUsersChange}
          />
        ) : (
          <div style={{ padding: '8px', borderRadius: '4px', backgroundColor: '#e0e0e0', color: '#555', wordBreak: 'break-word' }}>
            {task.NotifyUsersDisplayName && task.NotifyUsersDisplayName.length > 0
              ? task.NotifyUsersDisplayName.join(', ')
              : 'No notify users'}
          </div>
        )}
      </div>

      {/* Attachments */}
      <div style={{ marginBottom: '15px' }}>
        <b>Attachments:</b>
        {canEditAssignedByFull || canEditAssignedTo ? (
          <div style={{ marginBottom: '10px', display: 'flex', flexWrap: 'wrap', gap: '10px', alignItems: 'center' }}>
            <input type="file" multiple onChange={handleFileChange} style={{ flex: '1 1 auto', minWidth: '200px' }} />
            <button
              style={{ padding: '5px 12px', backgroundColor: '#0d6efd', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer', flexShrink: 0 }}
              onClick={handleUploadAttachments}
              disabled={!uploadingFiles || uploadingFiles.length === 0}
            >
              Upload
            </button>
          </div>
        ) : null}

        {attachments && attachments.length > 0 ? (
          <ul style={{ listStyle: 'none', padding: 0, margin: 0 }}>
            {attachments.map((file, idx) => (
              <li
                key={idx}
                style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '6px 10px', marginBottom: '6px', borderRadius: '4px', backgroundColor: '#f1f3f5', flexWrap: 'wrap' }}
              >
                <a href={`${window.location.origin}${file.ServerRelativeUrl}`} target="_blank" rel="noreferrer" style={{ textDecoration: 'none', color: '#0d6efd', wordBreak: 'break-word', flex: '1 1 auto', minWidth: '150px' }}>
                  📎 {file.FileName}
                </a>
                {(canEditAssignedByFull || canEditAssignedTo) && (
                  <button
                    style={{ padding: '2px 6px', backgroundColor: '#dc3545', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '12px', flexShrink: 0, marginLeft: '10px' }}
                    onClick={() => handleDeleteAttachment(file.FileName)}
                  >
                    ❌ Remove
                  </button>
                )}
              </li>
            ))}
          </ul>
        ) : (
          <div style={{ padding: '8px', borderRadius: '4px', backgroundColor: '#f8f9fa', fontStyle: 'italic', color: '#6c757d' }}>
            No attachments uploaded.
          </div>
        )}
      </div>
{showStatusComments && (
  <>
    {/* Status */}
    <div style={{ marginBottom: '15px' }}>
      <label><b>Status:</b></label>
      <select
        style={{ width: '100%', padding: '6px', borderRadius: '4px', border: '1px solid #ccc' }}
        value={status}
        onChange={e => setStatus(e.target.value)}
        disabled={!(canEditAssignedTo || isSameUser || canEditAssignedByStatusOnly)}
      >
        <option value="">-- Select Status --</option>
        {statusOptions.map(opt => <option key={opt} value={opt}>{opt}</option>)}
      </select>
    </div>

    {/* Comments */}
    <div style={{ marginBottom: '15px' }}>
      <label><b>Comments <span className="text-danger">*</span> :</b></label>
      <textarea
        style={{ width: '100%', padding: '6px', borderRadius: '4px', border: '1px solid #ccc' }}
        value={comments}
        onChange={e => setComments(e.target.value)}
        rows={3}
        disabled={!(canEditAssignedTo || isSameUser || canEditAssignedByStatusOnly)}
      />
    </div>

    {/* Buttons */}
    {showSaveButton && (
      <div style={{ display: 'flex', gap: '10px' }}>
        <button
          style={{ padding: '6px 12px', backgroundColor: '#198754', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer' }}
          onClick={handleSave}
        >
          Save
        </button>
        <button
          style={{ padding: '6px 12px', backgroundColor: '#6c757d', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer' }}
          onClick={onCancel}
        >
          Cancel
        </button>
      </div>
    )}
  </>
)}

      {/* Comments History */}
      <div style={{ marginTop: '20px' }}>
        <b>Comments History:</b>
        <div style={{ padding: '8px', borderRadius: '4px', backgroundColor: '#f8f9fa', maxHeight: '150px', overflowY: 'auto', whiteSpace: 'pre-wrap' }}>
          {task.CommentsHistory || '-'}
        </div>
      </div>
    </div>
  );
};

export default FormDetails;