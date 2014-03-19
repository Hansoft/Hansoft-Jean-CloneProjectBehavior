using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

using HPMSdk;

using Hansoft.Jean.Behavior;
using Hansoft.ObjectWrapper;

namespace Hansoft.Jean.Behavior.CloneProjectBehavior
{
    public class CloneProjectBehavior : AbstractBehavior
    {

        string title;
        int reportCounter = 0;

        public CloneProjectBehavior(XmlElement configuration)
            : base(configuration) 
        {
            title = "CloneProjectBehavior";
        }

        public override string Title
        {
            get { return title; }
        }

        public override void Initialize()
        {
        }

        // TODO A bit of a hack. Should be more robust.
        void PatchupReleaseTags(List<HansoftItem> sourceSet, List<HansoftItem> targetSet, List<Release> targetReleases)
        {
            for (int i = 0; i < sourceSet.Count; i += 1)
            {
                List<Release> taggedReleases = ((Task)sourceSet[i]).TaggedToReleases;
                if (taggedReleases.Count > 0)
                {
                    HPMTaskLinkedToMilestones linkedReleases = new HPMTaskLinkedToMilestones();
                    linkedReleases.m_Milestones = new HPMUniqueID[taggedReleases.Count];
                    for (int j=0; j < taggedReleases.Count; j+=1)
                    {
                        Release taggedRelease = taggedReleases[j];
                        linkedReleases.m_Milestones[j] = (targetReleases.Find(r => r.Name == taggedRelease.Name && r.Parent.Name == taggedRelease.Parent.Name)).UniqueID;
                    }
                    SessionManager.Session.TaskSetLinkedToMilestones(((Task)targetSet[i]).UniqueTaskID, linkedReleases);
                }
            }
        }

        void CloneChildTasks(HPMUniqueID sourceParent, HPMUniqueID targetParent, HPMUniqueID sourceProject, HPMUniqueID sourceMainProject, HPMUniqueID targetProject, HPMUniqueID targetMainProject, HPMProjectCustomColumns customColumns, bool reverse)
        {
            HPMSdkSession session = SessionManager.Session;
            HPMUniqueID[] childIDs = session.TaskRefUtilEnumChildren(sourceParent, false).m_Tasks;

            HPMUniqueID newTaskRefID = -1;
            for (int i = 0; i < childIDs.Length; i += 1)
            {
                HPMUniqueID originalTaskRefID;
                if (reverse)
                    originalTaskRefID = childIDs[childIDs.Length - 1 - i];
                else
                    originalTaskRefID = childIDs[i];
                HPMUniqueID originalTaskID = session.TaskRefGetTask(originalTaskRefID);
                HPMTaskCreateUnifiedReference prevRefID = new HPMTaskCreateUnifiedReference();
                prevRefID.m_RefID = -1;
                prevRefID.m_bLocalID = false;

                HPMTaskCreateUnifiedReference prevWorkPrioRefID = new HPMTaskCreateUnifiedReference();
                prevWorkPrioRefID.m_RefID = -2;
                prevWorkPrioRefID.m_bLocalID = false;

                HPMTaskCreateUnifiedReference parentRefId = new HPMTaskCreateUnifiedReference();
                parentRefId.m_RefID = targetParent;
                parentRefId.m_bLocalID = false;

                HPMTaskCreateUnified createTaskData = new HPMTaskCreateUnified();
                createTaskData.m_Tasks = new HPMTaskCreateUnifiedEntry[1];
                createTaskData.m_Tasks[0] = new HPMTaskCreateUnifiedEntry();
                createTaskData.m_Tasks[0].m_bIsProxy = false;
                createTaskData.m_Tasks[0].m_LocalID = -1;

                createTaskData.m_Tasks[0].m_ParentRefIDs = new HPMTaskCreateUnifiedReference[1];
                createTaskData.m_Tasks[0].m_ParentRefIDs[0] = parentRefId;
                createTaskData.m_Tasks[0].m_PreviousRefID = prevRefID;
                createTaskData.m_Tasks[0].m_PreviousWorkPrioRefID = prevWorkPrioRefID;
                createTaskData.m_Tasks[0].m_NonProxy_ReuseID = 0;

                createTaskData.m_Tasks[0].m_TaskLockedType = session.TaskGetLockedType(originalTaskID);
                createTaskData.m_Tasks[0].m_TaskType = session.TaskGetType(originalTaskID);

                HPMChangeCallbackData_TaskCreateUnified createdData = session.TaskCreateUnifiedBlock(targetProject, createTaskData);
                if (createdData.m_Tasks.Length == 1)
                {
                    newTaskRefID = createdData.m_Tasks[0].m_TaskRefID;
                    HPMUniqueID newTaskID = session.TaskRefGetTask(newTaskRefID);
                    session.TaskSetBacklogCategory(newTaskID, session.TaskGetBacklogCategory(originalTaskID));
                    session.TaskSetConfidence(newTaskID, session.TaskGetConfidence(originalTaskID));
                    session.TaskSetDetailedDescription(newTaskID, session.TaskGetDetailedDescription(originalTaskID));
                    session.TaskSetEstimatedIdealDays(newTaskID, session.TaskGetEstimatedIdealDays(originalTaskID));
                    session.TaskSetHyperlink(newTaskID, session.TaskGetHyperlink(originalTaskID));
                    session.TaskSetDescription(newTaskID, session.TaskGetDescription(originalTaskID));
                    session.TaskSetComplexityPoints(newTaskID, session.TaskGetComplexityPoints(originalTaskID));
                    session.TaskSetBacklogPriority(newTaskID, session.TaskGetBacklogPriority(originalTaskID));
                    session.TaskSetSprintPriority(newTaskID, session.TaskGetSprintPriority(originalTaskID));
                    session.TaskSetRisk(newTaskID, session.TaskGetRisk(originalTaskID));

                    session.TaskSetSeverity(newTaskID, session.TaskGetSeverity(originalTaskID));
                    session.TaskSetStatus(newTaskID, session.TaskGetStatus(originalTaskID), true, EHPMTaskSetStatusFlag.All);
                    session.TaskSetWorkRemaining(newTaskID, session.TaskGetWorkRemaining(originalTaskID));
                    session.TaskSetRisk(newTaskID, session.TaskGetRisk(originalTaskID));

                    session.TaskSetForceSubProject(newTaskID, session.TaskGetForceSubProject(originalTaskID));

                    session.TaskSetDelegateTo(newTaskID, session.TaskGetDelegateTo(originalTaskID));

                    session.TaskSetFullyCreated(newTaskID);

                    // Cloning custom column data
                    for (int j = 0; j < customColumns.m_ShowingColumns.Length; j += 1)
                    {
                        HPMProjectCustomColumnsColumn column = customColumns.m_ShowingColumns[j];
                        if (column.m_Type == EHPMProjectCustomColumnsColumnType.Hyperlink)
                        {
                            string oldLink = session.TaskGetCustomColumnData(originalTaskID, column.m_Hash);
                            if (oldLink.StartsWith("hansoft://") && oldLink.Contains("/ReportGUID/"))
                            {
                                int ind = oldLink.LastIndexOf("/");
                                string start = oldLink.Substring(0, ind);
                                int ind2 = start.LastIndexOf("/");
                                start = start.Substring(0, ind2);
                                string rest = oldLink.Substring(ind + 1);
                                ulong reportGUID = UInt64.Parse(rest);

                                // TODO This will not work very well in the real world when different users might do this.
                                User adminUser = HPMUtilities.GetUsers().Find(u=> u.Name == "Hansoft Admin");
                                if (adminUser != null)
                                {
                                    HPMUniqueID sourceBacklogProject = session.ProjectUtilGetBacklog(sourceMainProject);
                                    HPMUniqueID sourceQAProject = session.ProjectUtilGetQA(sourceMainProject);
                                    HPMUniqueID targetBacklogProject = session.ProjectUtilGetBacklog(targetMainProject);
                                    HPMUniqueID targetQAProject = session.ProjectUtilGetQA(targetMainProject);
                                    HPMUniqueID sourceReportProject = sourceMainProject;

                                    HPMReports sourceReports = session.ProjectGetReports(sourceMainProject, adminUser.UniqueID);
                                    HPMReport sourceReport = sourceReports.m_Reports.FirstOrDefault(r => r.m_ReportGUID == reportGUID);
                                    if (sourceReport == null)
                                    {
                                        sourceReportProject = sourceBacklogProject;
                                        sourceReports = session.ProjectGetReports(sourceBacklogProject, adminUser.UniqueID);
                                        sourceReport = sourceReports.m_Reports.FirstOrDefault(r => r.m_ReportGUID == reportGUID);
                                    }
                                    if (sourceReport == null)
                                    {
                                        sourceReportProject = sourceQAProject;
                                        sourceReports = session.ProjectGetReports(sourceQAProject, adminUser.UniqueID);
                                        sourceReport = sourceReports.m_Reports.FirstOrDefault(r => r.m_ReportGUID == reportGUID);
                                    }
                                    if (sourceReport != null)
                                    {
                                        HPMReports targetReports;
                                        HPMReport targetReport;
                                        HPMUniqueID targetReportProject;
                                        if (sourceReportProject == sourceMainProject)
                                        {
                                            targetReportProject = targetMainProject;
                                            targetReports = session.ProjectGetReports(targetMainProject, adminUser.UniqueID);
                                            targetReport = targetReports.m_Reports.FirstOrDefault(r => r.m_Name == sourceReport.m_Name);
                                        }
                                        else if (sourceReportProject == sourceBacklogProject)
                                        {
                                            targetReportProject = targetBacklogProject;
                                            targetReports = session.ProjectGetReports(targetBacklogProject, adminUser.UniqueID);
                                            targetReport = targetReports.m_Reports.FirstOrDefault(r => r.m_Name == sourceReport.m_Name);
                                        }
                                        else
                                        {
                                            targetReportProject = targetQAProject;
                                            targetReports = session.ProjectGetReports(targetQAProject, adminUser.UniqueID);
                                            targetReport = targetReports.m_Reports.FirstOrDefault(r => r.m_Name == sourceReport.m_Name);
                                        }
                                        if (targetReport != null)
                                        {
                                            string newLink = start + "/" + targetReportProject.m_ID.ToString() + "/" +targetReport.m_ReportGUID.ToString();
                                            session.TaskSetCustomColumnData(newTaskID, column.m_Hash, newLink, false);
                                        }
                                        else
                                            session.TaskSetCustomColumnData(newTaskID, column.m_Hash, oldLink, false);
                                    }
                                    else
                                      session.TaskSetCustomColumnData(newTaskID, column.m_Hash, oldLink, false);
                                }
                                else
                                    session.TaskSetCustomColumnData(newTaskID, column.m_Hash, oldLink, false);
                            }
                            else
                                session.TaskSetCustomColumnData(newTaskID, column.m_Hash, oldLink, false);
                        }
                        else
                            session.TaskSetCustomColumnData(newTaskID, column.m_Hash, session.TaskGetCustomColumnData(originalTaskID, column.m_Hash), false);
                    }
                    
                    string tempPath = System.IO.Path.GetTempPath();
                    if (!session.VersionControlUtilIsInitialized())
                        session.VersionControlInit(tempPath);
                    string oldAttachmentPath = session.TaskGetAttachmentPath(originalTaskID);
                    string newAttachmentPath = session.TaskGetAttachmentPath(newTaskID);
                    HPMTaskAttachedDocuments oldAttachements = session.TaskGetAttachedDocuments(originalTaskID);
                    int nAttachments = oldAttachements.m_AttachedDocuments.Length;
                    HPMTaskAttachedDocuments attachedDocs = new HPMTaskAttachedDocuments();
                    // Create attachment directory if needed
                    if (nAttachments > 0)
                    {
                        HPMVersionControlFile fInfo = session.VersionControlGetFileInfoBlock(oldAttachmentPath);
                        HPMVersionControlCreateDirectories directories = new HPMVersionControlCreateDirectories();
                        directories.m_Comment = "Created as part of project template.";
                        directories.m_Files = new HPMVersionControlFileSpec[1];
                        directories.m_Files[0] = new HPMVersionControlFileSpec();
                        directories.m_Files[0].m_MetaDataEntries = fInfo.m_MetaDataEntries;
                        directories.m_Files[0].m_Path = newAttachmentPath;
                        session.VersionControlCreateDirectoriesBlock(directories);
                        attachedDocs.m_AttachedDocuments = new HPMTaskAttachedDocumentsEntry[nAttachments];
                    }

                    // Cloning attached documents
                    for (int idoc = 0; idoc < oldAttachements.m_AttachedDocuments.Length; idoc += 1)
                    {
                        HPMTaskAttachedDocumentsEntry oldAttachment = oldAttachements.m_AttachedDocuments[idoc];
                        ulong oldFileId = oldAttachment.m_FileID;
                        string oldFilePath = session.VersionControlUtilFileIDToPath(oldFileId);
                        HPMVersionControlFileList fileList = new HPMVersionControlFileList();
                        fileList.m_Files = new string[1];
                        fileList.m_Files[0] = oldFilePath;
                        HPMChangeCallbackData_VersionControlSyncFilesResponse synchResponse = session.VersionControlSyncFilesBlock(fileList, -1);
                        // TODO There is probably an SDK bug here in that synchResponse always seems to be empty
                        HPMVersionControlFile fInfo = session.VersionControlGetFileInfoBlock(oldFilePath);
                        HPMVersionControlAddFiles addFiles = new HPMVersionControlAddFiles();
                        addFiles.m_bDeleteSourceFiles = false;
                        addFiles.m_Comment = "Added as part of project template.";
                        addFiles.m_FilesToAdd = new HPMVersionControlLocalFilePair[1];
                        addFiles.m_FilesToAdd[0] = new HPMVersionControlLocalFilePair();
                        addFiles.m_FilesToAdd[0].m_LocalPath = tempPath.Replace('\\', '/') + "/" + oldAttachmentPath + "/" + fInfo.m_FileName;
                        addFiles.m_FilesToAdd[0].m_FileSpec.m_Path = newAttachmentPath + "/" + fInfo.m_FileName;
                        HPMChangeCallbackData_VersionControlAddFilesResponse addResponse = session.VersionControlAddFilesBlock(addFiles);
                        if (addResponse.m_Succeeded.Length == 1)
                        {
                            ulong newFileID = session.VersionControlUtilPathToFileID(addFiles.m_FilesToAdd[0].m_FileSpec.m_Path);
                            attachedDocs.m_AttachedDocuments[idoc] = new HPMTaskAttachedDocumentsEntry();
                            attachedDocs.m_AttachedDocuments[idoc].m_AddedByResource = session.ResourceGetLoggedIn();
                            attachedDocs.m_AttachedDocuments[idoc].m_FileID = newFileID;
                        }
                    }
                    if (nAttachments > 0)
                        session.TaskSetAttachedDocuments(newTaskID, attachedDocs);
                    CloneChildTasks(originalTaskRefID, newTaskRefID, sourceProject, sourceMainProject, targetProject, targetMainProject, customColumns, false);
                }
            }
        }

        void CloneColumns(HPMUniqueID sourceProjectViewID, HPMUniqueID targetProjectViewID)
        {
            HPMSdkSession session = SessionManager.Session;
            session.ProjectCustomColumnsSet(targetProjectViewID, session.ProjectCustomColumnsGet(sourceProjectViewID));
            HPMProjectDefaultColumns activatedColumns = session.ProjectGetDefaultActivatedColumns(sourceProjectViewID);
            session.ProjectSetDefaultActivatedColumns(targetProjectViewID, activatedColumns);
        }

        void CloneReports(HPMUniqueID sourceProjectID, HPMUniqueID sourceProjectViewID, HPMUniqueID targetProjectViewID, HPMUniqueID creatorID)
        {
            HPMSdkSession session = SessionManager.Session;
            HPMProjectResourceEnum sourceResources = session.ProjectResourceEnum(sourceProjectID);
            for (int i = 0; i < sourceResources.m_Resources.Length; i += 1)
            {
                HPMReports oldReports = session.ProjectGetReports(sourceProjectViewID, sourceResources.m_Resources[i]);
                if (oldReports.m_Reports.Length > 0)
                {
                    for (int iReport = 0; iReport < oldReports.m_Reports.Length; iReport += 1)
                    {
                        reportCounter += 1;
                        oldReports.m_Reports[iReport].m_ProjectID = targetProjectViewID;
                        oldReports.m_Reports[iReport].m_ReportGUID = 0;
                        oldReports.m_Reports[iReport].m_ReportID = reportCounter;
                        oldReports.m_Reports[iReport].m_ResourceID = creatorID;
                    }
                    session.ProjectSetReports(targetProjectViewID, creatorID, oldReports);
                }
            }
        }

        void ClonePresets(HPMUniqueID sourceProjectViewID, HPMUniqueID targetProjectViewID)
        {
            HPMSdkSession session = SessionManager.Session;
            HPMProjectViewPresets oldPresets = session.ProjectGetViewPresets(sourceProjectViewID);
            session.ProjectSetViewPresets(targetProjectViewID, oldPresets);
        }

        void CloneWorkflows(HPMUniqueID sourceProjectViewID, HPMUniqueID targetProjectViewID)
        {
            HPMSdkSession session = SessionManager.Session;
            uint[] workflows = session.ProjectWorkflowEnum(sourceProjectViewID, true).m_Workflows;
            for (int i = 0; i < workflows.Length; i += 1)
            {
                HPMProjectWorkflowSettings settings = session.ProjectWorkflowGetSettings(sourceProjectViewID, workflows[i]);
                session.ProjectWorkflowCreate(targetProjectViewID, settings.m_Properties);
                session.ProjectWorkflowSetSettings(targetProjectViewID, settings.m_Identifier, settings);
            }
        }

        void CloneBugWorkflow(HPMUniqueID sourceQAProjectID, HPMUniqueID targetQAProjectID)
        {
        }

        public override void OnProjectCreate(ProjectCreateEventArgs e)
        {
            Project createdProject = HPMUtilities.GetProjects().Find(p => p.UniqueID == e.Data.m_ProjectID);
            if (createdProject != null)
            {
                HPMSdkSession session = SessionManager.Session;
                HPMProjectProperties properties = session.ProjectGetProperties(createdProject.UniqueID);
                if (properties.m_SortName.Contains(":Template - "))
                {
                    int colonIndex = properties.m_SortName.IndexOf(':');
                    string creatorIDString = properties.m_SortName.Substring(0, colonIndex);
                    string templateName = properties.m_SortName.Substring(colonIndex + 1);
                    Project templateProject = HPMUtilities.FindProject(templateName);
                    HPMUniqueID creatorID = new HPMUniqueID();
                    creatorID.m_ID = Int32.Parse(creatorIDString);
                    properties.m_SortName = "";
                    session.ProjectSetProperties(createdProject.UniqueID, properties);
                    if (templateProject != null)
                    {
                        HPMUniqueID sourceProjectID = templateProject.UniqueID;
                        HPMUniqueID sourceBacklogID = templateProject.ProductBacklog.UniqueID;
                        HPMUniqueID targetProjectID = createdProject.UniqueID;
                        HPMUniqueID targetBacklogID = createdProject.ProductBacklog.UniqueID;
                        HPMUniqueID targetQAProjectID = createdProject.BugTracker.UniqueID;
                        HPMUniqueID sourceQAProjectID = createdProject.BugTracker.UniqueID;

                        CloneColumns(sourceProjectID, targetProjectID);
                        CloneColumns(sourceBacklogID, targetBacklogID);
                        CloneColumns(sourceQAProjectID, targetQAProjectID);

                        CloneReports(sourceProjectID, sourceProjectID, targetProjectID, creatorID);
                        CloneReports(sourceProjectID, sourceBacklogID, targetBacklogID, creatorID);
                        CloneReports(sourceProjectID, sourceQAProjectID, targetQAProjectID, creatorID);

                        ClonePresets(sourceProjectID, targetProjectID);
                        ClonePresets(sourceBacklogID, targetBacklogID);
                        ClonePresets(sourceQAProjectID, targetQAProjectID);

                        CloneWorkflows(sourceProjectID, targetProjectID);

                        CloneBugWorkflow(sourceQAProjectID, targetQAProjectID);


                        CloneChildTasks(sourceProjectID, createdProject.UniqueID, sourceProjectID, sourceProjectID, createdProject.UniqueID, createdProject.UniqueID, session.ProjectCustomColumnsGet(sourceProjectID), true);
                        CloneChildTasks(sourceBacklogID, createdProject.ProductBacklog.UniqueID, sourceBacklogID, sourceProjectID, createdProject.ProductBacklog.UniqueID, createdProject.UniqueID, session.ProjectCustomColumnsGet(sourceBacklogID), true);
                        
                        Project sourceProject = HPMUtilities.GetProjects().Find(p => p.UniqueID == sourceProjectID);
                        Project targetProject = HPMUtilities.GetProjects().Find(p => p.UniqueID == targetProjectID);

                        PatchupReleaseTags(sourceProject.Schedule.DeepChildren, targetProject.Schedule.DeepChildren, targetProject.Releases);
                        PatchupReleaseTags(sourceProject.ProductBacklog.DeepChildren, targetProject.ProductBacklog.DeepChildren, targetProject.Releases);
                    }
                }
            }
        }
    }
}
