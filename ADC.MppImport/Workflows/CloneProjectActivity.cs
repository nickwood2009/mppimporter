using System;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using ADC.MppImport.Services;
using System.Collections.Generic;

namespace ADC.MppImport.Workflows
{
    /// <summary>
    /// Custom workflow activity that clones an existing Project Operations project
    /// into a new project linked to an adc_case, with a new start date.
    /// Uses the msdyn_CopyProjectV3 PSS action.
    /// </summary>
    public class CloneProjectActivity : BaseCodeActivity
    {
        [Input("Source Project")]
        [ReferenceTarget("msdyn_project")]
        [RequiredArgument]
        public InArgument<EntityReference> SourceProject { get; set; }

        [Input("Target Case")]
        [ReferenceTarget("adc_case")]
        [RequiredArgument]
        public InArgument<EntityReference> TargetCase { get; set; }

        [Input("New Start Date")]
        [RequiredArgument]
        public InArgument<DateTime> NewStartDate { get; set; }

        [Input("Clear Teams And Assignments"), Default("True")]
        public InArgument<bool> ClearTeamsAndAssignments { get; set; }

        [Output("New Project")]
        [ReferenceTarget("msdyn_project")]
        public OutArgument<EntityReference> NewProject { get; set; }

        [Output("Success")]
        public OutArgument<bool> Success { get; set; }

        [Output("Result Message")]
        public OutArgument<string> ResultMessage { get; set; }

        private const int POLL_INTERVAL_SECONDS = 10;
        private const int POLL_MAX_ATTEMPTS = 30; // 5 minutes max
        private const int TASK_WAIT_INTERVAL_SECONDS = 5;
        private const int TASK_WAIT_MAX_ATTEMPTS = 12; // 60 seconds max

        protected override void ExecuteActivity(CodeActivityContext executionContext)
        {
            var sourceRef = SourceProject.Get(executionContext);
            var caseRef = TargetCase.Get(executionContext);
            var newStartDate = NewStartDate.Get(executionContext);
            bool clearTeams = ClearTeamsAndAssignments.Get(executionContext);

            if (sourceRef == null)
                throw new InvalidPluginExecutionException("Source Project input is required.");
            if (caseRef == null)
                throw new InvalidPluginExecutionException("Target Case input is required.");
            if (newStartDate == default(DateTime))
                throw new InvalidPluginExecutionException("New Start Date input is required.");

            TracingService.Trace("CloneProject: Source={0}, Case={1}, NewStart={2:yyyy-MM-dd}, ClearTeams={3}",
                sourceRef.Id, caseRef.Id, newStartDate, clearTeams);

            // Resolve initiating user for notifications
            Guid? initiatingUserId = null;
            var wfCtx = executionContext.GetExtension<IWorkflowContext>();
            if (wfCtx != null)
                initiatingUserId = wfCtx.InitiatingUserId;


            // Read source project name + calendar + start time fields to replicate on target
            var sourceProject = OrganizationService.Retrieve("msdyn_project", sourceRef.Id,
                new ColumnSet("msdyn_subject", "msdyn_calendarid", "msdyn_workhourtemplate", "msdyn_scheduledstart"));
            string sourceName = sourceProject.GetAttributeValue<string>("msdyn_subject") ?? "Project";

            // Capture calendar + start-time info from source so we can replicate on target
            string sourceCalendarId = sourceProject.GetAttributeValue<string>("msdyn_calendarid");
            var sourceWorkHourTemplate = sourceProject.GetAttributeValue<EntityReference>("msdyn_workhourtemplate");
            var sourceStart = sourceProject.GetAttributeValue<DateTime?>("msdyn_scheduledstart");
            TracingService.Trace("CloneProject: Source calendar ID = {0}, WorkHourTemplate = {1}",
                sourceCalendarId ?? "(null)",
                sourceWorkHourTemplate != null ? sourceWorkHourTemplate.Id.ToString() : "(null)");
            TracingService.Trace("CloneProject: Source scheduledstart = {0}",
                sourceStart.HasValue ? sourceStart.Value.ToString("o") : "(null)");

            // Use source project's start-time-of-day so PSS anchors tasks at same hour (e.g. 9 AM)
            // NOTE: CopyProjectV3 reschedules tasks internally — DST-aware timezone conversion
            // (as used in MppAsyncImportService) is NOT compatible here and causes fractional-day
            // durations. Keep this as a simple UTC hour copy for now.
            // Fallback to midnight UTC if source has no start (Dataverse will snap to calendar working-hours start)
            int startHour = sourceStart.HasValue ? sourceStart.Value.Hour : 0;
            int startMinute = sourceStart.HasValue ? sourceStart.Value.Minute : 0;
            DateTime startDateNormalized = newStartDate.Date.AddHours(startHour).AddMinutes(startMinute);
            TracingService.Trace("CloneProject: Target start = {0:o} (hour={1}, min={2} from source)",
                startDateNormalized, startHour, startMinute);

            // Read case details for naming
            var caseRecord = OrganizationService.Retrieve("adc_case", caseRef.Id,
                new ColumnSet("adc_name", "adc_casenumber"));
            string caseName = caseRecord.GetAttributeValue<string>("adc_name") ?? "Case";
            string caseNumber = caseRecord.GetAttributeValue<string>("adc_casenumber");
            string targetProjectName = !string.IsNullOrEmpty(caseNumber)
                ? string.Format("{0} - {1}", caseName, caseNumber)
                : caseName;

            TracingService.Trace("CloneProject: Creating target project '{0}' with start {1:yyyy-MM-dd}...",
                targetProjectName, startDateNormalized);

            // Update case status to Processing + send start notification
            UpdateCaseStatus(caseRef.Id, 1, "Cloning template project...");
            if (initiatingUserId.HasValue)
            {
                SendNotification(initiatingUserId.Value, "Project Clone Started",
                    string.Format("Cloning template project into '{0}' — running in background.", targetProjectName),
                    NotificationIconType.Info, caseRef.Id);
            }

            // 1. Create empty target project with desired start date + same work-hour template as source
            //    Do NOT copy msdyn_calendarid — each project needs its own calendar record;
            //    setting msdyn_workhourtemplate tells Dataverse which template to generate the calendar from.
            var targetProject = new Entity("msdyn_project");
            targetProject["msdyn_subject"] = targetProjectName;
            targetProject["msdyn_scheduledstart"] = startDateNormalized;
            targetProject["adc_parentadccase"] = new EntityReference("adc_case", caseRef.Id);
            if (sourceWorkHourTemplate != null)
                targetProject["msdyn_workhourtemplate"] = sourceWorkHourTemplate;
            Guid targetProjectId = OrganizationService.Create(targetProject);

            TracingService.Trace("CloneProject: Target project created: {0}", targetProjectId);

            // Verify target project calendar + start time were set correctly
            try
            {
                var targetCheck = OrganizationService.Retrieve("msdyn_project", targetProjectId,
                    new ColumnSet("msdyn_calendarid", "msdyn_workhourtemplate", "msdyn_scheduledstart"));
                string targetCalId = targetCheck.GetAttributeValue<string>("msdyn_calendarid");
                var targetWht = targetCheck.GetAttributeValue<EntityReference>("msdyn_workhourtemplate");
                var targetStart = targetCheck.GetAttributeValue<DateTime?>("msdyn_scheduledstart");
                TracingService.Trace("CloneProject: Target calendar ID = {0}, WorkHourTemplate = {1}, Start = {2}",
                    targetCalId ?? "(null)",
                    targetWht != null ? targetWht.Id.ToString() : "(null)",
                    targetStart.HasValue ? targetStart.Value.ToString("o") : "(null)");

                // Work-hour template match is what matters (calendarid will differ since each project gets its own copy)
                bool whtMatch = (sourceWorkHourTemplate == null && targetWht == null)
                    || (sourceWorkHourTemplate != null && targetWht != null
                        && sourceWorkHourTemplate.Id == targetWht.Id);
                TracingService.Trace("CloneProject: WorkHourTemplate match = {0}", whtMatch);
                if (!whtMatch)
                    TracingService.Trace("CloneProject: WARNING — WorkHourTemplate mismatch! Source='{0}', Target='{1}'",
                        sourceWorkHourTemplate != null ? sourceWorkHourTemplate.Id.ToString() : "(null)",
                        targetWht != null ? targetWht.Id.ToString() : "(null)");
            }
            catch (Exception ex)
            {
                TracingService.Trace("CloneProject: Calendar verification failed (non-fatal): {0}", ex.Message);
            }

            // Link project to case
            try
            {
                var caseUpdate = new Entity("adc_case", caseRef.Id);
                caseUpdate["adc_projectid"] = new EntityReference("msdyn_project", targetProjectId);
                OrganizationService.Update(caseUpdate);
            }
            catch (Exception ex)
            {
                TracingService.Trace("CloneProject: Failed to link project to case (non-fatal): {0}", ex.Message);
            }

            // 2. Call msdyn_CopyProjectV3
            TracingService.Trace("CloneProject: Calling msdyn_CopyProjectV3...");
            UpdateCaseStatus(caseRef.Id, 1, "Copying project tasks and dependencies...");
            try
            {
                var copyRequest = new OrganizationRequest("msdyn_CopyProjectV3");
                copyRequest["SourceProject"] = sourceRef;
                copyRequest["Target"] = new EntityReference("msdyn_project", targetProjectId);
                if (clearTeams)
                    copyRequest["ClearTeamsAndAssignments"] = true;
                else
                    copyRequest["ReplaceNamedResources"] = true;
                OrganizationService.Execute(copyRequest);
                TracingService.Trace("CloneProject: CopyProjectV3 initiated.");
            }
            catch (Exception ex)
            {
                string errMsg = string.Format("CopyProjectV3 failed: {0}", ex.Message);
                TracingService.Trace("CloneProject: {0}", errMsg);
                UpdateCaseStatus(caseRef.Id, 4, errMsg);
                if (initiatingUserId.HasValue)
                    SendNotification(initiatingUserId.Value, "Project Clone Failed", errMsg, NotificationIconType.Failure, caseRef.Id);
                Success.Set(executionContext, false);
                ResultMessage.Set(executionContext, errMsg);
                NewProject.Set(executionContext, new EntityReference("msdyn_project", targetProjectId));
                return;
            }

            // 3. Poll for copy completion
            int wfMode = wfCtx != null ? wfCtx.Mode : -1;
            TracingService.Trace("CloneProject: Workflow Mode={0} (0=Async, 1=Sync). Proceeding with poll...", wfMode);
            bool copySucceeded = false;
            string finalMessage = "Copy timed out after polling.";

            for (int attempt = 0; attempt < POLL_MAX_ATTEMPTS; attempt++)
            {
                System.Threading.Thread.Sleep(POLL_INTERVAL_SECONDS * 1000);

                var check = OrganizationService.Retrieve("msdyn_project", targetProjectId,
                    new ColumnSet("statuscode"));
                int statusCode = check.GetAttributeValue<OptionSetValue>("statuscode")?.Value ?? 0;

                TracingService.Trace("CloneProject: Poll {0}/{1} — statuscode={2}",
                    attempt + 1, POLL_MAX_ATTEMPTS, statusCode);

                // statuscode 1 = Active (copy done), 192350000 = Project copy failed
                if (statusCode == 1)
                {
                    copySucceeded = true;
                    finalMessage = string.Format("Project cloned successfully as '{0}'.", targetProjectName);
                    break;
                }
                else if (statusCode == 192350000)
                {
                    finalMessage = "Project copy failed. Check PSS Error Logs for details.";
                    break;
                }
                else
                {
                    UpdateCaseStatus(caseRef.Id, 1, string.Format("Copying project... (poll {0}/{1})", attempt + 1, POLL_MAX_ATTEMPTS));
                }
            }

            TracingService.Trace("CloneProject: {0}", finalMessage);

            // Update case with final status + recalculate day counts + send notification
            if (copySucceeded)
            {
                // Wait for tasks to be created by PSS, then recalculate day counts
                UpdateCaseStatus(caseRef.Id, 1, "Project copied — calculating day counts...");
                bool dayCountSuccess = WaitForTasksAndRecalc(targetProjectId);

                if (dayCountSuccess)
                    finalMessage = string.Format("Project cloned and day counts calculated for '{0}'.", targetProjectName);
                else
                    finalMessage = string.Format("Project cloned as '{0}' but day count calculation had issues — see trace log.", targetProjectName);

                UpdateCaseStatus(caseRef.Id, 2, finalMessage);
                if (initiatingUserId.HasValue)
                    SendNotification(initiatingUserId.Value, "Project Clone Complete", finalMessage, NotificationIconType.Success, caseRef.Id);
            }
            else
            {
                UpdateCaseStatus(caseRef.Id, 4, finalMessage);
                if (initiatingUserId.HasValue)
                    SendNotification(initiatingUserId.Value, "Project Clone Failed", finalMessage, NotificationIconType.Failure, caseRef.Id);
            }

            NewProject.Set(executionContext, new EntityReference("msdyn_project", targetProjectId));
            Success.Set(executionContext, copySucceeded);
            ResultMessage.Set(executionContext, finalMessage);
        }

        /// <summary>
        /// Waits for tasks to appear on the project, then recalculates day counts.
        /// Waits until the task count stabilises AND tasks have finish dates before
        /// running the recalc — PSS creates tasks asynchronously after CopyProjectV3.
        /// Returns true if recalc ran successfully.
        /// </summary>
        private bool WaitForTasksAndRecalc(Guid projectId)
        {
            TracingService.Trace("CloneProject: Waiting for tasks to appear on project {0}...", projectId);

            int taskCount = 0;
            int prevCount = -1;
            int stableRounds = 0;
            const int STABLE_REQUIRED = 3; // count must be unchanged for 3 consecutive polls

            for (int attempt = 0; attempt < TASK_WAIT_MAX_ATTEMPTS; attempt++)
            {
                var query = new QueryExpression("msdyn_projecttask")
                {
                    ColumnSet = new ColumnSet("msdyn_scheduledend"),
                    Criteria =
                    {
                        Conditions =
                        {
                            new ConditionExpression("msdyn_project", ConditionOperator.Equal, projectId)
                        }
                    }
                };
                var results = OrganizationService.RetrieveMultiple(query);
                taskCount = results.Entities.Count;

                // Count how many already have finish dates
                int withFinish = 0;
                foreach (var t in results.Entities)
                {
                    if (t.GetAttributeValue<DateTime?>("msdyn_scheduledend") != null)
                        withFinish++;
                }

                TracingService.Trace("CloneProject: Task wait {0}/{1} — {2} tasks found ({3} with finish dates).",
                    attempt + 1, TASK_WAIT_MAX_ATTEMPTS, taskCount, withFinish);

                // Wait until count > 0, count is stable, and all tasks have finish dates
                if (taskCount > 0 && taskCount == prevCount && withFinish == taskCount)
                {
                    stableRounds++;
                    if (stableRounds >= STABLE_REQUIRED)
                    {
                        TracingService.Trace("CloneProject: Task count stable at {0} for {1} rounds — proceeding.",
                            taskCount, stableRounds);
                        break;
                    }
                }
                else
                {
                    stableRounds = 0;
                }

                prevCount = taskCount;
                System.Threading.Thread.Sleep(TASK_WAIT_INTERVAL_SECONDS * 1000);
            }

            if (taskCount == 0)
            {
                TracingService.Trace("CloneProject: No tasks found after {0}s — skipping day count recalc.",
                    TASK_WAIT_INTERVAL_SECONDS * TASK_WAIT_MAX_ATTEMPTS);
                return false;
            }

            try
            {
                var dayCountService = new DayCountService(OrganizationService, TracingService);
                dayCountService.RecalcAllTasks(projectId);
                return true;
            }
            catch (Exception ex)
            {
                TracingService.Trace("CloneProject: Day count recalc failed (non-fatal): {0}", ex.Message);
                return false;
            }
        }

        private void UpdateCaseStatus(Guid caseId, int importStatus, string message)
        {
            try
            {
                var caseUpdate = new Entity("adc_case", caseId);
                caseUpdate["adc_importstatus"] = new OptionSetValue(importStatus);
                caseUpdate["adc_importmessage"] = message != null && message.Length > 100
                    ? message.Substring(0, 97) + "..." : message;
                OrganizationService.Update(caseUpdate);
            }
            catch (Exception ex)
            {
                TracingService.Trace("CloneProject: UpdateCaseStatus failed (non-fatal): {0}", ex.Message);
            }
        }

        private void SendNotification(Guid recipientUserId, string title, string body, int iconType, Guid? caseId = null)
        {
            try
            {
                var notification = new Entity("appnotification");
                notification["title"] = title;
                notification["body"] = body;
                notification["ownerid"] = new EntityReference("systemuser", recipientUserId);
                notification["icontype"] = new OptionSetValue(iconType);
                notification["toasttype"] = new OptionSetValue(200000000); // Timed (shows toast)

                if (caseId.HasValue)
                {
                    notification["data"] = string.Format(
                        "{{\"actions\":[{{\"title\":\"Open Case\",\"data\":{{\"url\":\"?pagetype=entityrecord&etn=adc_case&id={0}\",\"navigationTarget\":\"dialog\"}}}}]}}",
                        caseId.Value);
                }

                OrganizationService.Create(notification);
                TracingService.Trace("CloneProject: Notification sent to {0}: {1}", recipientUserId, title);
            }
            catch (Exception ex)
            {
                TracingService.Trace("CloneProject: SendNotification failed (non-fatal): {0}", ex.Message);
            }
        }
    }
}
