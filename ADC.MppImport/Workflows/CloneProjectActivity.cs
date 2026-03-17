using System;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;

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

            // Normalize start date to noon UTC to avoid timezone edge issues
            DateTime startDateNormalized = newStartDate.Date.AddHours(12);

            // Read source project name to build target name
            var sourceProject = OrganizationService.Retrieve("msdyn_project", sourceRef.Id,
                new ColumnSet("msdyn_subject"));
            string sourceName = sourceProject.GetAttributeValue<string>("msdyn_subject") ?? "Project";

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

            // 1. Create empty target project with desired start date
            var targetProject = new Entity("msdyn_project");
            targetProject["msdyn_subject"] = targetProjectName;
            targetProject["msdyn_scheduledstart"] = startDateNormalized;
            targetProject["adc_parentadccase"] = new EntityReference("adc_case", caseRef.Id);
            Guid targetProjectId = OrganizationService.Create(targetProject);

            TracingService.Trace("CloneProject: Target project created: {0}", targetProjectId);

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
                Success.Set(executionContext, false);
                ResultMessage.Set(executionContext, errMsg);
                NewProject.Set(executionContext, new EntityReference("msdyn_project", targetProjectId));
                return;
            }

            // 3. Poll for completion (async workflows only — synchronous/real-time
            //    workflows run inside a transaction where Thread.Sleep is not allowed)
            var wfContext = executionContext.GetExtension<IWorkflowContext>();
            bool isSynchronous = wfContext != null && wfContext.Mode == 1; // 0=Async, 1=Synchronous

            if (isSynchronous)
            {
                TracingService.Trace("CloneProject: Running in real-time workflow — skipping poll. Copy will complete async.");
                NewProject.Set(executionContext, new EntityReference("msdyn_project", targetProjectId));
                Success.Set(executionContext, true);
                ResultMessage.Set(executionContext, string.Format(
                    "Project copy initiated for '{0}'. Copy is running in the background.", targetProjectName));
                return;
            }

            TracingService.Trace("CloneProject: Polling for copy completion...");
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
            }

            TracingService.Trace("CloneProject: {0}", finalMessage);

            NewProject.Set(executionContext, new EntityReference("msdyn_project", targetProjectId));
            Success.Set(executionContext, copySucceeded);
            ResultMessage.Set(executionContext, finalMessage);
        }
    }
}
