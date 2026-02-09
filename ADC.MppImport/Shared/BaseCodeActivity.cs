using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace ADC.MppImport
{
    public abstract class BaseCodeActivity : CodeActivity
    {
        [Input("Fail on Exception"), Default("true")]
        public InArgument<bool> FailOnException
        {
            get;
            set;
        }

        [Output("Exception Occured"), Default("false")]
        public OutArgument<bool> ExceptionOccured
        {
            get;
            set;
        }

        [Output("Exception Message"), Default("")]
        public OutArgument<string> ExceptionMessage
        {
            get;
            set;
        }

        public IWorkflowContext WorkflowContext
        {
            get;
            set;
        }

        public IOrganizationServiceFactory OrganizationServiceFactory
        {
            get;
            set;
        }

        public IOrganizationService OrganizationService
        {
            get;
            set;
        }

        public ITracingService TracingService
        {
            get;
            set;
        }

        /// <summary>
        /// Executes the workflow activity.
        /// </summary>
        /// <param name="executionContext">The execution context.</param>
        protected override void Execute(CodeActivityContext executionContext)
        {
            if (executionContext == null)
                throw new ArgumentNullException("Code Activity Context is null");

            TracingService = executionContext.GetExtension<ITracingService>();

            if (TracingService == null)
                throw new InvalidPluginExecutionException("Failed to retrieve tracing service.");

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            this.WorkflowContext = context;

            if (context == null)
                throw new InvalidPluginExecutionException("Failed to retrieve workflow context.");

            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            OrganizationServiceFactory = serviceFactory;
            OrganizationService = service;

            try
            {
                this.ExecuteActivity(executionContext);
            }
            catch (Exception e)
            {
                ExceptionOccured.Set(executionContext, true);
                ExceptionMessage.Set(executionContext, e.Message);

                if (FailOnException.Get<bool>(executionContext))
                {
                    throw new InvalidPluginExecutionException(e.Message, e);
                }
            }
        }

        protected abstract void ExecuteActivity(CodeActivityContext executionContext);

    }
}
