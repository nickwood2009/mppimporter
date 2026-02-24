using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using ADC.MppImport.MppReader.Mpp;
using ADC.MppImport.MppReader.Model;
using ADC.MppImport.Services;

namespace ADC.MppImport.Test
{
    class Program
    {
        // Well-known sample app registration for Dynamics 365 OAuth
        private const string DefaultAppId = "51f81489-12ee-4a9e-aaae-a2591f45987d";
        private const string DefaultRedirectUri = "app://58145B91-0C36-4500-8554-080854F2AC97";

        static int Main(string[] args)
        {
            // CLI mode: parse <path>
            if (args.Length >= 2 && args[0].Equals("parse", StringComparison.OrdinalIgnoreCase))
            {
                return ParseLocalMppFile(args[1]);
            }

            // CLI mode: import <crmUrl> <templateId> <projectId> [appId]
            if (args.Length >= 4 && args[0].Equals("import", StringComparison.OrdinalIgnoreCase))
            {
                return RunCrmImportFromArgs(args);
            }

            // CLI mode: imports2s <crmUrl> <templateId> <projectId> <clientId> <clientSecret>
            if (args.Length >= 6 && args[0].Equals("imports2s", StringComparison.OrdinalIgnoreCase))
            {
                return RunCrmImportS2SFromArgs(args);
            }

            // Legacy CLI: single arg = mpp file path
            if (args.Length == 1 && File.Exists(args[0]))
            {
                return ParseLocalMppFile(args[0]);
            }

            // Show usage if bad args were passed
            if (args.Length > 0)
            {
                PrintUsage();
                return 1;
            }

            // Interactive menu loop
            while (true)
            {
                Console.WriteLine();
                Console.WriteLine("========================================");
                Console.WriteLine("  ADC MPP Import - Test Console");
                Console.WriteLine("========================================");
                Console.WriteLine("  1. Parse a local MPP file");
                Console.WriteLine("  2. Import MPP to Dynamics 365 (browser login)");
                Console.WriteLine("  3. Import MPP to Dynamics 365 (client secret)");
                Console.WriteLine("  4. Import MPP to Dynamics 365 (from .env file)");
                Console.WriteLine("  5. Clear all project tasks (from .env file)");
                Console.WriteLine("  6. Clear + Import (from .env file)");
                Console.WriteLine("  7. Exit");
                Console.WriteLine("========================================");
                Console.Write("Select option: ");

                string choice = Console.ReadLine()?.Trim();
                switch (choice)
                {
                    case "1":
                        RunLocalMppParse();
                        break;
                    case "2":
                        RunCrmImport();
                        break;
                    case "3":
                        RunCrmImportS2S();
                        break;
                    case "4":
                        RunCrmImportFromEnv();
                        break;
                    case "5":
                        RunClearProjectTasks();
                        break;
                    case "6":
                        RunClearAndImport();
                        break;
                    case "7":
                        return 0;
                    default:
                        Console.WriteLine("Invalid option. Please enter 1-7.");
                        break;
                }
            }
        }

        static void PrintUsage()
        {
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("  ADC.MppImport.Test.exe                                           (interactive menu)");
            Console.WriteLine("  ADC.MppImport.Test.exe parse <mppFilePath>                       (parse local MPP)");
            Console.WriteLine("  ADC.MppImport.Test.exe import <crmUrl> <templateId> <projectId>  (CRM import)");
            Console.WriteLine("  ADC.MppImport.Test.exe import <crmUrl> <templateId> <projectId> <appId>");
            Console.WriteLine("  ADC.MppImport.Test.exe imports2s <crmUrl> <templateId> <projectId> <clientId> <clientSecret>");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  ADC.MppImport.Test.exe parse \"C:\\files\\project.mpp\"");
            Console.WriteLine("  ADC.MppImport.Test.exe import https://myorg.crm6.dynamics.com 1a2b3c4d-... 5e6f7a8b-...");
            Console.WriteLine("  ADC.MppImport.Test.exe imports2s https://myorg.crm6.dynamics.com 1a2b3c4d-... 5e6f7a8b-... <clientId> <secret>");
            Console.WriteLine();
            Console.WriteLine("The default Azure App ID is the well-known Microsoft sample app:");
            Console.WriteLine("  {0}", DefaultAppId);
            Console.WriteLine("This works for any Dynamics 365 org without registering your own app.");
            Console.WriteLine("To use your own, create an App Registration in Azure Entra ID");
            Console.WriteLine("(portal.azure.com > App registrations > add Dynamics CRM user_impersonation).");
        }

        #region Option 1: Parse Local MPP File

        static void RunLocalMppParse()
        {
            Console.WriteLine();
            Console.Write("Enter path to MPP file: ");
            string filePath = Console.ReadLine()?.Trim().Trim('"');

            if (string.IsNullOrEmpty(filePath))
            {
                Console.WriteLine("No path entered.");
                return;
            }

            ParseLocalMppFile(filePath);
        }

        static int ParseLocalMppFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine("File not found: " + filePath);
                return 1;
            }

            try
            {
                byte[] data = File.ReadAllBytes(filePath);
                Console.WriteLine();
                Console.WriteLine("File: {0} ({1:N0} bytes)", Path.GetFileName(filePath), data.Length);

                var reader = new MppFileReader();
                ProjectFile project = reader.Read(data);

                Console.WriteLine("MPP Type:     {0}", project.ProjectProperties.MppFileType);
                Console.WriteLine("Tasks:        {0}", project.Tasks.Count);
                Console.WriteLine("Resources:    {0}", project.Resources.Count);
                Console.WriteLine("Assignments:  {0}", project.Assignments.Count);
                Console.WriteLine("Calendars:    {0}", project.Calendars.Count);

                int totalPredecessors = 0;
                foreach (var t2 in project.Tasks) totalPredecessors += t2.Predecessors.Count;
                Console.WriteLine("Relations:    {0}", totalPredecessors);
                Console.WriteLine("Errors:       {0}", project.IgnoredErrors.Count);
                if (project.IgnoredErrors.Count > 0)
                {
                    foreach (var err in project.IgnoredErrors)
                        Console.WriteLine("  ERR: {0}", err.Message);
                }

                // Dump reader diagnostics
                if (project.DiagnosticMessages.Count > 0)
                {
                    Console.WriteLine();
                    Console.WriteLine("--- Reader Diagnostics ---");
                    foreach (var msg in project.DiagnosticMessages)
                        Console.WriteLine("  {0}", msg);
                }

                Console.WriteLine();
                Console.WriteLine("--- Tasks ---");
                int shown = 0;
                foreach (var t in project.Tasks)
                {
                    if (shown++ >= 50) { Console.WriteLine("  ... ({0} more)", project.Tasks.Count - 50); break; }
                    string preds = t.Predecessors.Count > 0
                        ? string.Join(",", t.Predecessors.Select(p => p.SourceTaskUniqueID + "(" + p.Type + ")"))
                        : "";
                    string durStr = t.Duration != null
                        ? string.Format("{0} {1}", t.Duration.Value, t.Duration.Units)
                        : "null";
                    Console.WriteLine("  [{0,3}] L{1} {2,-40} Start={3:d}  Dur={4,-18} Pred={5}",
                        t.UniqueID, t.OutlineLevel, Trunc(t.Name, 40),
                        t.Start, durStr, preds);
                }

                if (project.Resources.Count > 0)
                {
                    Console.WriteLine();
                    Console.WriteLine("--- Resources ---");
                    shown = 0;
                    foreach (var r in project.Resources)
                    {
                        if (shown++ >= 20) { Console.WriteLine("  ... ({0} more)", project.Resources.Count - 20); break; }
                        Console.WriteLine("  [{0,3}] {1}", r.UniqueID, r.Name);
                    }
                }

                if (project.Assignments.Count > 0)
                {
                    Console.WriteLine();
                    Console.WriteLine("--- Assignments ---");
                    shown = 0;
                    foreach (var a in project.Assignments)
                    {
                        if (shown++ >= 20) { Console.WriteLine("  ... ({0} more)", project.Assignments.Count - 20); break; }
                        Console.WriteLine("  Task={0,3}  Resource={1,3}  Work={2}",
                            a.TaskUniqueID, a.ResourceUniqueID, a.Work);
                    }
                }

                Console.WriteLine();
                Console.WriteLine("=== PASS ===");
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine("=== FAIL ===");
                Console.WriteLine(ex.ToString());
                return 2;
            }
        }

        #endregion

        #region Option 2: CRM Import

        static int RunCrmImportFromArgs(string[] args)
        {
            // args: import <crmUrl> <templateId> <projectId> [appId]
            string crmUrl = args[1].TrimEnd('/');
            if (!crmUrl.StartsWith("http", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("ERROR: CRM URL must start with https://");
                return 1;
            }

            Guid caseTemplateId;
            if (!Guid.TryParse(args[2], out caseTemplateId))
            {
                Console.WriteLine("ERROR: '{0}' is not a valid GUID for Case Template ID.", args[2]);
                return 1;
            }

            Guid projectId;
            if (!Guid.TryParse(args[3], out projectId))
            {
                Console.WriteLine("ERROR: '{0}' is not a valid GUID for Project ID.", args[3]);
                return 1;
            }

            string appId = args.Length >= 5 ? args[4] : DefaultAppId;
            string redirectUri = args.Length >= 5 ? "http://localhost" : DefaultRedirectUri;

            Console.WriteLine();
            Console.WriteLine("  CRM URL:          {0}", crmUrl);
            Console.WriteLine("  Case Template:    {0}", caseTemplateId);
            Console.WriteLine("  Target Project:   {0}", projectId);
            Console.WriteLine("  App ID:           {0}", appId);

            return ExecuteCrmImport(crmUrl, caseTemplateId, projectId, appId, redirectUri);
        }

        static void RunCrmImport()
        {
            Console.WriteLine();
            Console.WriteLine("--- Dynamics 365 MPP Import ---");
            Console.WriteLine();

            // Step 1: Get CRM URL
            string crmUrl = PromptRequired("Dynamics 365 URL (e.g. https://yourorg.crm6.dynamics.com): ");
            if (crmUrl == null) return;
            crmUrl = crmUrl.TrimEnd('/');
            if (!crmUrl.StartsWith("http", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("ERROR: URL must start with https://");
                return;
            }

            // Step 2: Get Case Template ID
            Guid caseTemplateId = PromptGuid("adc_adccasetemplate ID (GUID): ");
            if (caseTemplateId == Guid.Empty) return;

            // Step 3: Get Project ID
            Guid projectId = PromptGuid("msdyn_project ID (GUID): ");
            if (projectId == Guid.Empty) return;

            // Step 4: Optional custom App ID
            Console.Write("Azure App ID (Enter for default sample app): ");
            string appIdInput = Console.ReadLine()?.Trim();
            string appId = string.IsNullOrEmpty(appIdInput) ? DefaultAppId : appIdInput;
            string redirectUri = string.IsNullOrEmpty(appIdInput) ? DefaultRedirectUri : "http://localhost";

            // Confirm
            Console.WriteLine();
            Console.WriteLine("  CRM URL:          {0}", crmUrl);
            Console.WriteLine("  Case Template:    {0}", caseTemplateId);
            Console.WriteLine("  Target Project:   {0}", projectId);
            Console.WriteLine("  App ID:           {0}", appId);
            Console.WriteLine();
            Console.Write("Proceed? (Y/N): ");
            string confirm = Console.ReadLine()?.Trim().ToUpper();
            if (confirm != "Y") { Console.WriteLine("Cancelled."); return; }

            ExecuteCrmImport(crmUrl, caseTemplateId, projectId, appId, redirectUri);
        }

        static int ExecuteCrmImport(string crmUrl, Guid caseTemplateId, Guid projectId, string appId, string redirectUri)
        {
            // Connect to CRM (browser login)
            Console.WriteLine();
            Console.WriteLine("Connecting to Dynamics 365 (browser login will open)...");

            IOrganizationService service = null;
            try
            {
                service = ConnectToCrm(crmUrl, appId, redirectUri);
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Failed to connect to CRM.");
                Console.WriteLine(ex.Message);
                return 1;
            }

            if (service == null)
            {
                Console.WriteLine("ERROR: CRM connection returned null. Check your credentials and URL.");
                return 1;
            }

            Console.WriteLine("Connected successfully.");

            // Run the import
            Console.WriteLine();
            Console.WriteLine("Starting MPP import...");
            try
            {
                var traceService = new ConsoleTracingService();
                var importService = new MppProjectImportService(service, traceService);
                ImportResult result = importService.Execute(caseTemplateId, projectId);

                Console.WriteLine();
                Console.WriteLine("========================================");
                Console.WriteLine("  Import Complete");
                Console.WriteLine("  Tasks Created:  {0}", result.TasksCreated);
                Console.WriteLine("  Tasks Updated:  {0}", result.TasksUpdated);
                Console.WriteLine("  Total:          {0}", result.TotalProcessed);
                if (result.Warnings.Count > 0)
                {
                    Console.WriteLine("  Warnings:");
                    foreach (var w in result.Warnings)
                        Console.WriteLine("    - {0}", w);
                }
                Console.WriteLine("========================================");
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine("=== IMPORT FAILED ===");
                Console.WriteLine(ex.Message);
                if (ex.InnerException != null)
                    Console.WriteLine("Inner: " + ex.InnerException.Message);
                return 2;
            }
        }

        static int RunCrmImportS2SFromArgs(string[] args)
        {
            // args: imports2s <crmUrl> <templateId> <projectId> <clientId> <clientSecret>
            string crmUrl = args[1].TrimEnd('/');
            if (!crmUrl.StartsWith("http", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("ERROR: CRM URL must start with https://");
                return 1;
            }

            Guid caseTemplateId;
            if (!Guid.TryParse(args[2], out caseTemplateId))
            {
                Console.WriteLine("ERROR: '{0}' is not a valid GUID for Case Template ID.", args[2]);
                return 1;
            }

            Guid projectId;
            if (!Guid.TryParse(args[3], out projectId))
            {
                Console.WriteLine("ERROR: '{0}' is not a valid GUID for Project ID.", args[3]);
                return 1;
            }

            string clientId = args[4];
            string clientSecret = args[5];

            Console.WriteLine();
            Console.WriteLine("  CRM URL:          {0}", crmUrl);
            Console.WriteLine("  Case Template:    {0}", caseTemplateId);
            Console.WriteLine("  Target Project:   {0}", projectId);
            Console.WriteLine("  Client ID:        {0}", clientId);
            Console.WriteLine("  Auth:             Client Secret (S2S)");

            return ExecuteCrmImportS2S(crmUrl, caseTemplateId, projectId, clientId, clientSecret);
        }

        static void RunCrmImportS2S()
        {
            Console.WriteLine();
            Console.WriteLine("--- Dynamics 365 MPP Import (Client Secret / S2S) ---");
            Console.WriteLine();

            string crmUrl = PromptRequired("Dynamics 365 URL (e.g. https://yourorg.crm6.dynamics.com): ");
            if (crmUrl == null) return;
            crmUrl = crmUrl.TrimEnd('/');
            if (!crmUrl.StartsWith("http", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("ERROR: URL must start with https://");
                return;
            }

            Guid caseTemplateId = PromptGuid("adc_adccasetemplate ID (GUID): ");
            if (caseTemplateId == Guid.Empty) return;

            Guid projectId = PromptGuid("msdyn_project ID (GUID): ");
            if (projectId == Guid.Empty) return;

            string clientId = PromptRequired("Azure App (Client) ID: ");
            if (clientId == null) return;

            string clientSecret = PromptRequired("Client Secret: ");
            if (clientSecret == null) return;

            Console.WriteLine();
            Console.WriteLine("  CRM URL:          {0}", crmUrl);
            Console.WriteLine("  Case Template:    {0}", caseTemplateId);
            Console.WriteLine("  Target Project:   {0}", projectId);
            Console.WriteLine("  Client ID:        {0}", clientId);
            Console.WriteLine("  Auth:             Client Secret (S2S)");
            Console.WriteLine();
            Console.Write("Proceed? (Y/N): ");
            string confirm = Console.ReadLine()?.Trim().ToUpper();
            if (confirm != "Y") { Console.WriteLine("Cancelled."); return; }

            ExecuteCrmImportS2S(crmUrl, caseTemplateId, projectId, clientId, clientSecret);
        }

        static int ExecuteCrmImportS2S(string crmUrl, Guid caseTemplateId, Guid projectId, string clientId, string clientSecret)
        {
            Console.WriteLine();
            Console.WriteLine("Connecting to Dynamics 365 (client secret)...");

            IOrganizationService service = null;
            try
            {
                service = ConnectToCrmS2S(crmUrl, clientId, clientSecret);
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Failed to connect to CRM.");
                Console.WriteLine(ex.Message);
                return 1;
            }

            if (service == null)
            {
                Console.WriteLine("ERROR: CRM connection returned null. Check your credentials and URL.");
                return 1;
            }

            Console.WriteLine("Connected successfully.");

            Console.WriteLine();
            Console.WriteLine("Starting MPP import...");
            try
            {
                var traceService = new ConsoleTracingService();
                var importService = new MppProjectImportService(service, traceService);
                ImportResult result = importService.Execute(caseTemplateId, projectId);

                Console.WriteLine();
                Console.WriteLine("========================================");
                Console.WriteLine("  Import Complete");
                Console.WriteLine("  Tasks Created:  {0}", result.TasksCreated);
                Console.WriteLine("  Tasks Updated:  {0}", result.TasksUpdated);
                Console.WriteLine("  Total:          {0}", result.TotalProcessed);
                if (result.Warnings.Count > 0)
                {
                    Console.WriteLine("  Warnings:");
                    foreach (var w in result.Warnings)
                        Console.WriteLine("    - {0}", w);
                }
                Console.WriteLine("========================================");
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine("=== IMPORT FAILED ===");
                Console.WriteLine(ex.Message);
                if (ex.InnerException != null)
                    Console.WriteLine("Inner: " + ex.InnerException.Message);
                return 2;
            }
        }

        static IOrganizationService ConnectToCrmS2S(string crmUrl, string clientId, string clientSecret)
        {
            string connString = string.Format(
                "AuthType=ClientSecret;Url={0};ClientId={1};ClientSecret={2};RequireNewInstance=True",
                crmUrl, clientId, clientSecret);

            var client = new CrmServiceClient(connString);

            if (!client.IsReady)
            {
                string error = client.LastCrmError;
                Exception lastEx = client.LastCrmException;
                throw new Exception(
                    string.Format("CrmServiceClient not ready: {0}",
                        !string.IsNullOrEmpty(error) ? error : lastEx?.Message ?? "Unknown error"),
                    lastEx);
            }

            Console.WriteLine("  Org: {0}", client.ConnectedOrgFriendlyName);

            return (IOrganizationService)client;
        }

        static IOrganizationService ConnectToCrm(string crmUrl, string appId, string redirectUri)
        {
            string connString = string.Format(
                "AuthType=OAuth;Url={0};AppId={1};RedirectUri={2};LoginPrompt=Always;RequireNewInstance=True",
                crmUrl, appId, redirectUri);

            var client = new CrmServiceClient(connString);

            if (!client.IsReady)
            {
                string error = client.LastCrmError;
                Exception lastEx = client.LastCrmException;
                throw new Exception(
                    string.Format("CrmServiceClient not ready: {0}",
                        !string.IsNullOrEmpty(error) ? error : lastEx?.Message ?? "Unknown error"),
                    lastEx);
            }

            Console.WriteLine("  Org: {0}", client.ConnectedOrgFriendlyName);
            Console.WriteLine("  User: {0}", client.OAuthUserId);

            return (IOrganizationService)client;
        }

        #endregion

        #region .env Import

        static void RunCrmImportFromEnv()
        {
            // Look for .env next to the exe, then in the project directory
            string exeDir = AppDomain.CurrentDomain.BaseDirectory;
            string envPath = FindEnvFile(exeDir);

            if (envPath == null)
            {
                Console.WriteLine();
                Console.WriteLine("ERROR: .env file not found.");
                Console.WriteLine("Searched in: {0}", exeDir);
                Console.WriteLine("Copy .env.example to .env and fill in your values.");
                return;
            }

            Console.WriteLine();
            Console.WriteLine("Loading settings from: {0}", envPath);
            var env = ParseEnvFile(envPath);

            string crmUrl = GetEnv(env, "CRM_URL");
            if (crmUrl == null) { Console.WriteLine("ERROR: CRM_URL not set in .env"); return; }
            crmUrl = crmUrl.TrimEnd('/');

            Guid caseTemplateId = GetEnvGuid(env, "CASE_TEMPLATE_ID");
            if (caseTemplateId == Guid.Empty) return;

            Guid projectId = GetEnvGuid(env, "PROJECT_ID");
            if (projectId == Guid.Empty) return;

            string authType = GetEnv(env, "AUTH_TYPE") ?? "OAuth";

            Console.WriteLine("  CRM URL:          {0}", crmUrl);
            Console.WriteLine("  Case Template:    {0}", caseTemplateId);
            Console.WriteLine("  Target Project:   {0}", projectId);
            Console.WriteLine("  Auth Type:        {0}", authType);
            Console.WriteLine();
            Console.Write("Proceed? (Y/N): ");
            string confirm = Console.ReadLine()?.Trim().ToUpper();
            if (confirm != "Y") { Console.WriteLine("Cancelled."); return; }

            if (authType.Equals("ClientSecret", StringComparison.OrdinalIgnoreCase))
            {
                string clientId = GetEnv(env, "CLIENT_ID");
                string clientSecret = GetEnv(env, "CLIENT_SECRET");
                if (clientId == null) { Console.WriteLine("ERROR: CLIENT_ID not set in .env"); return; }
                if (clientSecret == null) { Console.WriteLine("ERROR: CLIENT_SECRET not set in .env"); return; }
                ExecuteCrmImportS2S(crmUrl, caseTemplateId, projectId, clientId, clientSecret);
            }
            else
            {
                string appId = GetEnv(env, "APP_ID") ?? DefaultAppId;
                string redirectUri = (appId == DefaultAppId) ? DefaultRedirectUri : "http://localhost";
                ExecuteCrmImport(crmUrl, caseTemplateId, projectId, appId, redirectUri);
            }
        }

        static string FindEnvFile(string startDir)
        {
            // Check exe directory, then walk up to find .env
            string dir = startDir;
            for (int i = 0; i < 5; i++)
            {
                string candidate = Path.Combine(dir, ".env");
                if (File.Exists(candidate)) return candidate;
                string parent = Directory.GetParent(dir)?.FullName;
                if (parent == null || parent == dir) break;
                dir = parent;
            }
            return null;
        }

        static Dictionary<string, string> ParseEnvFile(string path)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string rawLine in File.ReadAllLines(path))
            {
                string line = rawLine.Trim();
                if (string.IsNullOrEmpty(line) || line.StartsWith("#")) continue;
                int eq = line.IndexOf('=');
                if (eq <= 0) continue;
                string key = line.Substring(0, eq).Trim();
                string value = line.Substring(eq + 1).Trim();
                // Strip surrounding quotes if present
                if (value.Length >= 2 && ((value.StartsWith("\"") && value.EndsWith("\"")) || (value.StartsWith("'") && value.EndsWith("'"))))
                    value = value.Substring(1, value.Length - 2);
                dict[key] = value;
            }
            return dict;
        }

        static string GetEnv(Dictionary<string, string> env, string key)
        {
            string value;
            if (env.TryGetValue(key, out value) && !string.IsNullOrWhiteSpace(value))
                return value;
            return null;
        }

        static Guid GetEnvGuid(Dictionary<string, string> env, string key)
        {
            string raw = GetEnv(env, key);
            if (raw == null)
            {
                Console.WriteLine("ERROR: {0} not set in .env", key);
                return Guid.Empty;
            }
            Guid result;
            if (!Guid.TryParse(raw, out result))
            {
                Console.WriteLine("ERROR: {0} value '{1}' is not a valid GUID.", key, raw);
                return Guid.Empty;
            }
            return result;
        }

        #endregion

        #region Clear Project Tasks

        static void RunClearProjectTasks()
        {
            var envData = LoadEnvAndConnect(requireCaseTemplate: false);
            if (envData == null) return;

            // Confirmation
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine();
            Console.Write("  Type 'DELETE' to confirm clearing all tasks for project {0}: ", envData.Value.projectId);
            Console.ResetColor();
            string confirm = Console.ReadLine()?.Trim();
            if (confirm != "DELETE")
            {
                Console.WriteLine("  Cancelled.");
                return;
            }

            DoClearProjectTasks(envData.Value.service, envData.Value.projectId);
        }

        static void RunClearAndImport()
        {
            var envData = LoadEnvAndConnect(requireCaseTemplate: true);
            if (envData == null) return;

            Console.WriteLine();
            Console.Write("This will CLEAR all tasks then IMPORT from MPP. Continue? (Y/N): ");
            string confirm = Console.ReadLine()?.Trim().ToUpper();
            if (confirm != "Y") { Console.WriteLine("Cancelled."); return; }

            // Step 1: Clear existing tasks
            DoClearProjectTasks(envData.Value.service, envData.Value.projectId);

            // Step 2: Wait a moment for PSS to finish processing deletes
            Console.WriteLine();
            Console.WriteLine("Waiting for clear to complete...");
            System.Threading.Thread.Sleep(5000);

            // Step 3: Import
            Console.WriteLine("Starting import...");
            var trace = new ConsoleTracingService();
            var importService = new MppProjectImportService(envData.Value.service, trace);
            try
            {
                var result = importService.Execute(envData.Value.caseTemplateId, envData.Value.projectId);
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine();
                Console.WriteLine("=== IMPORT COMPLETE ===");
                Console.WriteLine("  Tasks created:  {0}", result.TasksCreated);
                Console.WriteLine("  Tasks updated:  {0}", result.TasksUpdated);
                Console.WriteLine("  Dependencies:   {0}", result.DependenciesCreated);
                Console.ResetColor();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine();
                Console.WriteLine("=== IMPORT FAILED ===");
                Console.WriteLine(ex.Message);
                Console.ResetColor();
            }
        }

        /// <summary>
        /// Loads .env, connects to CRM, and returns the connection info.
        /// Returns null if any required value is missing or connection fails.
        /// </summary>
        static (IOrganizationService service, Guid projectId, Guid caseTemplateId)?
            LoadEnvAndConnect(bool requireCaseTemplate)
        {
            string exeDir = AppDomain.CurrentDomain.BaseDirectory;
            string envPath = FindEnvFile(exeDir);
            if (envPath == null)
            {
                Console.WriteLine();
                Console.WriteLine("ERROR: .env file not found.");
                Console.WriteLine("Copy .env.example to .env and fill in your values.");
                return null;
            }

            var env = ParseEnvFile(envPath);

            string crmUrl = GetEnv(env, "CRM_URL");
            if (crmUrl == null) { Console.WriteLine("ERROR: CRM_URL not set in .env"); return null; }
            crmUrl = crmUrl.TrimEnd('/');

            Guid projectId = GetEnvGuid(env, "PROJECT_ID");
            if (projectId == Guid.Empty) return null;

            Guid caseTemplateId = Guid.Empty;
            if (requireCaseTemplate)
            {
                caseTemplateId = GetEnvGuid(env, "CASE_TEMPLATE_ID");
                if (caseTemplateId == Guid.Empty) return null;
            }

            // Connect to CRM
            IOrganizationService service = null;
            string authType = GetEnv(env, "AUTH_TYPE") ?? "OAuth";
            try
            {
                if (authType.Equals("ClientSecret", StringComparison.OrdinalIgnoreCase))
                {
                    string clientId = GetEnv(env, "CLIENT_ID");
                    string clientSecret = GetEnv(env, "CLIENT_SECRET");
                    if (clientId == null || clientSecret == null)
                    {
                        Console.WriteLine("ERROR: CLIENT_ID and CLIENT_SECRET required in .env for ClientSecret auth.");
                        return null;
                    }
                    Console.WriteLine("Connecting to Dynamics 365 (client secret)...");
                    service = ConnectToCrmS2S(crmUrl, clientId, clientSecret);
                }
                else
                {
                    string appId = GetEnv(env, "APP_ID") ?? DefaultAppId;
                    string redirectUri = (appId == DefaultAppId) ? DefaultRedirectUri : "http://localhost";
                    Console.WriteLine("Connecting to Dynamics 365 (browser login)...");
                    service = ConnectToCrm(crmUrl, appId, redirectUri);
                }
                Console.WriteLine("  Connected successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: Failed to connect: {0}", ex.Message);
                return null;
            }

            return (service, projectId, caseTemplateId);
        }

        /// <summary>
        /// Deletes all tasks and dependencies for the given project via PSS.
        /// Can be called standalone or as part of clear+import.
        /// </summary>
        static void DoClearProjectTasks(IOrganizationService service, Guid projectId)
        {
            Console.WriteLine();
            Console.WriteLine("Querying project tasks for project {0}...", projectId);

            var taskQuery = new Microsoft.Xrm.Sdk.Query.QueryExpression("msdyn_projecttask")
            {
                ColumnSet = new Microsoft.Xrm.Sdk.Query.ColumnSet("msdyn_subject", "msdyn_outlinelevel", "msdyn_parenttask"),
                Criteria = new Microsoft.Xrm.Sdk.Query.FilterExpression
                {
                    Conditions =
                    {
                        new Microsoft.Xrm.Sdk.Query.ConditionExpression("msdyn_project", Microsoft.Xrm.Sdk.Query.ConditionOperator.Equal, projectId)
                    }
                }
            };
            taskQuery.PageInfo = new Microsoft.Xrm.Sdk.Query.PagingInfo { PageNumber = 1, Count = 5000 };

            var allTasks = new System.Collections.Generic.List<Entity>();
            while (true)
            {
                var resp = service.RetrieveMultiple(taskQuery);
                allTasks.AddRange(resp.Entities);
                if (resp.MoreRecords) { taskQuery.PageInfo.PageNumber++; taskQuery.PageInfo.PagingCookie = resp.PagingCookie; }
                else break;
            }

            // Sort deepest tasks first so children are deleted before parents
            allTasks.Sort((a, b) =>
            {
                int levelA = a.Contains("msdyn_outlinelevel") ? (int)a["msdyn_outlinelevel"] : 0;
                int levelB = b.Contains("msdyn_outlinelevel") ? (int)b["msdyn_outlinelevel"] : 0;
                return levelB.CompareTo(levelA);
            });

            var depQuery = new Microsoft.Xrm.Sdk.Query.QueryExpression("msdyn_projecttaskdependency")
            {
                ColumnSet = new Microsoft.Xrm.Sdk.Query.ColumnSet("msdyn_projecttaskdependencyid"),
                Criteria = new Microsoft.Xrm.Sdk.Query.FilterExpression
                {
                    Conditions =
                    {
                        new Microsoft.Xrm.Sdk.Query.ConditionExpression("msdyn_project", Microsoft.Xrm.Sdk.Query.ConditionOperator.Equal, projectId)
                    }
                }
            };
            depQuery.PageInfo = new Microsoft.Xrm.Sdk.Query.PagingInfo { PageNumber = 1, Count = 5000 };

            var allDeps = new System.Collections.Generic.List<Entity>();
            while (true)
            {
                var resp = service.RetrieveMultiple(depQuery);
                allDeps.AddRange(resp.Entities);
                if (resp.MoreRecords) { depQuery.PageInfo.PageNumber++; depQuery.PageInfo.PagingCookie = resp.PagingCookie; }
                else break;
            }

            Console.WriteLine("  Found {0} project task(s) and {1} dependency record(s).", allTasks.Count, allDeps.Count);

            if (allTasks.Count == 0 && allDeps.Count == 0)
            {
                Console.WriteLine("  Nothing to delete.");
                return;
            }

            // Delete dependencies first via PSS OperationSet
            Console.WriteLine();
            Console.WriteLine("Creating OperationSet for deletion...");
            var createOsReq = new OrganizationRequest("msdyn_CreateOperationSetV1");
            createOsReq["ProjectId"] = projectId.ToString();
            createOsReq["Description"] = "Clear project tasks";
            var createOsResp = service.Execute(createOsReq);
            string operationSetId = (string)createOsResp["OperationSetId"];
            Console.WriteLine("  OperationSet: {0}", operationSetId);

            // Queue dependency deletes
            int depCount = 0;
            foreach (var dep in allDeps)
            {
                var delReq = new OrganizationRequest("msdyn_PssDeleteV1");
                delReq["RecordId"] = dep.Id.ToString();
                delReq["EntityLogicalName"] = dep.LogicalName;
                delReq["OperationSetId"] = operationSetId;
                service.Execute(delReq);
                depCount++;
            }
            Console.WriteLine("  Queued {0} dependency deletes.", depCount);

            // Queue task deletes (sorted deepest-first)
            int taskCount = 0;
            foreach (var task in allTasks)
            {
                int level = task.Contains("msdyn_outlinelevel") ? (int)task["msdyn_outlinelevel"] : -1;
                string name = task.Contains("msdyn_subject") ? (string)task["msdyn_subject"] : "(no name)";
                Console.WriteLine("    [{0}] Lvl {1}: {2} ({3})", taskCount, level, name, task.Id);

                var delReq = new OrganizationRequest("msdyn_PssDeleteV1");
                delReq["RecordId"] = task.Id.ToString();
                delReq["EntityLogicalName"] = task.LogicalName;
                delReq["OperationSetId"] = operationSetId;
                service.Execute(delReq);
                taskCount++;
            }
            Console.WriteLine("  Queued {0} task deletes.", taskCount);

            // Execute
            Console.WriteLine("  Executing OperationSet...");
            try
            {
                var execReq = new OrganizationRequest("msdyn_ExecuteOperationSetV1");
                execReq["OperationSetId"] = operationSetId;
                service.Execute(execReq);
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("  Done. Deleted {0} dependencies and {1} tasks.", depCount, taskCount);
                Console.ResetColor();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("  Delete failed: {0}", ex.Message);
                Console.ResetColor();
            }
        }

        #endregion

        #region Helpers

        static string PromptRequired(string prompt)
        {
            Console.Write(prompt);
            string value = Console.ReadLine()?.Trim();
            if (string.IsNullOrEmpty(value))
            {
                Console.WriteLine("ERROR: Value is required.");
                return null;
            }
            return value;
        }

        static Guid PromptGuid(string prompt)
        {
            Console.Write(prompt);
            string input = Console.ReadLine()?.Trim();
            if (string.IsNullOrEmpty(input))
            {
                Console.WriteLine("ERROR: GUID is required.");
                return Guid.Empty;
            }

            Guid result;
            if (!Guid.TryParse(input, out result))
            {
                Console.WriteLine("ERROR: '{0}' is not a valid GUID.", input);
                return Guid.Empty;
            }
            return result;
        }

        static string Trunc(string s, int max)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Length <= max ? s : s.Substring(0, max - 3) + "...";
        }

        #endregion
    }

    /// <summary>
    /// ITracingService implementation that writes to Console.
    /// </summary>
    class ConsoleTracingService : ITracingService
    {
        public void Trace(string format, params object[] args)
        {
            Console.WriteLine("  [TRACE] " + format, args);
        }
    }
}
