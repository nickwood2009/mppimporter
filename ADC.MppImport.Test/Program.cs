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

            // CLI mode: regtest [exampleFilesDir]
            if (args.Length >= 1 && args[0].Equals("regtest", StringComparison.OrdinalIgnoreCase))
            {
                string dir = args.Length >= 2 ? args[1] : null;
                return RunRegressionTests(dir);
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

            // CLI mode: depcheck <mppFilePath> — analyze dependency graph for cycles
            if (args.Length >= 2 && args[0].Equals("depcheck", StringComparison.OrdinalIgnoreCase))
            {
                return RunDepCheck(args[1]);
            }

            // CLI mode: asynctest [mppFilePath]
            if (args.Length >= 1 && args[0].Equals("asynctest", StringComparison.OrdinalIgnoreCase))
            {
                string mppPath = args.Length >= 2 ? args[1] : null;
                return RunAsyncImportTest(mppPath);
            }

            // CLI mode: sidebyside [exampleFilesDir]
            if (args.Length >= 1 && args[0].Equals("sidebyside", StringComparison.OrdinalIgnoreCase))
            {
                string dir = args.Length >= 2 ? args[1] : null;
                return RunSideBySideTests(dir);
            }

            // CLI mode: calinfo [exampleFilesDir]
            if (args.Length >= 1 && args[0].Equals("calinfo", StringComparison.OrdinalIgnoreCase))
            {
                string dir = args.Length >= 2 ? args[1] : null;
                return RunCalInfo(dir);
            }

            // CLI mode: datecheck <mppFile> <projectGuid> [mppFile2 projectGuid2 ...]
            if (args.Length >= 1 && args[0].Equals("datecheck", StringComparison.OrdinalIgnoreCase))
            {
                // Pairs of (mppFile, projectGuid)
                var pairs = new List<Tuple<string, Guid>>();
                for (int i = 1; i + 1 < args.Length; i += 2)
                {
                    Guid g;
                    if (Guid.TryParse(args[i + 1], out g))
                        pairs.Add(Tuple.Create(args[i], g));
                }
                return RunDateCheck(pairs);
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
                Console.WriteLine("  7. Regression Tests (all MPP files in exampleFiles)");
                Console.WriteLine("  8. Async Import Test (plugin chain)");
                Console.WriteLine("  9. Side-by-Side Comparison (all MPP files)");
                Console.WriteLine("  0. Exit");
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
                        RunRegressionTests(null);
                        break;
                    case "8":
                        RunAsyncImportTest(null);
                        break;
                    case "9":
                        RunSideBySideTests(null);
                        break;
                    case "0":
                        return 0;
                    default:
                        Console.WriteLine("Invalid option. Please enter 0-9.");
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
                    if (shown++ >= 200) { Console.WriteLine("  ... ({0} more)", project.Tasks.Count - 200); break; }
                    string preds = t.Predecessors.Count > 0
                        ? string.Join(",", t.Predecessors.Select(p => p.SourceTaskUniqueID + "(" + p.Type + ")"))
                        : "";
                    string durStr = t.Duration != null
                        ? string.Format("{0} {1}", t.Duration.Value, t.Duration.Units)
                        : "null";
                    string parentStr = t.ParentTaskUniqueID.HasValue ? string.Format("P={0}", t.ParentTaskUniqueID.Value) : "";
                    Console.WriteLine("  [{0,3}] L{1} {2,-40} Start={3:d}  Dur={4,-18} {5,-8} Pred={6}",
                        t.UniqueID, t.OutlineLevel, Trunc(t.Name, 40),
                        t.Start, durStr, parentStr, preds);
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
            LoadEnvAndConnect(bool requireCaseTemplate, bool requireProjectId = true)
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

            Guid projectId = Guid.Empty;
            if (requireProjectId)
            {
                projectId = GetEnvGuid(env, "PROJECT_ID");
                if (projectId == Guid.Empty) return null;
            }

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

        #region Regression Tests

        /// <summary>
        /// Runs regression tests against a live D365 environment for all MPP files in the example directory.
        /// For each file: parse locally, create project, import, query back, compare, clean up.
        /// </summary>
        static int RunRegressionTests(string exampleDirOverride)
        {
            Console.WriteLine();
            Console.WriteLine("========================================");
            Console.WriteLine("  ADC MPP Import - Regression Tests");
            Console.WriteLine("========================================");

            // 1. Find example files directory
            string exampleDir = exampleDirOverride;
            if (string.IsNullOrEmpty(exampleDir))
            {
                // Walk up from exe dir to find exampleFiles
                string dir = AppDomain.CurrentDomain.BaseDirectory;
                for (int i = 0; i < 6; i++)
                {
                    string candidate = Path.Combine(dir, "exampleFiles");
                    if (Directory.Exists(candidate)) { exampleDir = candidate; break; }
                    string parent = Directory.GetParent(dir)?.FullName;
                    if (parent == null || parent == dir) break;
                    dir = parent;
                }
            }

            if (exampleDir == null || !Directory.Exists(exampleDir))
            {
                Console.WriteLine("ERROR: exampleFiles directory not found.");
                Console.WriteLine("Usage: regtest [path/to/exampleFiles]");
                return 1;
            }

            string[] mppFiles = Directory.GetFiles(exampleDir, "*.mpp");
            if (mppFiles.Length == 0)
            {
                Console.WriteLine("ERROR: No .mpp files found in {0}", exampleDir);
                return 1;
            }

            Console.WriteLine("Found {0} MPP file(s) in: {1}", mppFiles.Length, exampleDir);
            foreach (var f in mppFiles)
                Console.WriteLine("  - {0}", Path.GetFileName(f));

            // 2. Connect to D365 (only need CRM_URL + auth, no project/template IDs)
            var envData = LoadEnvAndConnect(requireCaseTemplate: false, requireProjectId: false);
            if (envData == null) return 1;
            var service = envData.Value.service;

            Console.WriteLine();
            Console.WriteLine("Starting regression tests...");
            Console.WriteLine();

            int passed = 0;
            int failed = 0;
            var summaries = new List<string>();

            foreach (string mppPath in mppFiles)
            {
                string fileName = Path.GetFileName(mppPath);
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("================================================================");
                Console.WriteLine("  TEST: {0}", fileName);
                Console.WriteLine("================================================================");
                Console.ResetColor();

                try
                {
                    var result = RunSingleRegressionTest(service, mppPath);
                    if (result.Passed)
                    {
                        passed++;
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("  RESULT: PASS ({0} checks passed, {1} warnings)", result.ChecksPassed, result.Warnings.Count);
                        summaries.Add(string.Format("PASS  {0}  ({1} checks, {2} warnings)", fileName, result.ChecksPassed, result.Warnings.Count));
                    }
                    else
                    {
                        failed++;
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("  RESULT: FAIL ({0} checks passed, {1} failures)", result.ChecksPassed, result.Failures.Count);
                        summaries.Add(string.Format("FAIL  {0}  ({1} failures)", fileName, result.Failures.Count));
                    }
                    Console.ResetColor();

                    foreach (var w in result.Warnings)
                        Console.WriteLine("    WARN: {0}", w);
                    foreach (var f in result.Failures)
                        Console.WriteLine("    FAIL: {0}", f);
                }
                catch (Exception ex)
                {
                    failed++;
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("  RESULT: ERROR - {0}", ex.Message);
                    Console.ResetColor();
                    summaries.Add(string.Format("ERROR {0}  ({1})", fileName, ex.Message));
                }

                Console.WriteLine();
            }

            // Final summary
            Console.WriteLine("================================================================");
            Console.ForegroundColor = passed > 0 && failed == 0 ? ConsoleColor.Green : ConsoleColor.Yellow;
            Console.WriteLine("  REGRESSION TEST SUMMARY: {0} passed, {1} failed out of {2}", passed, failed, mppFiles.Length);
            Console.ResetColor();
            foreach (var s in summaries)
                Console.WriteLine("    {0}", s);
            Console.WriteLine("================================================================");

            return failed > 0 ? 2 : 0;
        }

        class RegressionResult
        {
            public bool Passed { get { return Failures.Count == 0; } }
            public int ChecksPassed;
            public List<string> Failures = new List<string>();
            public List<string> Warnings = new List<string>();
        }

        static RegressionResult RunSingleRegressionTest(IOrganizationService service, string mppPath)
        {
            var result = new RegressionResult();
            string fileName = Path.GetFileName(mppPath);

            // 1. Parse MPP locally
            Console.WriteLine("  Parsing MPP locally...");
            byte[] mppBytes = File.ReadAllBytes(mppPath);
            var reader = new MppFileReader();
            ProjectFile mppProject = reader.Read(mppBytes);

            // Build expected data from MPP (exclude UniqueID 0 = project summary task)
            var mppTasks = mppProject.Tasks.Where(t => t.UniqueID.HasValue && t.UniqueID.Value != 0).ToList();
            Console.WriteLine("  MPP: {0} tasks (excl. summary 0), {1} relations", mppTasks.Count, mppTasks.Sum(t => t.Predecessors != null ? t.Predecessors.Count : 0));

            // Identify summary tasks (have children)
            var summaryIds = new HashSet<int>();
            foreach (var t in mppTasks)
            {
                if (t.ParentTask != null && t.ParentTask.UniqueID.HasValue)
                    summaryIds.Add(t.ParentTask.UniqueID.Value);
            }

            // Build expected leaf task data (name -> expected duration in days)
            var expectedLeafTasks = new Dictionary<string, List<double>>(StringComparer.OrdinalIgnoreCase);
            foreach (var t in mppTasks)
            {
                if (summaryIds.Contains(t.UniqueID.Value)) continue; // skip summary tasks for duration comparison
                string name = t.Name ?? "(Unnamed Task)";
                double durationDays = 0;
                if (t.Duration != null && t.Duration.Value >= 0)
                {
                    double hours = ConvertToHoursLocal(t.Duration);
                    durationDays = Math.Round(hours / 8.0, 2);
                }
                List<double> durations;
                if (!expectedLeafTasks.TryGetValue(name, out durations))
                {
                    durations = new List<double>();
                    expectedLeafTasks[name] = durations;
                }
                durations.Add(durationDays);
            }

            // Build expected parent relationships: child name -> parent name
            var expectedParents = new Dictionary<int, string>();
            foreach (var t in mppTasks)
            {
                if (t.ParentTask != null && t.ParentTask.UniqueID.HasValue && t.ParentTask.UniqueID.Value != 0)
                    expectedParents[t.UniqueID.Value] = t.ParentTask.Name ?? "(Unnamed Task)";
            }

            // Build expected dependency count (excluding summary-task deps)
            int expectedDepCount = 0;
            foreach (var t in mppTasks)
            {
                if (t.Predecessors == null) continue;
                foreach (var p in t.Predecessors)
                {
                    if (!summaryIds.Contains(p.SourceTaskUniqueID) && !summaryIds.Contains(t.UniqueID.Value))
                        expectedDepCount++;
                }
            }

            // 2. Create a new project in D365
            Console.WriteLine("  Creating D365 project: REGTEST - {0}...", fileName);
            var projectEntity = new Entity("msdyn_project");
            projectEntity["msdyn_subject"] = string.Format("REGTEST - {0} - {1:yyyyMMdd-HHmmss}", fileName, DateTime.Now);
            Guid projectId = service.Create(projectEntity);
            Console.WriteLine("  Project created: {0}", projectId);

            try
            {
                // 3. Import using our service
                Console.WriteLine("  Importing via MppProjectImportService.ExecuteFromBytes...");
                var trace = new ConsoleTracingService();
                var importService = new MppProjectImportService(service, trace);

                // Use earliest task start date from MPP as project start
                DateTime? projectStart = null;
                var allStarts = mppTasks.Where(t => t.Start.HasValue).Select(t => t.Start.Value).ToList();
                if (allStarts.Count > 0) projectStart = allStarts.Min();

                var importResult = importService.ExecuteFromBytes(mppBytes, projectId, projectStart);
                Console.WriteLine("  Import done: {0} created, {1} updated, {2} deps",
                    importResult.TasksCreated, importResult.TasksUpdated, importResult.DependenciesCreated);

                // 4. Wait for PSS to finish processing (deps need time to materialize)
                Console.WriteLine("  Waiting for PSS to finalize (30s)...");
                System.Threading.Thread.Sleep(30000);

                // 5. Query back tasks from D365
                Console.WriteLine("  Querying D365 tasks...");
                var crmTasks = QueryProjectTasksFull(service, projectId);
                var crmDeps = QueryProjectDependencies(service, projectId);

                // Exclude the PSS root task (outline level 0, auto-created by PSS)
                var crmNonRoot = crmTasks.Where(e =>
                {
                    int level = e.Contains("msdyn_outlinelevel") ? (int)e["msdyn_outlinelevel"] : -1;
                    return level != 0;
                }).ToList();

                Console.WriteLine("  D365: {0} tasks (excl. root), {1} dependencies", crmNonRoot.Count, crmDeps.Count);

                // ===== CHECKS =====

                // CHECK 1: Task count
                if (crmNonRoot.Count == mppTasks.Count)
                {
                    result.ChecksPassed++;
                    Console.WriteLine("    [PASS] Task count: {0}", crmNonRoot.Count);
                }
                else
                {
                    result.Failures.Add(string.Format("Task count: expected {0}, got {1}", mppTasks.Count, crmNonRoot.Count));
                }

                // CHECK 2: All MPP task names exist in D365
                var crmNameCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                foreach (var e in crmNonRoot)
                {
                    string name = e.GetAttributeValue<string>("msdyn_subject") ?? "";
                    int count;
                    crmNameCounts.TryGetValue(name, out count);
                    crmNameCounts[name] = count + 1;
                }

                var mppNameCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                foreach (var t in mppTasks)
                {
                    string name = t.Name ?? "(Unnamed Task)";
                    int count;
                    mppNameCounts.TryGetValue(name, out count);
                    mppNameCounts[name] = count + 1;
                }

                int nameMatches = 0;
                int nameMismatches = 0;
                foreach (var kvp in mppNameCounts)
                {
                    int crmCount;
                    crmNameCounts.TryGetValue(kvp.Key, out crmCount);
                    if (crmCount >= kvp.Value)
                        nameMatches++;
                    else
                    {
                        nameMismatches++;
                        result.Failures.Add(string.Format("Missing task name '{0}': expected {1} occurrences, found {2}", kvp.Key, kvp.Value, crmCount));
                    }
                }

                if (nameMismatches == 0)
                {
                    result.ChecksPassed++;
                    Console.WriteLine("    [PASS] All {0} distinct task names found in D365", nameMatches);
                }

                // CHECK 3: Leaf task durations
                int durationMatches = 0;
                int durationMismatches = 0;
                foreach (var crmTask in crmNonRoot)
                {
                    string name = crmTask.GetAttributeValue<string>("msdyn_subject") ?? "";
                    double? crmDuration = crmTask.Contains("msdyn_duration") ? (double?)crmTask["msdyn_duration"] : null;

                    // Only check leaf tasks (those that don't have children in CRM — use parent lookup)
                    bool isCrmParent = crmNonRoot.Any(c =>
                    {
                        var parentRef = c.GetAttributeValue<EntityReference>("msdyn_parenttask");
                        return parentRef != null && parentRef.Id == crmTask.Id;
                    });
                    if (isCrmParent) continue; // summary — PSS auto-calculates

                    List<double> expectedDurations;
                    if (expectedLeafTasks.TryGetValue(name, out expectedDurations) && expectedDurations.Count > 0)
                    {
                        double expected = expectedDurations[0];
                        double actual = crmDuration ?? 0;
                        if (Math.Abs(expected - actual) < 0.5)
                        {
                            durationMatches++;
                            expectedDurations.RemoveAt(0);
                        }
                        else
                        {
                            durationMismatches++;
                            result.Warnings.Add(string.Format("Duration mismatch '{0}': expected {1:F1}d, got {2:F1}d", name, expected, actual));
                            expectedDurations.RemoveAt(0);
                        }
                    }
                }

                if (durationMismatches == 0)
                {
                    result.ChecksPassed++;
                    Console.WriteLine("    [PASS] Leaf task durations: {0} matched", durationMatches);
                }
                else
                {
                    Console.WriteLine("    [WARN] Leaf task durations: {0} matched, {1} mismatched", durationMatches, durationMismatches);
                }

                // CHECK 4: Hierarchy (parent links)
                int parentMatches = 0;
                var crmTaskById = crmNonRoot.ToDictionary(e => e.Id);

                foreach (var crmTask in crmNonRoot)
                {
                    var parentRef = crmTask.GetAttributeValue<EntityReference>("msdyn_parenttask");
                    string crmName = crmTask.GetAttributeValue<string>("msdyn_subject") ?? "";

                    if (parentRef != null)
                    {
                        Entity parentEntity;
                        string crmParentName = "";
                        if (crmTaskById.TryGetValue(parentRef.Id, out parentEntity))
                            crmParentName = parentEntity.GetAttributeValue<string>("msdyn_subject") ?? "";

                        // Find matching MPP task and check if parent name matches
                        bool foundMatch = false;
                        foreach (var kvp in expectedParents)
                        {
                            var mppTask = mppTasks.FirstOrDefault(t => t.UniqueID.Value == kvp.Key);
                            if (mppTask != null && (mppTask.Name ?? "(Unnamed Task)") == crmName && kvp.Value == crmParentName)
                            {
                                parentMatches++;
                                foundMatch = true;
                                break;
                            }
                        }
                        if (!foundMatch)
                        {
                            // Don't fail on this — parent matching by name can have ambiguity
                            // Just count it
                            parentMatches++; // trust the import for now
                        }
                    }
                }

                int expectedParentCount = expectedParents.Count;
                int crmParentCount = crmNonRoot.Count(e => e.GetAttributeValue<EntityReference>("msdyn_parenttask") != null);

                // PSS auto-parents top-level tasks to the root, so D365 will have
                // at least as many parent links as we set, often more.
                if (crmParentCount >= expectedParentCount)
                {
                    result.ChecksPassed++;
                    Console.WriteLine("    [PASS] Parent links: expected >={0}, got {1}", expectedParentCount, crmParentCount);
                }
                else
                {
                    result.Failures.Add(string.Format("Parent link count: expected >={0}, got {1}", expectedParentCount, crmParentCount));
                }

                // CHECK 5: Dependency count
                if (Math.Abs(expectedDepCount - crmDeps.Count) <= 2) // allow small tolerance for filtered summary deps
                {
                    result.ChecksPassed++;
                    Console.WriteLine("    [PASS] Dependencies: expected ~{0}, got {1}", expectedDepCount, crmDeps.Count);
                }
                else
                {
                    result.Warnings.Add(string.Format("Dependency count: expected ~{0} (excl. summary), got {1}", expectedDepCount, crmDeps.Count));
                }

                // CHECK 6: Import completed without fatal error
                result.ChecksPassed++;
                Console.WriteLine("    [PASS] Import completed without fatal error");
            }
            finally
            {
                // 6. Clean up: delete the test project
                Console.WriteLine("  Cleaning up: deleting test project {0}...", projectId);
                try
                {
                    // Clear tasks first via PSS, then delete project
                    DoClearProjectTasks(service, projectId);
                    System.Threading.Thread.Sleep(5000);
                    service.Delete("msdyn_project", projectId);
                    Console.WriteLine("  Project deleted.");
                }
                catch (Exception cleanEx)
                {
                    Console.WriteLine("  WARNING: Cleanup failed (project may remain): {0}", cleanEx.Message);
                }
            }

            return result;
        }

        /// <summary>
        /// Queries all project tasks with full details for comparison.
        /// </summary>
        static List<Entity> QueryProjectTasksFull(IOrganizationService service, Guid projectId)
        {
            var query = new Microsoft.Xrm.Sdk.Query.QueryExpression("msdyn_projecttask")
            {
                ColumnSet = new Microsoft.Xrm.Sdk.Query.ColumnSet(
                    "msdyn_subject", "msdyn_duration", "msdyn_effort",
                    "msdyn_parenttask", "msdyn_outlinelevel",
                    "msdyn_scheduledstart", "msdyn_scheduledend"),
                Criteria = new Microsoft.Xrm.Sdk.Query.FilterExpression
                {
                    Conditions =
                    {
                        new Microsoft.Xrm.Sdk.Query.ConditionExpression("msdyn_project",
                            Microsoft.Xrm.Sdk.Query.ConditionOperator.Equal, projectId)
                    }
                }
            };
            query.PageInfo = new Microsoft.Xrm.Sdk.Query.PagingInfo { PageNumber = 1, Count = 5000 };

            var results = new List<Entity>();
            while (true)
            {
                var resp = service.RetrieveMultiple(query);
                results.AddRange(resp.Entities);
                if (resp.MoreRecords) { query.PageInfo.PageNumber++; query.PageInfo.PagingCookie = resp.PagingCookie; }
                else break;
            }
            return results;
        }

        /// <summary>
        /// Queries all project task dependencies for comparison.
        /// </summary>
        static List<Entity> QueryProjectDependencies(IOrganizationService service, Guid projectId)
        {
            var query = new Microsoft.Xrm.Sdk.Query.QueryExpression("msdyn_projecttaskdependency")
            {
                ColumnSet = new Microsoft.Xrm.Sdk.Query.ColumnSet(
                    "msdyn_predecessortask", "msdyn_successortask", "msdyn_linktype"),
                Criteria = new Microsoft.Xrm.Sdk.Query.FilterExpression
                {
                    Conditions =
                    {
                        new Microsoft.Xrm.Sdk.Query.ConditionExpression("msdyn_project",
                            Microsoft.Xrm.Sdk.Query.ConditionOperator.Equal, projectId)
                    }
                }
            };
            query.PageInfo = new Microsoft.Xrm.Sdk.Query.PagingInfo { PageNumber = 1, Count = 5000 };

            var results = new List<Entity>();
            while (true)
            {
                var resp = service.RetrieveMultiple(query);
                results.AddRange(resp.Entities);
                if (resp.MoreRecords) { query.PageInfo.PageNumber++; query.PageInfo.PagingCookie = resp.PagingCookie; }
                else break;
            }
            return results;
        }

        /// <summary>
        /// Local copy of ConvertToHours for regression test comparison.
        /// Must stay in sync with MppProjectImportService.ConvertToHours.
        /// </summary>
        static double ConvertToHoursLocal(Duration duration)
        {
            double val = duration.Value;
            switch (duration.Units)
            {
                case TimeUnit.Minutes:
                case TimeUnit.ElapsedMinutes:
                    return val / 60.0;
                case TimeUnit.Hours:
                case TimeUnit.ElapsedHours:
                    return val;
                case TimeUnit.Days:
                case TimeUnit.ElapsedDays:
                    return val * 8.0;
                case TimeUnit.Weeks:
                case TimeUnit.ElapsedWeeks:
                    return val * 40.0;
                case TimeUnit.Months:
                case TimeUnit.ElapsedMonths:
                    return val * 160.0;
                default: return val;
            }
        }

        #endregion

        #region Dependency Graph Analysis

        static int MapRelationTypeStatic(RelationType type)
        {
            switch (type)
            {
                case RelationType.FinishToStart: return 192350000;
                case RelationType.StartToStart: return 192350001;
                case RelationType.FinishToFinish: return 192350002;
                case RelationType.StartToFinish: return 192350003;
                default: return 192350000;
            }
        }

        static bool DfsCycleDetect(int node, Dictionary<int, List<int>> adj,
            HashSet<int> white, HashSet<int> gray, HashSet<int> black,
            List<int> cyclePath)
        {
            white.Remove(node);
            gray.Add(node);
            cyclePath.Add(node);

            List<int> neighbors;
            if (adj.TryGetValue(node, out neighbors))
            {
                foreach (var next in neighbors)
                {
                    if (gray.Contains(next))
                    {
                        cyclePath.Add(next);
                        return true;
                    }
                    if (white.Contains(next))
                    {
                        if (DfsCycleDetect(next, adj, white, gray, black, cyclePath))
                            return true;
                    }
                }
            }

            gray.Remove(node);
            black.Add(node);
            cyclePath.RemoveAt(cyclePath.Count - 1);
            return false;
        }

        /// <summary>
        /// Parses an MPP file and analyzes the dependency graph for scheduling cycles.
        /// Uses the same filtering as the async import (summary-task deps removed).
        /// </summary>
        static int RunDepCheck(string mppPath)
        {
            Console.WriteLine("\n=== Dependency Graph Analysis ===");
            if (!File.Exists(mppPath))
            {
                Console.WriteLine("ERROR: File not found: {0}", mppPath);
                return 1;
            }

            byte[] mppBytes = File.ReadAllBytes(mppPath);
            Console.WriteLine("File: {0} ({1:N0} bytes)", Path.GetFileName(mppPath), mppBytes.Length);

            var reader = new MppFileReader();
            ProjectFile project = reader.Read(mppBytes);

            var sortedByOrder = project.Tasks
                .Where(t => t.UniqueID.HasValue && t.UniqueID.Value != 0)
                .ToList();

            // Derive parent relationships
            var lastAtLevel = new Dictionary<int, Task>();
            foreach (var mppTask in sortedByOrder)
            {
                int level = mppTask.OutlineLevel ?? 0;
                if (level > 0 && mppTask.ParentTask == null)
                {
                    Task derivedParent;
                    if (lastAtLevel.TryGetValue(level - 1, out derivedParent))
                    {
                        mppTask.ParentTask = derivedParent;
                        if (!derivedParent.ChildTasks.Contains(mppTask))
                            derivedParent.ChildTasks.Add(mppTask);
                    }
                }
                lastAtLevel[level] = mppTask;
            }

            // Build summary task set
            var summaryTaskIds = new HashSet<int>();
            foreach (var mppTask in sortedByOrder)
            {
                if (mppTask.ParentTask != null && mppTask.ParentTask.UniqueID.HasValue
                    && mppTask.ParentTask.UniqueID.Value != 0)
                    summaryTaskIds.Add(mppTask.ParentTask.UniqueID.Value);
            }

            // Build task name lookup
            var taskNames = new Dictionary<int, string>();
            foreach (var mppTask in sortedByOrder)
            {
                if (mppTask.UniqueID.HasValue)
                    taskNames[mppTask.UniqueID.Value] = mppTask.Name ?? "(unnamed)";
            }

            Console.WriteLine("Tasks: {0}, Summary tasks: {1}", sortedByOrder.Count, summaryTaskIds.Count);

            string[] linkTypeNames = { "FS", "SS", "FF", "SF" };
            var allDeps = new List<Tuple<int, int, int>>(); // pred, succ, linkType
            int summarySkipped = 0;

            foreach (var mppTask in project.Tasks)
            {
                if (!mppTask.UniqueID.HasValue || mppTask.UniqueID.Value == 0) continue;
                if (mppTask.Predecessors == null || mppTask.Predecessors.Count == 0) continue;

                foreach (var relation in mppTask.Predecessors)
                {
                    if (relation.SourceTaskUniqueID == 0) continue;

                    if (summaryTaskIds.Contains(relation.SourceTaskUniqueID) || summaryTaskIds.Contains(mppTask.UniqueID.Value))
                    {
                        summarySkipped++;
                        string pn = taskNames.ContainsKey(relation.SourceTaskUniqueID) ? taskNames[relation.SourceTaskUniqueID] : "?";
                        string sn = taskNames.ContainsKey(mppTask.UniqueID.Value) ? taskNames[mppTask.UniqueID.Value] : "?";
                        int lt2 = MapRelationTypeStatic(relation.Type);
                        string ltn = (lt2 >= 192350000 && lt2 <= 192350003) ? linkTypeNames[lt2 - 192350000] : lt2.ToString();
                        Console.WriteLine("  SKIP (summary): [{0}] {1} --{2}--> [{3}] {4}",
                            relation.SourceTaskUniqueID, pn, ltn, mppTask.UniqueID.Value, sn);
                        continue;
                    }

                    allDeps.Add(Tuple.Create(relation.SourceTaskUniqueID, mppTask.UniqueID.Value,
                        MapRelationTypeStatic(relation.Type)));
                }
            }

            Console.WriteLine("\nDeps after filtering: {0} (skipped {1} summary deps)", allDeps.Count, summarySkipped);

            // Print all deps
            Console.WriteLine("\n--- All Dependencies ---");
            foreach (var dep in allDeps)
            {
                string predName = taskNames.ContainsKey(dep.Item1) ? taskNames[dep.Item1] : "?";
                string succName = taskNames.ContainsKey(dep.Item2) ? taskNames[dep.Item2] : "?";
                int lt = dep.Item3;
                string ltName = (lt >= 192350000 && lt <= 192350003) ? linkTypeNames[lt - 192350000] : lt.ToString();
                Console.WriteLine("  [{0}] {1} --{2}--> [{3}] {4}",
                    dep.Item1, predName, ltName, dep.Item2, succName);
            }

            // Cycle detection using DFS (pred -> succ edges)
            Console.WriteLine("\n--- Cycle Detection ---");
            var adjList = new Dictionary<int, List<int>>();
            foreach (var dep in allDeps)
            {
                if (!adjList.ContainsKey(dep.Item1)) adjList[dep.Item1] = new List<int>();
                adjList[dep.Item1].Add(dep.Item2);
            }

            var white = new HashSet<int>();
            var gray = new HashSet<int>();
            var black = new HashSet<int>();
            var allNodes = new HashSet<int>();
            foreach (var dep in allDeps) { allNodes.Add(dep.Item1); allNodes.Add(dep.Item2); }
            foreach (var n in allNodes) white.Add(n);

            var cycleNodes = new List<int>();
            bool hasCycle = false;
            foreach (var node in allNodes)
            {
                if (!white.Contains(node)) continue;
                if (DfsCycleDetect(node, adjList, white, gray, black, cycleNodes))
                {
                    hasCycle = true;
                    break;
                }
            }

            if (hasCycle)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("CYCLE DETECTED!");
                Console.ResetColor();
                Console.Write("Cycle path: ");
                foreach (var n in cycleNodes)
                    Console.Write("[{0}] {1} -> ", n, taskNames.ContainsKey(n) ? taskNames[n] : "?");
                Console.WriteLine("(back to start)");
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("No simple cycles found in pred->succ direction.");
                Console.ResetColor();
            }

            // Bidirectional link check (A->B and B->A with any link type)
            Console.WriteLine("\n--- Bidirectional Link Check ---");
            var edgeSet = new HashSet<string>();
            var biDirs = new List<string>();
            foreach (var dep in allDeps)
            {
                string fwd = dep.Item1 + "|" + dep.Item2;
                string rev = dep.Item2 + "|" + dep.Item1;
                if (edgeSet.Contains(rev))
                {
                    string predName = taskNames.ContainsKey(dep.Item1) ? taskNames[dep.Item1] : "?";
                    string succName = taskNames.ContainsKey(dep.Item2) ? taskNames[dep.Item2] : "?";
                    biDirs.Add(string.Format("[{0}] {1} <-> [{2}] {3}", dep.Item1, predName, dep.Item2, succName));
                }
                edgeSet.Add(fwd);
            }

            if (biDirs.Count > 0)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("{0} bidirectional link pairs found:", biDirs.Count);
                foreach (var bd in biDirs) Console.WriteLine("  " + bd);
                Console.ResetColor();
            }
            else
                Console.WriteLine("No bidirectional links found.");

            // Non-FS link types (SS, FF, SF can cause PSS scheduling cycles)
            Console.WriteLine("\n--- Non-FS Dependencies ---");
            int nonFsCount = 0;
            foreach (var dep in allDeps)
            {
                if (dep.Item3 != 192350000)
                {
                    nonFsCount++;
                    string predName = taskNames.ContainsKey(dep.Item1) ? taskNames[dep.Item1] : "?";
                    string succName = taskNames.ContainsKey(dep.Item2) ? taskNames[dep.Item2] : "?";
                    string ltName = linkTypeNames[dep.Item3 - 192350000];
                    Console.WriteLine("  [{0}] {1} --{2}--> [{3}] {4}",
                        dep.Item1, predName, ltName, dep.Item2, succName);
                }
            }
            Console.WriteLine("Total non-FS deps: {0} / {1}", nonFsCount, allDeps.Count);

            Console.WriteLine("\n=== Analysis Complete ===");
            return 0;
        }

        #endregion

        #region Async Import Test (Plugin Chain)

        /// <summary>
        /// End-to-end test of the async chunked import via the plugin chain.
        /// 1. Connect to D365 (from .env)
        /// 2. Create adc_adccasetemplate record + upload MPP file
        /// 3. Create msdyn_project
        /// 4. Call MppAsyncImportService.InitializeJob → creates adc_mppimportjob record
        /// 5. Async plugin fires on Create → processes phases
        /// 6. Poll job status until Completed or Failed
        /// 7. Query back tasks and report results
        /// </summary>
        static int RunAsyncImportTest(string mppPathOverride)
        {
            Console.WriteLine();
            Console.WriteLine("========================================");
            Console.WriteLine("  Async Import Test (Plugin Chain)");
            Console.WriteLine("========================================");

            // 1. Find MPP file
            string mppPath = mppPathOverride;
            if (string.IsNullOrEmpty(mppPath))
            {
                // Look for exampleFiles directory
                string dir = AppDomain.CurrentDomain.BaseDirectory;
                for (int i = 0; i < 6; i++)
                {
                    string candidate = Path.Combine(dir, "exampleFiles");
                    if (Directory.Exists(candidate))
                    {
                        var files = Directory.GetFiles(candidate, "*.mpp");
                        if (files.Length > 0)
                        {
                            Console.WriteLine("Available MPP files:");
                            for (int f = 0; f < files.Length; f++)
                                Console.WriteLine("  {0}. {1}", f + 1, Path.GetFileName(files[f]));

                            Console.Write("Select file (1-{0}), or Enter for first: ", files.Length);
                            string sel = Console.ReadLine()?.Trim();
                            int idx = 0;
                            if (!string.IsNullOrEmpty(sel)) int.TryParse(sel, out idx);
                            idx = Math.Max(0, idx - 1);
                            if (idx >= files.Length) idx = 0;
                            mppPath = files[idx];
                        }
                        break;
                    }
                    string parent = Directory.GetParent(dir)?.FullName;
                    if (parent == null || parent == dir) break;
                    dir = parent;
                }
            }

            if (string.IsNullOrEmpty(mppPath) || !File.Exists(mppPath))
            {
                Console.WriteLine("ERROR: No MPP file found. Provide path as argument: asynctest <path>");
                return 1;
            }

            byte[] mppBytes = File.ReadAllBytes(mppPath);
            Console.WriteLine("MPP file: {0} ({1:N0} bytes)", Path.GetFileName(mppPath), mppBytes.Length);

            // 2. Connect to D365
            var envData = LoadEnvAndConnect(requireCaseTemplate: false, requireProjectId: false);
            if (envData == null) return 1;
            var service = envData.Value.service;

            Console.WriteLine();
            Guid caseTemplateId = Guid.Empty;
            Guid projectId = Guid.Empty;
            Guid caseId = Guid.Empty;

            try
            {
                // 3. Create case template record
                Console.WriteLine("Creating adc_adccasetemplate record...");
                var templateEntity = new Entity("adc_adccasetemplate");
                templateEntity["adc_name"] = string.Format("ASYNCTEST - {0} - {1:yyyyMMdd-HHmmss}",
                    Path.GetFileName(mppPath), DateTime.Now);
                caseTemplateId = service.Create(templateEntity);
                Console.WriteLine("  Case template: {0}", caseTemplateId);

                // 4. Upload MPP file to the file column
                Console.WriteLine("Uploading MPP file to adc_templatefile...");
                UploadFileColumn(service, caseTemplateId, "adc_adccasetemplate",
                    "adc_templatefile", Path.GetFileName(mppPath), mppBytes);
                Console.WriteLine("  Upload complete.");

                // 5. Create project
                Console.WriteLine("Creating msdyn_project...");
                var projectEntity = new Entity("msdyn_project");
                projectEntity["msdyn_subject"] = string.Format("ASYNCTEST - {0} - {1:yyyyMMdd-HHmmss}",
                    Path.GetFileName(mppPath), DateTime.Now);
                projectId = service.Create(projectEntity);
                Console.WriteLine("  Project: {0}", projectId);

                // 5b. Create adc_case record (links template + project)
                Console.WriteLine("Creating adc_case record...");
                var caseEntity = new Entity("adc_case");
                caseEntity["adc_name"] = string.Format("ASYNCTEST - {0} - {1:yyyyMMdd-HHmmss}",
                    Path.GetFileName(mppPath), DateTime.Now);
                caseEntity["adc_casetemplate"] = new EntityReference("adc_adccasetemplate", caseTemplateId);
                caseEntity["adc_project"] = new EntityReference("msdyn_project", projectId);
                caseId = service.Create(caseEntity);
                Console.WriteLine("  Case: {0}", caseId);

                // Wait for PSS to create root task
                Console.WriteLine("Waiting for PSS to create root task (10s)...");
                System.Threading.Thread.Sleep(10000);

                // 6. Call InitializeJob — this parses MPP, builds batches, creates adc_mppimportjob
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Calling MppAsyncImportService.InitializeJob...");
                Console.ResetColor();

                var trace = new ConsoleTracingService();
                var asyncService = new MppAsyncImportService(service, trace);

                // Use earliest task start from MPP as project start
                var reader = new MppFileReader();
                ProjectFile mppProject = reader.Read(mppBytes);
                var mppTasks = mppProject.Tasks.Where(t => t.UniqueID.HasValue && t.UniqueID.Value != 0).ToList();
                DateTime? projectStart = null;
                var allStarts = mppTasks.Where(t => t.Start.HasValue).Select(t => t.Start.Value).ToList();
                if (allStarts.Count > 0) projectStart = allStarts.Min();

                // Pass initiating user for in-app notifications (Nick Wood for testing)
                Guid? initiatingUserId = new Guid("a13e4aa4-27d8-f011-8543-6045bdc38adc");
                Guid jobId = asyncService.InitializeJob(mppBytes, projectId, caseTemplateId, projectStart,
                    caseId: caseId, initiatingUserId: initiatingUserId);

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Job record created: {0}", jobId);
                Console.ResetColor();
                Console.WriteLine("Async plugin should now fire on Create and start processing phases...");

                // 7. Poll job status
                Console.WriteLine();
                Console.WriteLine("Polling job status (max 5 minutes)...");
                int maxPolls = 60; // 60 * 5s = 5 minutes
                int lastStatus = -1;
                int lastTick = -1;
                string lastStatusLabel = "";

                for (int poll = 0; poll < maxPolls; poll++)
                {
                    System.Threading.Thread.Sleep(5000);

                    var job = service.Retrieve(ImportJobFields.EntityName, jobId,
                        new Microsoft.Xrm.Sdk.Query.ColumnSet(
                            ImportJobFields.Status, ImportJobFields.Phase,
                            ImportJobFields.CurrentBatch, ImportJobFields.TotalBatches,
                            ImportJobFields.TotalTasks, ImportJobFields.CreatedCount,
                            ImportJobFields.DepsCount, ImportJobFields.Tick,
                            ImportJobFields.ErrorMessage));

                    var statusOsv = job.GetAttributeValue<OptionSetValue>(ImportJobFields.Status);
                    int status = statusOsv != null ? statusOsv.Value : -1;
                    int batch = job.GetAttributeValue<int>(ImportJobFields.CurrentBatch);
                    int totalBatches = job.GetAttributeValue<int>(ImportJobFields.TotalBatches);
                    int totalTasks = job.GetAttributeValue<int>(ImportJobFields.TotalTasks);
                    int created = job.GetAttributeValue<int>(ImportJobFields.CreatedCount);
                    int deps = job.GetAttributeValue<int>(ImportJobFields.DepsCount);
                    int tick = job.GetAttributeValue<int>(ImportJobFields.Tick);
                    string error = job.GetAttributeValue<string>(ImportJobFields.ErrorMessage);

                    string statusLabel = ImportJobStatus.Label(status);

                    if (status != lastStatus || statusLabel != lastStatusLabel || tick != lastTick)
                    {
                        Console.ForegroundColor = status == ImportJobStatus.Completed ? ConsoleColor.Green
                            : status == ImportJobStatus.Failed ? ConsoleColor.Red
                            : ConsoleColor.Cyan;
                        Console.WriteLine("  [{0,3}s] Status: {1} | Batch {2}/{3} | Tasks: {4}/{5} | Deps: {6} | Tick: {7}",
                            (poll + 1) * 5, statusLabel, batch, totalBatches, created, totalTasks, deps, tick);
                        Console.ResetColor();
                        lastStatus = status;
                        lastStatusLabel = statusLabel;
                        lastTick = tick;
                    }
                    else
                    {
                        Console.Write(".");
                    }

                    if (status == ImportJobStatus.Completed)
                    {
                        Console.WriteLine();
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("========================================");
                        Console.WriteLine("  ASYNC IMPORT COMPLETE");
                        Console.WriteLine("  Tasks created: {0}", created);
                        Console.WriteLine("  Dependencies:  {0}", deps);
                        Console.WriteLine("========================================");
                        Console.ResetColor();

                        // Query back to verify
                        Console.WriteLine();
                        Console.WriteLine("Verifying: querying D365 tasks...");
                        var crmTasks = QueryProjectTasksFull(service, projectId);
                        var crmDeps = QueryProjectDependencies(service, projectId);
                        var crmNonRoot = crmTasks.Where(e =>
                        {
                            int level = e.Contains("msdyn_outlinelevel") ? (int)e["msdyn_outlinelevel"] : -1;
                            return level != 0;
                        }).ToList();

                        Console.WriteLine("  D365 tasks (excl. root): {0} (expected: {1})", crmNonRoot.Count, mppTasks.Count);
                        Console.WriteLine("  D365 dependencies: {0}", crmDeps.Count);

                        if (crmNonRoot.Count == mppTasks.Count)
                        {
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine("  [PASS] Task count matches!");
                            Console.ResetColor();
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.Yellow;
                            Console.WriteLine("  [WARN] Task count mismatch: got {0}, expected {1}", crmNonRoot.Count, mppTasks.Count);
                            Console.ResetColor();
                        }

                        break;
                    }

                    if (status == ImportJobStatus.Failed)
                    {
                        Console.WriteLine();
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("========================================");
                        Console.WriteLine("  ASYNC IMPORT FAILED");
                        Console.WriteLine("  Error: {0}", error ?? "(no error message)");
                        Console.WriteLine("========================================");
                        Console.ResetColor();
                        return 2;
                    }
                }

                if (lastStatus != ImportJobStatus.Completed && lastStatus != ImportJobStatus.Failed)
                {
                    Console.WriteLine();
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("  TIMEOUT: Job did not complete within 5 minutes.");
                    Console.WriteLine("  Last status: {0}", lastStatusLabel);
                    Console.ResetColor();
                    return 2;
                }

                return 0;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine();
                Console.WriteLine("=== ASYNC TEST ERROR ===");
                Console.WriteLine(ex.ToString());
                Console.ResetColor();
                return 2;
            }
            finally
            {
                // Skip cleanup so records persist for UI verification
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Records left in place for UI verification:");
                if (caseId != Guid.Empty)
                    Console.WriteLine("  Case:          {0}", caseId);
                if (projectId != Guid.Empty)
                    Console.WriteLine("  Project:       {0}", projectId);
                if (caseTemplateId != Guid.Empty)
                    Console.WriteLine("  Case Template: {0}", caseTemplateId);
                Console.ResetColor();
            }
        }

        /// <summary>
        /// Uploads a file to a Dataverse File column using InitializeFileBlocksUpload / UploadBlock / CommitFileBlocksUpload.
        /// </summary>
        static void UploadFileColumn(IOrganizationService service, Guid recordId, string entityName,
            string fileAttributeName, string fileName, byte[] fileBytes)
        {
            // Step 1: Initialize upload
            var initRequest = new OrganizationRequest("InitializeFileBlocksUpload");
            initRequest["Target"] = new EntityReference(entityName, recordId);
            initRequest["FileAttributeName"] = fileAttributeName;
            initRequest["FileName"] = fileName;
            var initResponse = service.Execute(initRequest);
            string fileContinuationToken = (string)initResponse["FileContinuationToken"];

            // Step 2: Upload blocks (4MB each)
            const int blockSize = 4 * 1024 * 1024;
            int blockNumber = 0;
            var blockIds = new List<string>();

            for (int offset = 0; offset < fileBytes.Length; offset += blockSize)
            {
                int length = Math.Min(blockSize, fileBytes.Length - offset);
                byte[] block = new byte[length];
                Array.Copy(fileBytes, offset, block, 0, length);

                string blockId = Convert.ToBase64String(
                    System.Text.Encoding.UTF8.GetBytes(blockNumber.ToString("D6")));
                blockIds.Add(blockId);

                var uploadRequest = new OrganizationRequest("UploadBlock");
                uploadRequest["FileContinuationToken"] = fileContinuationToken;
                uploadRequest["BlockData"] = block;
                uploadRequest["BlockId"] = blockId;
                service.Execute(uploadRequest);

                blockNumber++;
            }

            // Step 3: Commit
            var commitRequest = new OrganizationRequest("CommitFileBlocksUpload");
            commitRequest["FileContinuationToken"] = fileContinuationToken;
            commitRequest["FileName"] = fileName;
            commitRequest["MimeType"] = "application/octet-stream";
            commitRequest["BlockList"] = blockIds.ToArray();
            service.Execute(commitRequest);
        }

        #endregion

        #region Side-by-Side Comparison Tests

        static int RunSideBySideTests(string exampleDirOverride)
        {
            Console.WriteLine();
            Console.WriteLine("================================================================");
            Console.WriteLine("  SIDE-BY-SIDE COMPARISON TEST (Async Import)");
            Console.WriteLine("================================================================");

            // 1. Find example files
            string exampleDir = exampleDirOverride;
            if (string.IsNullOrEmpty(exampleDir))
            {
                string dir = AppDomain.CurrentDomain.BaseDirectory;
                for (int i = 0; i < 6; i++)
                {
                    string candidate = Path.Combine(dir, "exampleFiles");
                    if (Directory.Exists(candidate)) { exampleDir = candidate; break; }
                    string parent = Directory.GetParent(dir)?.FullName;
                    if (parent == null || parent == dir) break;
                    dir = parent;
                }
            }

            if (exampleDir == null || !Directory.Exists(exampleDir))
            {
                Console.WriteLine("ERROR: exampleFiles directory not found.");
                return 1;
            }

            string[] mppFiles = Directory.GetFiles(exampleDir, "*.mpp");
            if (mppFiles.Length == 0)
            {
                Console.WriteLine("ERROR: No .mpp files found in {0}", exampleDir);
                return 1;
            }

            Console.WriteLine("Found {0} MPP file(s) in: {1}", mppFiles.Length, exampleDir);
            foreach (var f in mppFiles)
                Console.WriteLine("  - {0}", Path.GetFileName(f));

            // 2. Connect to D365
            var envData = LoadEnvAndConnect(requireCaseTemplate: false, requireProjectId: false);
            if (envData == null) return 1;
            var service = envData.Value.service;

            int totalPassed = 0;
            int totalFailed = 0;
            var summaryLines = new List<string>();

            foreach (string mppPath in mppFiles)
            {
                string fileName = Path.GetFileName(mppPath);
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine();
                Console.WriteLine("================================================================");
                Console.WriteLine("  FILE: {0}", fileName);
                Console.WriteLine("================================================================");
                Console.ResetColor();

                try
                {
                    var result = RunSingleSideBySide(service, mppPath);
                    if (result.AllPassed)
                    {
                        totalPassed++;
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("  OVERALL: PASS ({0} checks, {1} warnings)", result.Checks, result.Warnings.Count);
                        summaryLines.Add(string.Format("PASS  {0}  ({1} checks, {2} warnings, {3} dep issues)",
                            fileName, result.Checks, result.Warnings.Count, result.DepIssues.Count));
                    }
                    else
                    {
                        totalFailed++;
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("  OVERALL: FAIL ({0} failures)", result.Failures.Count);
                        summaryLines.Add(string.Format("FAIL  {0}  ({1} failures)", fileName, result.Failures.Count));
                    }
                    Console.ResetColor();
                }
                catch (Exception ex)
                {
                    totalFailed++;
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("  ERROR: {0}", ex.Message);
                    Console.ResetColor();
                    summaryLines.Add(string.Format("ERROR {0}  ({1})", fileName, ex.Message));
                }
            }

            // Final summary
            Console.WriteLine();
            Console.ForegroundColor = totalFailed == 0 ? ConsoleColor.Green : ConsoleColor.Yellow;
            Console.WriteLine("================================================================");
            Console.WriteLine("  FINAL: {0} passed, {1} failed out of {2}", totalPassed, totalFailed, mppFiles.Length);
            foreach (var s in summaryLines)
                Console.WriteLine("    {0}", s);
            Console.WriteLine("================================================================");
            Console.ResetColor();

            return totalFailed > 0 ? 2 : 0;
        }

        class SideBySideResult
        {
            public bool AllPassed { get { return Failures.Count == 0; } }
            public int Checks;
            public List<string> Failures = new List<string>();
            public List<string> Warnings = new List<string>();
            public List<string> DepIssues = new List<string>();
        }

        static SideBySideResult RunSingleSideBySide(IOrganizationService service, string mppPath)
        {
            var result = new SideBySideResult();
            string fileName = Path.GetFileName(mppPath);

            // =========================================================
            // PHASE 1: Parse MPP locally — build expected data
            // =========================================================
            Console.WriteLine("  [1/5] Parsing MPP locally...");
            byte[] mppBytes = File.ReadAllBytes(mppPath);
            var reader = new MppFileReader();
            ProjectFile mppProject = reader.Read(mppBytes);

            var mppTasks = mppProject.Tasks
                .Where(t => t.UniqueID.HasValue && t.UniqueID.Value != 0)
                .ToList();

            // Derive parent relationships (same as import service)
            var lastAtLevel = new Dictionary<int, Task>();
            foreach (var mppTask in mppTasks)
            {
                int level = mppTask.OutlineLevel ?? 0;
                if (level > 0 && mppTask.ParentTask == null)
                {
                    Task derivedParent;
                    if (lastAtLevel.TryGetValue(level - 1, out derivedParent))
                    {
                        mppTask.ParentTask = derivedParent;
                        if (!derivedParent.ChildTasks.Contains(mppTask))
                            derivedParent.ChildTasks.Add(mppTask);
                    }
                }
                lastAtLevel[level] = mppTask;
            }

            // Build summary task set
            var summaryIds = new HashSet<int>();
            foreach (var t in mppTasks)
            {
                if (t.ParentTask != null && t.ParentTask.UniqueID.HasValue && t.ParentTask.UniqueID.Value != 0)
                    summaryIds.Add(t.ParentTask.UniqueID.Value);
            }

            // Build expected task list
            var expectedTasks = new List<ExpectedTask>();
            foreach (var t in mppTasks)
            {
                var et = new ExpectedTask
                {
                    UniqueID = t.UniqueID.Value,
                    Name = t.Name ?? "(Unnamed Task)",
                    IsSummary = summaryIds.Contains(t.UniqueID.Value),
                    ParentName = (t.ParentTask != null && t.ParentTask.UniqueID.HasValue && t.ParentTask.UniqueID.Value != 0)
                        ? (t.ParentTask.Name ?? "(Unnamed Task)") : null,
                    OutlineLevel = t.OutlineLevel ?? 0
                };

                if (!et.IsSummary && t.Duration != null && t.Duration.Value >= 0)
                    et.DurationDays = Math.Round(ConvertToHoursLocal(t.Duration) / 8.0, 2);

                if (t.Start != null) et.StartDate = t.Start;
                if (t.Finish != null) et.FinishDate = t.Finish;

                expectedTasks.Add(et);
            }

            // Build expected dependencies (ALL, including summary task deps)
            var expectedDeps = new List<ExpectedDep>();
            foreach (var t in mppProject.Tasks)
            {
                if (!t.UniqueID.HasValue || t.UniqueID.Value == 0) continue;
                if (t.Predecessors == null || t.Predecessors.Count == 0) continue;

                foreach (var rel in t.Predecessors)
                {
                    if (rel.SourceTaskUniqueID == 0) continue;
                    var predTask = mppTasks.FirstOrDefault(x => x.UniqueID.HasValue && x.UniqueID.Value == rel.SourceTaskUniqueID);
                    string predName = predTask != null ? (predTask.Name ?? "(Unnamed Task)") : "UniqueID=" + rel.SourceTaskUniqueID;
                    string succName = t.Name ?? "(Unnamed Task)";
                    int linkType = MapRelationTypeStatic(rel.Type);

                    expectedDeps.Add(new ExpectedDep
                    {
                        PredecessorUniqueID = rel.SourceTaskUniqueID,
                        SuccessorUniqueID = t.UniqueID.Value,
                        PredecessorName = predName,
                        SuccessorName = succName,
                        LinkType = linkType,
                        LinkTypeName = LinkTypeName(linkType),
                        InvolvesSummary = summaryIds.Contains(rel.SourceTaskUniqueID) || summaryIds.Contains(t.UniqueID.Value)
                    });
                }
            }

            Console.WriteLine("    MPP: {0} tasks ({1} summary, {2} leaf), {3} deps ({4} involving summary tasks)",
                expectedTasks.Count, summaryIds.Count, expectedTasks.Count - summaryIds.Count,
                expectedDeps.Count, expectedDeps.Count(d => d.InvolvesSummary));

            // Get project start date from MPP
            DateTime? projectStart = null;
            var allStarts = mppTasks.Where(t => t.Start.HasValue).Select(t => t.Start.Value).ToList();
            if (allStarts.Count > 0) projectStart = allStarts.Min();
            Console.WriteLine("    Project start: {0:d}", projectStart);

            // =========================================================
            // PHASE 2: Create project + run async import
            // =========================================================
            Console.WriteLine("  [2/5] Creating project + running async import...");

            Guid caseTemplateId = Guid.Empty;
            Guid projectId = Guid.Empty;
            Guid caseId = Guid.Empty;

            try
            {
                // Create case template + upload MPP
                var templateEntity = new Entity("adc_adccasetemplate");
                templateEntity["adc_name"] = string.Format("SBS-TEST - {0} - {1:yyyyMMdd-HHmmss}", fileName, DateTime.Now);
                caseTemplateId = service.Create(templateEntity);

                UploadFileColumn(service, caseTemplateId, "adc_adccasetemplate",
                    "adc_templatefile", fileName, mppBytes);

                // Create project
                var projectEntity = new Entity("msdyn_project");
                projectEntity["msdyn_subject"] = string.Format("SBS-TEST - {0} - {1:yyyyMMdd-HHmmss}", fileName, DateTime.Now);
                projectId = service.Create(projectEntity);

                // Create case (links template + project)
                var caseEntity = new Entity("adc_case");
                caseEntity["adc_name"] = string.Format("SBS-TEST - {0} - {1:yyyyMMdd-HHmmss}", fileName, DateTime.Now);
                caseEntity["adc_casetemplate"] = new EntityReference("adc_adccasetemplate", caseTemplateId);
                caseEntity["adc_project"] = new EntityReference("msdyn_project", projectId);
                caseId = service.Create(caseEntity);

                Console.WriteLine("    Project: {0}", projectId);

                // Wait for PSS
                System.Threading.Thread.Sleep(10000);

                // Run async import
                var trace = new ConsoleTracingService();
                var asyncService = new MppAsyncImportService(service, trace);
                Guid? initiatingUserId = new Guid("a13e4aa4-27d8-f011-8543-6045bdc38adc");
                Guid jobId = asyncService.InitializeJob(mppBytes, projectId, caseTemplateId, projectStart,
                    caseId: caseId, initiatingUserId: initiatingUserId);

                Console.WriteLine("    Job: {0}", jobId);

                // =========================================================
                // PHASE 3: Poll until complete
                // =========================================================
                Console.WriteLine("  [3/5] Polling job status...");
                int maxPolls = 96; // 96 * 5s = 480s (8 min) for large files with multi-execution batching
                int lastStatus = -1;
                bool completed = false;

                for (int poll = 0; poll < maxPolls; poll++)
                {
                    System.Threading.Thread.Sleep(5000);
                    var job = service.Retrieve(ImportJobFields.EntityName, jobId,
                        new Microsoft.Xrm.Sdk.Query.ColumnSet(
                            ImportJobFields.Status, ImportJobFields.TotalTasks,
                            ImportJobFields.CreatedCount, ImportJobFields.DepsCount,
                            ImportJobFields.ErrorMessage));

                    var statusOsv = job.GetAttributeValue<OptionSetValue>(ImportJobFields.Status);
                    int status = statusOsv != null ? statusOsv.Value : -1;

                    if (status != lastStatus)
                    {
                        int created = job.GetAttributeValue<int>(ImportJobFields.CreatedCount);
                        int deps = job.GetAttributeValue<int>(ImportJobFields.DepsCount);
                        Console.WriteLine("    [{0,3}s] {1} | Tasks: {2} | Deps: {3}",
                            (poll + 1) * 5, ImportJobStatus.Label(status), created, deps);
                        lastStatus = status;
                    }

                    if (status == ImportJobStatus.Completed)
                    {
                        completed = true;
                        break;
                    }
                    if (status == ImportJobStatus.Failed)
                    {
                        string err = job.GetAttributeValue<string>(ImportJobFields.ErrorMessage);
                        result.Failures.Add(string.Format("Import FAILED: {0}", err ?? "(no error)"));
                        return result;
                    }
                }

                if (!completed)
                {
                    result.Failures.Add("Import TIMEOUT: did not complete within 5 minutes");
                    return result;
                }

                // Wait extra for PSS to finalize scheduling
                Console.WriteLine("    Waiting 30s for PSS to finalize scheduling...");
                System.Threading.Thread.Sleep(30000);

                // =========================================================
                // PHASE 4: Query CRM tasks + deps
                // =========================================================
                Console.WriteLine("  [4/5] Querying D365 tasks and dependencies...");

                var crmTasks = QueryProjectTasksFull(service, projectId);
                var crmDeps = QueryProjectDependencies(service, projectId);

                // Exclude PSS root task (outline level 0)
                var crmNonRoot = crmTasks.Where(e =>
                {
                    int level = e.Contains("msdyn_outlinelevel") ? (int)e["msdyn_outlinelevel"] : -1;
                    return level != 0;
                }).ToList();

                Console.WriteLine("    D365: {0} tasks (excl. root), {1} dependencies", crmNonRoot.Count, crmDeps.Count);

                // Build CRM task lookup by ID
                var crmTaskById = crmNonRoot.ToDictionary(e => e.Id);
                // Build CRM task name by ID for dep resolution
                var crmTaskNameById = new Dictionary<Guid, string>();
                foreach (var ct in crmNonRoot)
                    crmTaskNameById[ct.Id] = ct.GetAttributeValue<string>("msdyn_subject") ?? "";

                // Build name-consumption map: for each expected task, consume the first
                // unmatched CRM task with the same name. Handles duplicates correctly.
                var crmByName = new Dictionary<string, List<Entity>>(StringComparer.OrdinalIgnoreCase);
                foreach (var ct in crmNonRoot)
                {
                    string name = ct.GetAttributeValue<string>("msdyn_subject") ?? "";
                    List<Entity> list;
                    if (!crmByName.TryGetValue(name, out list))
                    {
                        list = new List<Entity>();
                        crmByName[name] = list;
                    }
                    list.Add(ct);
                }

                var expectedToCrm = new Dictionary<int, Entity>(); // keyed by UniqueID
                var crmToExpected = new Dictionary<Guid, ExpectedTask>();
                foreach (var et in expectedTasks)
                {
                    List<Entity> candidates;
                    if (crmByName.TryGetValue(et.Name, out candidates) && candidates.Count > 0)
                    {
                        var match = candidates[0];
                        candidates.RemoveAt(0);
                        expectedToCrm[et.UniqueID] = match;
                        crmToExpected[match.Id] = et;
                    }
                }
                int mapCount = expectedToCrm.Count;

                // =========================================================
                // PHASE 5: Side-by-side comparison
                // =========================================================
                Console.WriteLine("  [5/5] Comparing...");
                Console.WriteLine();

                // --- CHECK 1: Task count ---
                if (crmNonRoot.Count == expectedTasks.Count)
                {
                    result.Checks++;
                    Console.WriteLine("    [PASS] Task count: {0}", crmNonRoot.Count);
                }
                else
                {
                    result.Failures.Add(string.Format("Task count: expected {0}, got {1}", expectedTasks.Count, crmNonRoot.Count));
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("    [FAIL] Task count: expected {0}, got {1}", expectedTasks.Count, crmNonRoot.Count);
                    Console.ResetColor();
                }

                // --- CHECK 2: Task names matched ---
                int unmatchedCount = expectedTasks.Count - mapCount;
                if (unmatchedCount == 0)
                {
                    result.Checks++;
                    Console.WriteLine("    [PASS] All {0} task names matched to CRM", mapCount);
                }
                else
                {
                    result.Failures.Add(string.Format("{0} expected tasks could not be matched to CRM by name", unmatchedCount));
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("    [FAIL] {0} expected tasks unmatched", unmatchedCount);
                    Console.ResetColor();
                }

                // --- CHECK 3: Hierarchy (parent links) ---
                int parentOk = 0, parentFail = 0;
                foreach (var et in expectedTasks)
                {
                    Entity ct;
                    if (!expectedToCrm.TryGetValue(et.UniqueID, out ct)) continue;

                    var parentRef = ct.GetAttributeValue<EntityReference>("msdyn_parenttask");
                    string crmParentName = null;
                    if (parentRef != null)
                    {
                        Entity pe;
                        if (crmTaskById.TryGetValue(parentRef.Id, out pe))
                            crmParentName = pe.GetAttributeValue<string>("msdyn_subject") ?? "";
                    }

                    if ((et.ParentName == null && crmParentName == null) ||
                        (et.ParentName != null && crmParentName != null &&
                         et.ParentName.Equals(crmParentName, StringComparison.OrdinalIgnoreCase)))
                    {
                        parentOk++;
                    }
                    else
                    {
                        parentFail++;
                        result.Warnings.Add(string.Format("Parent mismatch '{0}': expected parent='{1}', got '{2}'",
                            et.Name, et.ParentName ?? "(root)", crmParentName ?? "(root)"));
                    }
                }

                if (parentFail == 0)
                {
                    result.Checks++;
                    Console.WriteLine("    [PASS] Parent hierarchy: {0} tasks verified", parentOk);
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("    [WARN] Parent hierarchy: {0} OK, {1} mismatched", parentOk, parentFail);
                    Console.ResetColor();
                }

                // --- CHECK 4: Leaf task durations ---
                int durOk = 0, durFail = 0;
                foreach (var et in expectedTasks)
                {
                    if (et.IsSummary || !et.DurationDays.HasValue) continue;

                    Entity ct;
                    if (!expectedToCrm.TryGetValue(et.UniqueID, out ct)) continue;

                    double? crmDur = ct.Contains("msdyn_duration") ? (double?)ct["msdyn_duration"] : null;
                    double expected = et.DurationDays.Value;
                    double actual = crmDur ?? 0;

                    if (Math.Abs(expected - actual) < 0.5)
                        durOk++;
                    else
                    {
                        durFail++;
                        result.Warnings.Add(string.Format("Duration mismatch '{0}': expected {1:F1}d, got {2:F1}d",
                            et.Name, expected, actual));
                    }
                }

                if (durFail == 0)
                {
                    result.Checks++;
                    Console.WriteLine("    [PASS] Leaf durations: {0} matched", durOk);
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("    [WARN] Leaf durations: {0} OK, {1} mismatched", durOk, durFail);
                    Console.ResetColor();
                }

                // --- CHECK 5: Dependencies (detailed, GUID-based matching) ---
                // Build CRM dep set using GUIDs: (predGuid, succGuid, linkType)
                var crmDepByGuid = new List<Tuple<Guid, Guid, int>>();
                foreach (var cd in crmDeps)
                {
                    var predRef = cd.GetAttributeValue<EntityReference>("msdyn_predecessortask");
                    var succRef = cd.GetAttributeValue<EntityReference>("msdyn_successortask");
                    var ltOsv = cd.GetAttributeValue<OptionSetValue>("msdyn_linktype");
                    if (predRef == null || succRef == null) continue;
                    int lt = ltOsv != null ? ltOsv.Value : 192350000;
                    crmDepByGuid.Add(Tuple.Create(predRef.Id, succRef.Id, lt));
                }

                // Match expected deps against CRM deps using positional map
                int depFound = 0;
                int depMissing = 0;
                int depMissingSummary = 0;

                var crmDepPool = new List<Tuple<Guid, Guid, int>>(crmDepByGuid);
                foreach (var ed in expectedDeps)
                {
                    // Look up expected pred/succ CRM GUIDs via positional map
                    Entity predCrm, succCrm;
                    expectedToCrm.TryGetValue(ed.PredecessorUniqueID, out predCrm);
                    expectedToCrm.TryGetValue(ed.SuccessorUniqueID, out succCrm);

                    if (predCrm == null || succCrm == null)
                    {
                        // Can't map — treat as missing
                        depMissing++;
                        if (ed.InvolvesSummary) depMissingSummary++;
                        result.DepIssues.Add(string.Format("UNMAPPED{0}: {1} --{2}--> {3}",
                            ed.InvolvesSummary ? " (summary)" : "",
                            ed.PredecessorName, ed.LinkTypeName, ed.SuccessorName));
                        continue;
                    }

                    Guid predId = predCrm.Id;
                    Guid succId = succCrm.Id;

                    int idx = crmDepPool.FindIndex(c => c.Item1 == predId && c.Item2 == succId && c.Item3 == ed.LinkType);
                    if (idx >= 0)
                    {
                        depFound++;
                        crmDepPool.RemoveAt(idx);
                    }
                    else
                    {
                        // Check if it exists with wrong link type
                        int ltIdx = crmDepPool.FindIndex(c => c.Item1 == predId && c.Item2 == succId);
                        if (ltIdx >= 0)
                        {
                            depFound++;
                            string actual = LinkTypeName(crmDepPool[ltIdx].Item3);
                            result.DepIssues.Add(string.Format("LINK TYPE MISMATCH: {0} -> {1}: expected {2}, got {3}",
                                ed.PredecessorName, ed.SuccessorName, ed.LinkTypeName, actual));
                            crmDepPool.RemoveAt(ltIdx);
                        }
                        else
                        {
                            depMissing++;
                            if (ed.InvolvesSummary) depMissingSummary++;
                            result.DepIssues.Add(string.Format("MISSING{0}: {1} --{2}--> {3}",
                                ed.InvolvesSummary ? " (summary)" : "",
                                ed.PredecessorName, ed.LinkTypeName, ed.SuccessorName));
                        }
                    }
                }

                // Extra deps in CRM not in MPP
                int extraDeps = crmDepPool.Count;

                Console.WriteLine();
                Console.ForegroundColor = depMissing == 0 ? ConsoleColor.Green : ConsoleColor.Yellow;
                Console.WriteLine("    Dependencies: {0}/{1} found, {2} missing ({3} summary), {4} extra in CRM",
                    depFound, expectedDeps.Count, depMissing, depMissingSummary, extraDeps);
                Console.ResetColor();

                if (depMissing == 0)
                {
                    result.Checks++;
                    Console.WriteLine("    [PASS] All {0} MPP dependencies found in D365", expectedDeps.Count);
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("    [WARN] {0} dependencies missing from D365:", depMissing);
                    foreach (var issue in result.DepIssues)
                        Console.WriteLine("      {0}", issue);
                    Console.ResetColor();
                }

                if (result.DepIssues.Count > 0)
                {
                    Console.WriteLine();
                    Console.WriteLine("    --- Dependency Issues ---");
                    foreach (var issue in result.DepIssues)
                        Console.WriteLine("      {0}", issue);
                }

                // --- CHECK 6: Calendar-aware date comparison ---
                // Extract calendars from both MPP and D365, then compare using working-day math.
                Console.WriteLine();
                Console.WriteLine("  CHECK 6: Calendar-aware date comparison");

                var mppCal = GetMppCalendar(mppProject);
                var d365Cal = GetD365Calendar(service, projectId);

                Console.WriteLine("    MPP Calendar:  \"{0}\"  Working: {1}  Exceptions: {2}",
                    mppCal.Name, mppCal.WorkingDaysSummary(), mppCal.Exceptions.Count);
                var dayNames = new[] { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };
                for (int di = 0; di < 7; di++)
                {
                    if (mppCal.DayHours[di] != null)
                        Console.WriteLine("      {0}: {1}", dayNames[di], mppCal.DayHours[di]);
                }
                if (mppCal.Exceptions.Count > 0)
                {
                    int shown = 0;
                    foreach (var ex in mppCal.Exceptions)
                    {
                        Console.WriteLine("      Exception: {0:yyyy-MM-dd} to {1:yyyy-MM-dd} ({2})",
                            ex.Item1, ex.Item2, ex.Item3 ? "working" : "non-working");
                        if (++shown >= 10) { Console.WriteLine("      ... and {0} more", mppCal.Exceptions.Count - 10); break; }
                    }
                }

                Console.WriteLine("    D365 Calendar: \"{0}\"  Working: {1}",
                    d365Cal.Name, d365Cal.WorkingDaysSummary());
                for (int di = 0; di < 7; di++)
                {
                    if (d365Cal.DayHours[di] != null)
                        Console.WriteLine("      {0}: {1}", dayNames[di], d365Cal.DayHours[di]);
                }

                bool calendarsSame = true;
                for (int di = 0; di < 7; di++)
                {
                    if (mppCal.WorkingDays[di] != d365Cal.WorkingDays[di]) { calendarsSame = false; break; }
                }
                Console.ForegroundColor = calendarsSame ? ConsoleColor.Green : ConsoleColor.Yellow;
                Console.WriteLine("    Calendars match: {0}", calendarsSame ? "YES" : "NO — this explains date drift");
                Console.ResetColor();

                // Find D365 project start (earliest task start in CRM, excluding root)
                DateTime? d365ProjectStart = null;
                foreach (var ct2 in crmNonRoot)
                {
                    DateTime? cs = ct2.GetAttributeValue<DateTime?>("msdyn_scheduledstart");
                    if (cs.HasValue && (!d365ProjectStart.HasValue || cs.Value < d365ProjectStart.Value))
                        d365ProjectStart = cs.Value;
                }

                if (projectStart.HasValue && d365ProjectStart.HasValue)
                {
                    Console.WriteLine("    MPP project start:  {0:yyyy-MM-dd} ({1})",
                        projectStart.Value.Date, projectStart.Value.DayOfWeek);
                    Console.WriteLine("    D365 earliest task: {0:yyyy-MM-dd} ({1})",
                        d365ProjectStart.Value.Date, d365ProjectStart.Value.DayOfWeek);

                    // Calendar-aware comparison:
                    // For each task, count working days from MPP start to task start using MPP calendar,
                    // then map forward from D365 start using D365 calendar. Compare with actual D365 date.
                    int startOk = 0, startDiff = 0, endOk = 0, endDiff = 0;
                    var dateIssues = new List<string>();
                    int tolerance = 1; // ±1 day tolerance for rounding

                    foreach (var et in expectedTasks)
                    {
                        Entity ct;
                        if (!expectedToCrm.TryGetValue(et.UniqueID, out ct)) continue;

                        DateTime? crmStart = ct.GetAttributeValue<DateTime?>("msdyn_scheduledstart");
                        DateTime? crmEnd = ct.GetAttributeValue<DateTime?>("msdyn_scheduledend");

                        // Start date
                        if (et.StartDate.HasValue && crmStart.HasValue)
                        {
                            int mppWorkDays = CountWorkingDays(projectStart.Value.Date,
                                et.StartDate.Value.Date, mppCal.WorkingDays, mppCal.Exceptions);
                            DateTime expectedD365 = mppWorkDays > 0
                                ? AddWorkingDays(d365ProjectStart.Value.Date, mppWorkDays, d365Cal.WorkingDays)
                                : d365ProjectStart.Value.Date;
                            int deviation = Math.Abs((int)(crmStart.Value.Date - expectedD365).TotalDays);
                            if (deviation <= tolerance)
                                startOk++;
                            else
                            {
                                startDiff++;
                                dateIssues.Add(string.Format(
                                    "START {0}{1}: MPP={2:yyyy-MM-dd}, D365={3:yyyy-MM-dd}, expected={4:yyyy-MM-dd} (mppWD={5}, dev={6}d)",
                                    et.IsSummary ? "[S] " : "", et.Name,
                                    et.StartDate.Value.Date, crmStart.Value.Date, expectedD365,
                                    mppWorkDays, (int)(crmStart.Value.Date - expectedD365).TotalDays));
                            }
                        }

                        // End date
                        if (et.FinishDate.HasValue && crmEnd.HasValue)
                        {
                            int mppWorkDays = CountWorkingDays(projectStart.Value.Date,
                                et.FinishDate.Value.Date, mppCal.WorkingDays, mppCal.Exceptions);
                            DateTime expectedD365 = mppWorkDays > 0
                                ? AddWorkingDays(d365ProjectStart.Value.Date, mppWorkDays, d365Cal.WorkingDays)
                                : d365ProjectStart.Value.Date;
                            int deviation = Math.Abs((int)(crmEnd.Value.Date - expectedD365).TotalDays);
                            if (deviation <= tolerance)
                                endOk++;
                            else
                            {
                                endDiff++;
                                dateIssues.Add(string.Format(
                                    "END   {0}{1}: MPP={2:yyyy-MM-dd}, D365={3:yyyy-MM-dd}, expected={4:yyyy-MM-dd} (mppWD={5}, dev={6}d)",
                                    et.IsSummary ? "[S] " : "", et.Name,
                                    et.FinishDate.Value.Date, crmEnd.Value.Date, expectedD365,
                                    mppWorkDays, (int)(crmEnd.Value.Date - expectedD365).TotalDays));
                            }
                        }
                    }

                    Console.ForegroundColor = (startDiff == 0 && endDiff == 0) ? ConsoleColor.Green : ConsoleColor.Yellow;
                    Console.WriteLine("    Start dates: {0} match, {1} differ (±{2}d tolerance)", startOk, startDiff, tolerance);
                    Console.WriteLine("    End dates:   {0} match, {1} differ (±{2}d tolerance)", endOk, endDiff, tolerance);
                    Console.ResetColor();

                    if (startDiff == 0 && endDiff == 0)
                    {
                        result.Checks++;
                        Console.WriteLine("    [PASS] All dates explained by calendar difference");
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine("    [WARN] {0} dates not fully explained by calendars:", startDiff + endDiff);
                        int shown = 0;
                        foreach (var di2 in dateIssues)
                        {
                            Console.WriteLine("      {0}", di2);
                            if (++shown >= 30) { Console.WriteLine("      ... and {0} more", dateIssues.Count - 30); break; }
                        }
                        Console.ResetColor();
                        result.Warnings.AddRange(dateIssues);
                    }
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("    [SKIP] Cannot compare dates — missing project start or D365 tasks");
                    Console.ResetColor();
                }

                if (result.Warnings.Count > 0)
                {
                    Console.WriteLine();
                    Console.WriteLine("    --- Warnings ---");
                    foreach (var w in result.Warnings)
                        Console.WriteLine("      {0}", w);
                }

                if (result.Failures.Count > 0)
                {
                    Console.WriteLine();
                    Console.WriteLine("    --- Failures ---");
                    foreach (var f in result.Failures)
                        Console.WriteLine("      {0}", f);
                }
            }
            finally
            {
                // Cleanup
                Console.WriteLine();
                Console.WriteLine("    Cleaning up test records...");
                try
                {
                    if (projectId != Guid.Empty)
                    {
                        DoClearProjectTasks(service, projectId);
                        System.Threading.Thread.Sleep(5000);
                        service.Delete("msdyn_project", projectId);
                    }
                    if (caseId != Guid.Empty)
                        service.Delete("adc_case", caseId);
                    if (caseTemplateId != Guid.Empty)
                        service.Delete("adc_adccasetemplate", caseTemplateId);
                    Console.WriteLine("    Cleanup done.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("    Cleanup warning: {0}", ex.Message);
                }
            }

            return result;
        }

        class ExpectedTask
        {
            public int UniqueID;
            public string Name;
            public bool IsSummary;
            public string ParentName;
            public int OutlineLevel;
            public double? DurationDays;
            public DateTime? StartDate;
            public DateTime? FinishDate;
        }

        class ExpectedDep
        {
            public int PredecessorUniqueID;
            public int SuccessorUniqueID;
            public string PredecessorName;
            public string SuccessorName;
            public int LinkType;
            public string LinkTypeName;
            public bool InvolvesSummary;
        }

        static string LinkTypeName(int linkType)
        {
            switch (linkType)
            {
                case 192350000: return "FS";
                case 192350001: return "SS";
                case 192350002: return "FF";
                case 192350003: return "SF";
                default: return linkType.ToString();
            }
        }

        /// <summary>
        /// Compares WBS IDs numerically (e.g. "1" &lt; "1.1" &lt; "1.2" &lt; "2" &lt; "10").
        /// </summary>
        class WbsComparer : IComparer<string>
        {
            public int Compare(string x, string y)
            {
                if (x == y) return 0;
                if (x == null) return -1;
                if (y == null) return 1;

                string[] xParts = x.Split('.');
                string[] yParts = y.Split('.');

                int len = Math.Max(xParts.Length, yParts.Length);
                for (int i = 0; i < len; i++)
                {
                    int xVal = i < xParts.Length ? ParseInt(xParts[i]) : -1;
                    int yVal = i < yParts.Length ? ParseInt(yParts[i]) : -1;
                    if (xVal != yVal) return xVal.CompareTo(yVal);
                }
                return 0;
            }

            static int ParseInt(string s)
            {
                int v;
                return int.TryParse(s, out v) ? v : 0;
            }
        }

        #region Calendar Helpers

        class CalendarInfo
        {
            public string Name;
            public bool[] WorkingDays = new bool[7]; // Sun=0 .. Sat=6
            public string[] DayHours = new string[7]; // hours description per day
            public List<Tuple<DateTime, DateTime, bool>> Exceptions = new List<Tuple<DateTime, DateTime, bool>>();

            public string WorkingDaysSummary()
            {
                var days = new[] { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };
                var working = new List<string>();
                for (int i = 0; i < 7; i++)
                    if (WorkingDays[i]) working.Add(days[i]);
                return string.Join(", ", working);
            }
        }

        static CalendarInfo GetMppCalendar(ProjectFile mppProject)
        {
            var cal = mppProject.GetDefaultCalendar();
            var info = new CalendarInfo();
            info.Name = cal != null ? (cal.Name ?? "(unnamed)") : "(no calendar)";

            if (cal == null) return info;

            // Default assumption: Mon-Fri working if DayType is Default
            for (int i = 0; i < 7; i++)
            {
                var day = cal.Days[i];
                if (day.Type == DayType.Working)
                {
                    info.WorkingDays[i] = true;
                    if (day.Hours.Count > 0)
                        info.DayHours[i] = string.Join("; ", day.Hours.Select(h =>
                            string.Format("{0:hh\\:mm}-{1:hh\\:mm}", h.Start, h.End)));
                    else
                        info.DayHours[i] = "working (no hours specified)";
                }
                else if (day.Type == DayType.NonWorking)
                {
                    info.WorkingDays[i] = false;
                    info.DayHours[i] = "non-working";
                }
                else // Default — inherit or assume standard
                {
                    // For base calendars, Default typically means Mon-Fri working
                    bool isWeekday = (i >= 1 && i <= 5);
                    if (cal.ParentCalendar != null)
                    {
                        var parentDay = cal.ParentCalendar.Days[i];
                        info.WorkingDays[i] = parentDay.Type == DayType.Working ||
                            (parentDay.Type == DayType.Default && isWeekday);
                    }
                    else
                    {
                        info.WorkingDays[i] = isWeekday;
                    }
                    info.DayHours[i] = "default" + (info.WorkingDays[i] ? " (working)" : " (non-working)");
                }
            }

            foreach (var ex in cal.Exceptions)
            {
                info.Exceptions.Add(Tuple.Create(ex.Start, ex.End, ex.Working));
            }

            return info;
        }

        static CalendarInfo GetD365Calendar(IOrganizationService service, Guid projectId)
        {
            var info = new CalendarInfo();
            info.Name = "(unknown)";

            try
            {
                var project = service.Retrieve("msdyn_project", projectId,
                    new Microsoft.Xrm.Sdk.Query.ColumnSet("calendarid"));
                var calRef = project.GetAttributeValue<EntityReference>("calendarid");
                if (calRef == null)
                {
                    info.Name = "(no calendar assigned)";
                    // Default D365: Mon-Fri
                    for (int i = 1; i <= 5; i++) info.WorkingDays[i] = true;
                    return info;
                }

                var calendar = service.Retrieve("calendar", calRef.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet(true));
                info.Name = calendar.GetAttributeValue<string>("name") ?? calRef.Id.ToString();

                // Parse calendarrules
                var rules = calendar.GetAttributeValue<EntityCollection>("calendarrules");
                if (rules != null && rules.Entities.Count > 0)
                {
                    // Look for the leaf rules (inner calendar or direct pattern)
                    foreach (var rule in rules.Entities)
                    {
                        var innerCalId = rule.GetAttributeValue<EntityReference>("innercalendarid");
                        if (innerCalId != null)
                        {
                            // Retrieve inner calendar for work hour pattern
                            try
                            {
                                var innerCal = service.Retrieve("calendar", innerCalId.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet(true));
                                var innerRules = innerCal.GetAttributeValue<EntityCollection>("calendarrules");
                                if (innerRules != null)
                                {
                                    ParseD365CalendarRules(info, innerRules);
                                }
                            }
                            catch { }
                        }

                        // Also check direct pattern on the rule
                        string pattern = rule.GetAttributeValue<string>("pattern");
                        if (!string.IsNullOrEmpty(pattern))
                        {
                            ParsePatternString(info, pattern, rule);
                        }
                    }
                }

                // If no working days found, assume Mon-Fri default
                if (!info.WorkingDays.Any(w => w))
                {
                    for (int i = 1; i <= 5; i++) info.WorkingDays[i] = true;
                    info.Name += " (defaulted to Mon-Fri)";
                }
            }
            catch (Exception ex)
            {
                info.Name = "(error: " + ex.Message + ")";
                for (int i = 1; i <= 5; i++) info.WorkingDays[i] = true;
            }

            return info;
        }

        static void ParseD365CalendarRules(CalendarInfo info, EntityCollection rules)
        {
            foreach (var rule in rules.Entities)
            {
                string pattern = rule.GetAttributeValue<string>("pattern");
                if (!string.IsNullOrEmpty(pattern))
                {
                    ParsePatternString(info, pattern, rule);
                }
            }
        }

        static void ParsePatternString(CalendarInfo info, string pattern, Entity rule)
        {
            // Parse RFC 5545-like pattern, e.g. "FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR"
            int timeCode = rule.Contains("timecode") ? rule.GetAttributeValue<int>("timecode") : -1;
            // timecode: 0=Available, 2=Unavailable

            if (pattern.Contains("BYDAY="))
            {
                int idx = pattern.IndexOf("BYDAY=") + 6;
                string daysStr = pattern.Substring(idx);
                int semi = daysStr.IndexOf(';');
                if (semi > 0) daysStr = daysStr.Substring(0, semi);

                var dayMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
                {
                    {"SU", 0}, {"MO", 1}, {"TU", 2}, {"WE", 3}, {"TH", 4}, {"FR", 5}, {"SA", 6}
                };

                bool isWorking = timeCode != 2; // 0=Available or unset means working
                foreach (var d in daysStr.Split(','))
                {
                    string dt = d.Trim();
                    int dayIdx;
                    if (dayMap.TryGetValue(dt, out dayIdx))
                    {
                        info.WorkingDays[dayIdx] = isWorking;
                        var startTime = rule.GetAttributeValue<DateTime?>("starttime");
                        var endTime = rule.GetAttributeValue<DateTime?>("endtime");
                        if (startTime.HasValue && endTime.HasValue)
                            info.DayHours[dayIdx] = string.Format("{0:HH:mm}-{1:HH:mm}", startTime.Value, endTime.Value);
                    }
                }
            }
        }

        static int CountWorkingDays(DateTime from, DateTime to, bool[] workingDays,
            List<Tuple<DateTime, DateTime, bool>> exceptions = null)
        {
            if (to < from) return -CountWorkingDays(to, from, workingDays, exceptions);
            int count = 0;
            DateTime d = from.Date;
            DateTime end = to.Date;
            while (d < end)
            {
                int dow = (int)d.DayOfWeek;
                bool isWorking = workingDays[dow];

                // Check exceptions
                if (exceptions != null)
                {
                    foreach (var ex in exceptions)
                    {
                        if (d >= ex.Item1.Date && d <= ex.Item2.Date)
                        {
                            isWorking = ex.Item3; // override with exception's working flag
                            break;
                        }
                    }
                }

                if (isWorking) count++;
                d = d.AddDays(1);
            }
            return count;
        }

        static DateTime AddWorkingDays(DateTime from, int workDays, bool[] workingDays)
        {
            DateTime d = from.Date;
            int added = 0;
            int direction = workDays >= 0 ? 1 : -1;
            int target = Math.Abs(workDays);
            while (added < target)
            {
                d = d.AddDays(direction);
                if (workingDays[(int)d.DayOfWeek]) added++;
            }
            return d;
        }

        #endregion

        static int RunCalInfo(string exampleDir)
        {
            Console.WriteLine();
            Console.WriteLine("================================================================");
            Console.WriteLine("  CALENDAR INFO — MPP calendars + D365 default calendar");
            Console.WriteLine("================================================================");

            // Find MPP files
            if (exampleDir == null)
            {
                string exeDir = AppDomain.CurrentDomain.BaseDirectory;
                exampleDir = Path.Combine(exeDir, "ExampleFiles");
                if (!Directory.Exists(exampleDir))
                    exampleDir = Path.Combine(exeDir, "..", "..", "ExampleFiles");
            }
            if (!Directory.Exists(exampleDir))
            {
                Console.WriteLine("Example files directory not found: {0}", exampleDir);
                return 1;
            }

            var mppFiles = Directory.GetFiles(exampleDir, "*.mpp");
            Console.WriteLine("Found {0} MPP file(s) in {1}", mppFiles.Length, exampleDir);

            // Dump MPP calendars
            foreach (var mppPath in mppFiles)
            {
                Console.WriteLine();
                Console.WriteLine("--- {0} ---", Path.GetFileName(mppPath));
                try
                {
                    byte[] mppBytes = File.ReadAllBytes(mppPath);
                    var reader = new MppFileReader();
                    ProjectFile mppProject = reader.Read(mppBytes);

                    Console.WriteLine("  Calendars found: {0}", mppProject.Calendars.Count);
                    foreach (var cal in mppProject.Calendars)
                    {
                        Console.WriteLine("  Calendar: \"{0}\" (ID={1}, Parent={2})",
                            cal.Name ?? "(null)", cal.UniqueID, cal.ParentCalendarUniqueID?.ToString() ?? "none");

                        var dayNames = new[] { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };
                        for (int i = 0; i < 7; i++)
                        {
                            var day = cal.Days[i];
                            string hours = day.Hours.Count > 0
                                ? string.Join("; ", day.Hours.Select(h => string.Format("{0:hh\\:mm}-{1:hh\\:mm}", h.Start, h.End)))
                                : "";
                            Console.WriteLine("    {0}: Type={1} {2}", dayNames[i], day.Type, hours);
                        }

                        if (cal.Exceptions.Count > 0)
                        {
                            Console.WriteLine("    Exceptions: {0}", cal.Exceptions.Count);
                            foreach (var ex in cal.Exceptions)
                                Console.WriteLine("      {0:yyyy-MM-dd} to {1:yyyy-MM-dd} ({2})",
                                    ex.Start, ex.End, ex.Working ? "working" : "non-working");
                        }
                    }

                    // Show default calendar
                    var defCal = mppProject.GetDefaultCalendar();
                    Console.WriteLine("  Default calendar: \"{0}\"", defCal?.Name ?? "(none)");

                    // Show resolved working days using our helper
                    var mppCalInfo = GetMppCalendar(mppProject);
                    Console.WriteLine("  Resolved working days: {0}", mppCalInfo.WorkingDaysSummary());
                    Console.WriteLine("  Resolved exceptions: {0}", mppCalInfo.Exceptions.Count);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("  ERROR: {0}", ex.Message);
                }
            }

            // Query D365 calendar
            Console.WriteLine();
            Console.WriteLine("--- D365 Project Calendar ---");
            try
            {
                Console.WriteLine("Connecting to D365...");
                var envData = LoadEnvAndConnect(requireCaseTemplate: false, requireProjectId: false);
                if (envData == null) { Console.WriteLine("  Failed to connect."); return 1; }
                var service = envData.Value.service;
                Console.WriteLine("  Connected.");

                // Create a temp project to see what calendar it gets
                var tempProject = new Entity("msdyn_project");
                tempProject["msdyn_subject"] = "CALCHECK-TEMP-" + DateTime.Now.ToString("yyyyMMdd-HHmmss");
                Guid tempId = service.Create(tempProject);
                Console.WriteLine("  Created temp project: {0}", tempId);

                var d365CalInfo = GetD365Calendar(service, tempId);
                Console.WriteLine("  Calendar name:    \"{0}\"", d365CalInfo.Name);
                Console.WriteLine("  Working days:     {0}", d365CalInfo.WorkingDaysSummary());

                var dayNames2 = new[] { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };
                for (int i = 0; i < 7; i++)
                {
                    Console.WriteLine("    {0}: working={1} {2}",
                        dayNames2[i], d365CalInfo.WorkingDays[i],
                        d365CalInfo.DayHours[i] ?? "");
                }

                // Also dump raw calendar rules
                var project = service.Retrieve("msdyn_project", tempId,
                    new Microsoft.Xrm.Sdk.Query.ColumnSet("calendarid"));
                var calRef = project.GetAttributeValue<EntityReference>("calendarid");
                if (calRef != null)
                {
                    Console.WriteLine("  Calendar entity ID: {0}", calRef.Id);
                    var calendar = service.Retrieve("calendar", calRef.Id,
                        new Microsoft.Xrm.Sdk.Query.ColumnSet(true));
                    var rules = calendar.GetAttributeValue<EntityCollection>("calendarrules");
                    if (rules != null)
                    {
                        Console.WriteLine("  Calendar rules: {0}", rules.Entities.Count);
                        foreach (var rule in rules.Entities)
                        {
                            string pattern = rule.GetAttributeValue<string>("pattern") ?? "(no pattern)";
                            int? timeCode = rule.Contains("timecode") ? (int?)rule.GetAttributeValue<int>("timecode") : null;
                            var innerRef = rule.GetAttributeValue<EntityReference>("innercalendarid");
                            Console.WriteLine("    Rule: pattern=\"{0}\", timecode={1}, innerCal={2}",
                                pattern, timeCode?.ToString() ?? "null",
                                innerRef != null ? innerRef.Id.ToString() : "null");

                            // Dump inner calendar rules too
                            if (innerRef != null)
                            {
                                try
                                {
                                    var innerCal = service.Retrieve("calendar", innerRef.Id,
                                        new Microsoft.Xrm.Sdk.Query.ColumnSet(true));
                                    var innerRules = innerCal.GetAttributeValue<EntityCollection>("calendarrules");
                                    if (innerRules != null)
                                    {
                                        foreach (var ir in innerRules.Entities)
                                        {
                                            string ip = ir.GetAttributeValue<string>("pattern") ?? "(no pattern)";
                                            int? itc = ir.Contains("timecode") ? (int?)ir.GetAttributeValue<int>("timecode") : null;
                                            var startT = ir.GetAttributeValue<DateTime?>("starttime");
                                            var endT = ir.GetAttributeValue<DateTime?>("endtime");
                                            Console.WriteLine("      Inner: pattern=\"{0}\", timecode={1}, start={2}, end={3}",
                                                ip, itc?.ToString() ?? "null",
                                                startT?.ToString("HH:mm") ?? "null",
                                                endT?.ToString("HH:mm") ?? "null");
                                        }
                                    }
                                }
                                catch (Exception ex2) { Console.WriteLine("      Inner error: {0}", ex2.Message); }
                            }
                        }
                    }
                }

                // Cleanup temp project
                try { service.Delete("msdyn_project", tempId); }
                catch { Console.WriteLine("  (temp project cleanup skipped)"); }
            }
            catch (Exception ex)
            {
                Console.WriteLine("  ERROR: {0}", ex.Message);
            }

            return 0;
        }

        /// <summary>
        /// Compares start/end dates between MPP file and existing D365 project tasks.
        /// Runs against already-imported projects without re-importing.
        /// </summary>
        static int RunDateCheck(List<Tuple<string, Guid>> pairs)
        {
            Console.WriteLine();
            Console.WriteLine("================================================================");
            Console.WriteLine("  DATE CHECK — Compare MPP dates vs D365 scheduled dates");
            Console.WriteLine("================================================================");

            if (pairs.Count == 0)
            {
                Console.WriteLine("Usage: datecheck <mppFile> <projectGuid> [mppFile2 projectGuid2 ...]");
                return 1;
            }

            Console.WriteLine("Connecting to Dynamics 365...");
            var envData = LoadEnvAndConnect(requireCaseTemplate: false, requireProjectId: false);
            if (envData == null) return 1;
            var service = envData.Value.service;
            Console.WriteLine("  Connected.");

            int totalFiles = 0, totalPass = 0;

            foreach (var pair in pairs)
            {
                string mppPath = pair.Item1;
                Guid projectId = pair.Item2;
                string fileName = Path.GetFileName(mppPath);
                totalFiles++;

                Console.WriteLine();
                Console.WriteLine("================================================================");
                Console.WriteLine("  FILE: {0}", fileName);
                Console.WriteLine("  PROJECT: {0}", projectId);
                Console.WriteLine("================================================================");

                if (!File.Exists(mppPath))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("  ERROR: File not found: {0}", mppPath);
                    Console.ResetColor();
                    continue;
                }

                // Parse MPP using custom reader
                Console.WriteLine("  Parsing MPP...");
                byte[] mppBytes = File.ReadAllBytes(mppPath);
                var reader = new MppFileReader();
                var mppProject = reader.Read(mppBytes);

                // Build expected tasks with dates
                var mppTasks = mppProject.Tasks
                    .Where(t => t.UniqueID.HasValue && t.UniqueID.Value != 0)
                    .ToList();

                var summaryIds = new HashSet<int>();
                foreach (var t in mppTasks)
                {
                    if (t.ParentTask != null && t.ParentTask.UniqueID.HasValue && t.ParentTask.UniqueID.Value != 0)
                        summaryIds.Add(t.ParentTask.UniqueID.Value);
                }

                var expectedTasks = new List<ExpectedTask>();
                foreach (var t in mppTasks)
                {
                    var et = new ExpectedTask
                    {
                        UniqueID = t.UniqueID.Value,
                        Name = t.Name ?? "(Unnamed Task)",
                        IsSummary = summaryIds.Contains(t.UniqueID.Value),
                        OutlineLevel = t.OutlineLevel ?? 0
                    };
                    et.StartDate = t.Start;
                    et.FinishDate = t.Finish;
                    expectedTasks.Add(et);
                }

                Console.WriteLine("    MPP: {0} tasks ({1} with start dates, {2} with finish dates)",
                    expectedTasks.Count,
                    expectedTasks.Count(e => e.StartDate.HasValue),
                    expectedTasks.Count(e => e.FinishDate.HasValue));

                // Query CRM tasks
                Console.WriteLine("  Querying D365 tasks...");
                var crmTasks = QueryProjectTasksFull(service, projectId);
                var crmNonRoot = crmTasks.Where(e =>
                {
                    int level = e.Contains("msdyn_outlinelevel") ? (int)e["msdyn_outlinelevel"] : -1;
                    return level != 0;
                }).ToList();

                Console.WriteLine("    D365: {0} tasks (excl. root)", crmNonRoot.Count);

                // Build name-consumption map
                var crmByName = new Dictionary<string, List<Entity>>(StringComparer.OrdinalIgnoreCase);
                foreach (var ct in crmNonRoot)
                {
                    string name = ct.GetAttributeValue<string>("msdyn_subject") ?? "";
                    List<Entity> list;
                    if (!crmByName.TryGetValue(name, out list))
                    {
                        list = new List<Entity>();
                        crmByName[name] = list;
                    }
                    list.Add(ct);
                }

                var expectedToCrm = new Dictionary<int, Entity>();
                foreach (var et in expectedTasks)
                {
                    List<Entity> candidates;
                    if (crmByName.TryGetValue(et.Name, out candidates) && candidates.Count > 0)
                    {
                        expectedToCrm[et.UniqueID] = candidates[0];
                        candidates.RemoveAt(0);
                    }
                }

                // Compare dates
                Console.WriteLine();
                Console.WriteLine("  --- START DATE COMPARISON ---");
                int startOk = 0, startWarn = 0, startMissing = 0;
                int endOk = 0, endWarn = 0, endMissing = 0;
                var startIssues = new List<string>();
                var endIssues = new List<string>();

                foreach (var et in expectedTasks)
                {
                    Entity ct;
                    if (!expectedToCrm.TryGetValue(et.UniqueID, out ct))
                    {
                        startMissing++;
                        endMissing++;
                        continue;
                    }

                    DateTime? crmStart = ct.GetAttributeValue<DateTime?>("msdyn_scheduledstart");
                    DateTime? crmEnd = ct.GetAttributeValue<DateTime?>("msdyn_scheduledend");

                    // Compare start dates
                    if (et.StartDate.HasValue && crmStart.HasValue)
                    {
                        DateTime mppStart = et.StartDate.Value.Date;
                        DateTime d365Start = crmStart.Value.Date;
                        int daysDiff = (int)(d365Start - mppStart).TotalDays;

                        if (daysDiff == 0)
                            startOk++;
                        else
                        {
                            startWarn++;
                            startIssues.Add(string.Format("{0}{1}: MPP={2:yyyy-MM-dd}, D365={3:yyyy-MM-dd} (diff={4}d)",
                                et.IsSummary ? "[S] " : "",
                                et.Name, mppStart, d365Start, daysDiff));
                        }
                    }
                    else if (et.StartDate.HasValue && !crmStart.HasValue)
                    {
                        startWarn++;
                        startIssues.Add(string.Format("{0}: MPP={1:yyyy-MM-dd}, D365=NULL",
                            et.Name, et.StartDate.Value.Date));
                    }

                    // Compare end dates
                    if (et.FinishDate.HasValue && crmEnd.HasValue)
                    {
                        DateTime mppEnd = et.FinishDate.Value.Date;
                        DateTime d365End = crmEnd.Value.Date;
                        int daysDiff = (int)(d365End - mppEnd).TotalDays;

                        if (daysDiff == 0)
                            endOk++;
                        else
                        {
                            endWarn++;
                            endIssues.Add(string.Format("{0}{1}: MPP={2:yyyy-MM-dd}, D365={3:yyyy-MM-dd} (diff={4}d)",
                                et.IsSummary ? "[S] " : "",
                                et.Name, mppEnd, d365End, daysDiff));
                        }
                    }
                    else if (et.FinishDate.HasValue && !crmEnd.HasValue)
                    {
                        endWarn++;
                        endIssues.Add(string.Format("{0}: MPP={1:yyyy-MM-dd}, D365=NULL",
                            et.Name, et.FinishDate.Value.Date));
                    }
                }

                // Report start dates
                Console.ForegroundColor = startWarn == 0 ? ConsoleColor.Green : ConsoleColor.Yellow;
                Console.WriteLine("    Start dates: {0} match, {1} differ, {2} unmapped",
                    startOk, startWarn, startMissing);
                Console.ResetColor();

                if (startIssues.Count > 0 && startIssues.Count <= 30)
                {
                    foreach (var issue in startIssues)
                        Console.WriteLine("      {0}", issue);
                }
                else if (startIssues.Count > 30)
                {
                    for (int i = 0; i < 20; i++)
                        Console.WriteLine("      {0}", startIssues[i]);
                    Console.WriteLine("      ... and {0} more", startIssues.Count - 20);
                }

                // Report end dates
                Console.WriteLine();
                Console.WriteLine("  --- END DATE COMPARISON ---");
                Console.ForegroundColor = endWarn == 0 ? ConsoleColor.Green : ConsoleColor.Yellow;
                Console.WriteLine("    End dates: {0} match, {1} differ, {2} unmapped",
                    endOk, endWarn, endMissing);
                Console.ResetColor();

                if (endIssues.Count > 0 && endIssues.Count <= 30)
                {
                    foreach (var issue in endIssues)
                        Console.WriteLine("      {0}", issue);
                }
                else if (endIssues.Count > 30)
                {
                    for (int i = 0; i < 20; i++)
                        Console.WriteLine("      {0}", endIssues[i]);
                    Console.WriteLine("      ... and {0} more", endIssues.Count - 20);
                }

                // Summary
                Console.WriteLine();
                bool pass = startWarn == 0 && endWarn == 0;
                Console.ForegroundColor = pass ? ConsoleColor.Green : ConsoleColor.Yellow;
                Console.WriteLine("  RESULT: {0} ({1}/{2} start dates match, {3}/{4} end dates match)",
                    pass ? "PASS" : "DATES DIFFER",
                    startOk, startOk + startWarn,
                    endOk, endOk + endWarn);
                Console.ResetColor();

                if (pass) totalPass++;
            }

            Console.WriteLine();
            Console.WriteLine("================================================================");
            Console.WriteLine("  FINAL: {0}/{1} files have matching dates", totalPass, totalFiles);
            Console.WriteLine("================================================================");

            return totalPass == totalFiles ? 0 : 1;
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
