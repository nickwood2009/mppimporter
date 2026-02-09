using System;
using System.IO;
using ADC.MppImport.MppReader.Mpp;
using ADC.MppImport.MppReader.Model;

namespace ADC.MppImport.Test
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: ADC.MppImport.Test <mpp-file>");
                return 1;
            }

            string filePath = args[0];
            if (!File.Exists(filePath))
            {
                Console.WriteLine("File not found: " + filePath);
                return 1;
            }

            try
            {
                byte[] data = File.ReadAllBytes(filePath);
                Console.WriteLine("File: {0} ({1} bytes)", Path.GetFileName(filePath), data.Length);

                var reader = new MppFileReader();
                ProjectFile project = reader.Read(data);

                Console.WriteLine("MPP Type:     {0}", project.ProjectProperties.MppFileType);
                Console.WriteLine("Tasks:        {0}", project.Tasks.Count);
                Console.WriteLine("Resources:    {0}", project.Resources.Count);
                Console.WriteLine("Assignments:  {0}", project.Assignments.Count);
                Console.WriteLine("Calendars:    {0}", project.Calendars.Count);
                Console.WriteLine("Errors:       {0}", project.IgnoredErrors.Count);

                Console.WriteLine();
                Console.WriteLine("--- Tasks (first 10) ---");
                int shown = 0;
                foreach (var t in project.Tasks)
                {
                    if (shown++ >= 10) break;
                    Console.WriteLine("  [{0}] {1}  Start={2:d}  Dur={3}  %={4}",
                        t.UniqueID, t.Name,
                        t.Start, t.Duration, t.PercentComplete);
                }

                if (project.Resources.Count > 0)
                {
                    Console.WriteLine();
                    Console.WriteLine("--- Resources (first 10) ---");
                    shown = 0;
                    foreach (var r in project.Resources)
                    {
                        if (shown++ >= 10) break;
                        Console.WriteLine("  [{0}] {1}", r.UniqueID, r.Name);
                    }
                }

                Console.WriteLine();
                Console.WriteLine("=== PASS ===");
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine("=== FAIL ===");
                Console.WriteLine(ex.ToString());
                return 2;
            }
        }
    }
}
