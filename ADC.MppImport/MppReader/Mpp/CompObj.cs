using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using ADC.MppImport.MppReader.Common;

namespace ADC.MppImport.MppReader.Mpp
{
    /// <summary>
    /// Reads the CompObj block from an MPP file to determine the file format.
    /// Ported from org.mpxj.mpp.CompObj
    /// </summary>
    internal class CompObj
    {
        public string ApplicationName { get; }
        public int? ApplicationVersion { get; }
        public string ApplicationID { get; private set; }
        public string FileFormat { get; private set; }

        private static readonly Regex VersionPattern = new Regex(@"Microsoft.Project.(\d+)\.0");

        public CompObj(byte[] data)
        {
            using (var ms = new MemoryStream(data))
            using (var reader = new BinaryReader(ms))
            {
                // Skip first 28 bytes
                reader.ReadBytes(28);

                int length = reader.ReadInt32();
                byte[] nameBytes = reader.ReadBytes(length);
                ApplicationName = Encoding.ASCII.GetString(nameBytes, 0, length - 1);

                var match = VersionPattern.Match(ApplicationName);
                if (match.Success)
                {
                    ApplicationVersion = int.Parse(match.Groups[1].Value);
                }

                if (ApplicationName == "Microsoft Project 4.0")
                {
                    FileFormat = "MSProject.MPP4";
                    ApplicationID = "MSProject.Project.4";
                }
                else
                {
                    if (ms.Position < ms.Length)
                    {
                        length = reader.ReadInt32();
                        if (length > 0 && ms.Position + length <= ms.Length)
                        {
                            byte[] formatBytes = reader.ReadBytes(length);
                            FileFormat = Encoding.ASCII.GetString(formatBytes, 0, length - 1);

                            if (ms.Position + 4 <= ms.Length)
                            {
                                length = reader.ReadInt32();
                                if (length > 0 && ms.Position + length <= ms.Length)
                                {
                                    byte[] idBytes = reader.ReadBytes(length);
                                    ApplicationID = Encoding.ASCII.GetString(idBytes, 0, length - 1);
                                }
                            }
                        }
                    }
                }
            }
        }

        public override string ToString()
        {
            return $"[CompObj applicationName={ApplicationName} applicationID={ApplicationID} fileFormat={FileFormat}]";
        }
    }
}
