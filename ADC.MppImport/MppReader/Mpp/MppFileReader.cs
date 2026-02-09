using System;
using System.Collections.Generic;
using System.IO;
using ADC.MppImport.MppReader.Model;
using ADC.MppImport.Ole2;

namespace ADC.MppImport.MppReader.Mpp
{
    /// <summary>
    /// Main entry point for reading MPP files. Uses OpenMcdf to read OLE2 compound documents
    /// and dispatches to the appropriate version-specific reader.
    /// Ported from org.mpxj.mpp.MPPReader
    /// </summary>
    public class MppFileReader
    {
        public bool UseRawTimephasedData { get; set; }
        public bool ReadPresentationData { get; set; } = true;
        public bool ReadPropertiesOnly { get; set; }
        public bool RespectPasswordProtection { get; set; } = true;
        public string ReadPassword { get; set; }

        /// <summary>
        /// Read an MPP file from the specified file path.
        /// </summary>
        public ProjectFile Read(string filePath)
        {
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                return Read(fs);
            }
        }

        /// <summary>
        /// Read an MPP file from a stream.
        /// </summary>
        public ProjectFile Read(Stream stream)
        {
            byte[] data;
            using (var ms = new MemoryStream())
            {
                stream.CopyTo(ms);
                data = ms.ToArray();
            }
            return Read(data);
        }

        /// <summary>
        /// Read an MPP file from a byte array.
        /// </summary>
        public ProjectFile Read(byte[] data)
        {
            using (var cf = new CompoundFile(new MemoryStream(data)))
            {
                return Read(cf);
            }
        }

        /// <summary>
        /// Read an MPP file from an already-opened CompoundFile.
        /// </summary>
        public ProjectFile Read(CompoundFile cf)
        {
            var projectFile = new ProjectFile();
            var root = cf.RootStorage;

            // Read the CompObj to determine the file format
            byte[] compObjData = GetStreamData(root, "\u0001CompObj");
            if (compObjData == null)
                throw new MppReaderException("Cannot find CompObj stream - not a valid MPP file");

            var compObj = new CompObj(compObjData);
            var properties = projectFile.ProjectProperties;
            properties.FullApplicationName = compObj.ApplicationName;
            properties.ApplicationVersion = compObj.ApplicationVersion;

            string format = compObj.FileFormat;
            if (string.IsNullOrEmpty(format))
                throw new MppReaderException("Cannot determine file format from CompObj");

            // Dispatch to the appropriate version reader
            IMppVariantReader reader = GetVariantReader(format);
            if (reader == null)
                throw new MppReaderException("Unsupported MPP file format: " + format);

            reader.Process(this, projectFile, cf, root);

            // Post-processing
            projectFile.ResolveReferences();

            // Set analytics
            string projectFilePath = properties.ProjectFilePath;
            if (projectFilePath != null && projectFilePath.StartsWith("<>\\"))
                properties.FileApplication = "Microsoft Project Server";
            else
                properties.FileApplication = "Microsoft";
            properties.FileType = "MPP";

            return projectFile;
        }

        private IMppVariantReader GetVariantReader(string format)
        {
            switch (format)
            {
                case "MSProject.MPP14":
                case "MSProject.MPT14":
                case "MSProject.GLOBAL14":
                    return new Mpp14Reader();

                case "MSProject.MPP12":
                case "MSProject.MPT12":
                case "MSProject.GLOBAL12":
                    return new Mpp12Reader();

                case "MSProject.MPP9":
                case "MSProject.MPT9":
                case "MSProject.GLOBAL9":
                    return new Mpp9Reader();

                case "MSProject.MPP8":
                case "MSProject.MPT8":
                    return new Mpp8Reader();

                default:
                    return null;
            }
        }

        /// <summary>
        /// Utility to read a stream from the compound file, returning null if not found.
        /// Tries exact name first, then with/without special prefixes.
        /// </summary>
        internal static byte[] GetStreamData(CFStorage storage, string name)
        {
            // Try exact name
            try
            {
                var stream = storage.GetStream(name);
                return stream.GetData();
            }
            catch { }

            // Try stripping \x01 or \x05 prefix and looking up without it
            if (name.Length > 0 && (name[0] == '\x01' || name[0] == '\x05'))
            {
                string stripped = name.Substring(1);
                try
                {
                    var stream = storage.GetStream(stripped);
                    return stream.GetData();
                }
                catch { }
            }

            // Try adding \x01 prefix if not present
            if (name.Length > 0 && name[0] != '\x01' && name[0] != '\x05')
            {
                try
                {
                    var stream = storage.GetStream("\u0001" + name);
                    return stream.GetData();
                }
                catch { }
            }

            // Fallback: visit entries and match by suffix
            byte[] result = null;
            try
            {
                string matchName = name.TrimStart('\x01', '\x05');
                storage.VisitEntries(item =>
                {
                    if (!item.IsStorage && item.Name.EndsWith(matchName, StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            result = ((CFStream)item).GetData();
                        }
                        catch { }
                    }
                }, false);
            }
            catch { }

            return result;
        }

        /// <summary>
        /// Utility to get a sub-storage, returning null if not found.
        /// </summary>
        internal static CFStorage GetStorage(CFStorage parent, string name)
        {
            try
            {
                return parent.GetStorage(name);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Check if a storage contains a named entry.
        /// </summary>
        internal static bool HasEntry(CFStorage storage, string name)
        {
            try
            {
                storage.GetStream(name);
                return true;
            }
            catch
            {
                try
                {
                    storage.GetStorage(name);
                    return true;
                }
                catch
                {
                    return false;
                }
            }
        }
    }

    /// <summary>
    /// Exception thrown when reading MPP files fails.
    /// </summary>
    public class MppReaderException : Exception
    {
        public MppReaderException(string message) : base(message) { }
        public MppReaderException(string message, Exception innerException) : base(message, innerException) { }
    }
}
