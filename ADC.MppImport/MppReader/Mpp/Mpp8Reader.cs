using System;
using ADC.MppImport.MppReader.Model;
using ADC.MppImport.Ole2;

namespace ADC.MppImport.MppReader.Mpp
{
    /// <summary>
    /// Reads MPP8 files (MS Project 98/2000).
    /// Stub implementation - MPP8 is a significantly different format.
    /// TODO: Full implementation if needed.
    /// </summary>
    internal class Mpp8Reader : IMppVariantReader
    {
        public void Process(MppFileReader reader, ProjectFile file, CompoundFile cf, CFStorage root)
        {
            throw new MppReaderException(
                "MPP8 (MS Project 98/2000) format is not yet supported. " +
                "Please save the file in a newer format (MS Project 2003 or later).");
        }
    }
}
