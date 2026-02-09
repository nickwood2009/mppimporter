using ADC.MppImport.MppReader.Model;
using ADC.MppImport.Ole2;

namespace ADC.MppImport.MppReader.Mpp
{
    /// <summary>
    /// Interface for version-specific MPP file readers.
    /// Ported from org.mpxj.mpp.MPPVariantReader
    /// </summary>
    internal interface IMppVariantReader
    {
        void Process(MppFileReader reader, ProjectFile file, CompoundFile cf, CFStorage root);
    }
}
