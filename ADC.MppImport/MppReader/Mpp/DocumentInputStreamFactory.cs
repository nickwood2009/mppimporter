namespace ADC.MppImport.MppReader.Mpp
{
    /// <summary>
    /// Factory to handle encrypted/unencrypted document streams.
    /// Ported from org.mpxj.mpp.DocumentInputStreamFactory
    /// </summary>
    internal class DocumentInputStreamFactory
    {
        public bool Encrypted { get; }
        public byte EncryptionCode { get; }

        public DocumentInputStreamFactory(Props props)
        {
            Encrypted = props.GetByte(Props.PASSWORD_FLAG) != 0;
            byte code = props.GetByte(Props.ENCRYPTION_CODE);
            EncryptionCode = (byte)(code == 0x00 ? 0x00 : (0xFF - code));
        }

        /// <summary>
        /// Get the data for the given entry, decrypting if necessary.
        /// </summary>
        public byte[] GetData(byte[] rawData)
        {
            if (Encrypted && rawData != null)
            {
                byte[] result = new byte[rawData.Length];
                System.Array.Copy(rawData, result, rawData.Length);
                MppUtility.DecodeBuffer(result, EncryptionCode);
                return result;
            }
            return rawData;
        }
    }
}
