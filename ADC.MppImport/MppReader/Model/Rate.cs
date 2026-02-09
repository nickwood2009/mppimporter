namespace ADC.MppImport.MppReader.Model
{
    /// <summary>
    /// Represents a rate (amount per time unit).
    /// Ported from org.mpxj.Rate
    /// </summary>
    public class Rate
    {
        public double Amount { get; }
        public TimeUnit Units { get; }

        public Rate(double amount, TimeUnit units)
        {
            Amount = amount;
            Units = units;
        }

        public override string ToString()
        {
            return $"{Amount}/{Units}";
        }
    }
}
