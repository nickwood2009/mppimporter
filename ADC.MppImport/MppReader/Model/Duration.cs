namespace ADC.MppImport.MppReader.Model
{
    /// <summary>
    /// Represents a duration with value and time units.
    /// Ported from org.mpxj.Duration
    /// </summary>
    public class Duration
    {
        public double Value { get; }
        public TimeUnit Units { get; }

        private Duration(double value, TimeUnit units)
        {
            Value = value;
            Units = units;
        }

        public static Duration GetInstance(double value, TimeUnit units)
        {
            return new Duration(value, units);
        }

        public override string ToString()
        {
            return $"{Value} {Units}";
        }
    }
}
