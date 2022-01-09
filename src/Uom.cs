namespace DXPlus
{
    /// <summary>
    /// In order to avoid floating point calculations and still maintain high precision DOCX uses some odd measurement units - Dxa, Half-points, and EMUs.
    /// This class represents the various measurements used by docx files.
    /// Dxa is the main unit and states values in twentieths of a point (1/1440 of an inch).
    /// Half-points are the main unit for font sizes, calculate the value by multiplying the ptSize by 2.
    /// Emu is English Metric Units. EMUs are equivalent to 1/360,000 of a centimeter, 1/914,400 inches or 1/12,700 points.
    /// </summary>
    public class Uom
    {
        public const double DxaConversion = 20.0;
        public const double InchConversion = 1440.0;
        public const double EmuConversion = 914400.0 / 96.0; // 91440/inch, 96px/inch

        /// <summary>
        /// Internally we represent this as dxa (1/20 of a 
        /// </summary>
        private readonly double value;

        internal Uom(double value)
        {
            this.value = value;
        }

        public static Uom FromInches(double inches) => new(inches * InchConversion);
        public static Uom FromPoints(double pt) => new(pt * DxaConversion);
        public static Uom FromDxa(double value) => new(value);
        public static Uom FromHalfPoints(double hpt) => FromPoints(hpt/2.0);

        public double Inches => value / InchConversion;
        public double Points => value / DxaConversion;
        public double HalfPoints => (value / DxaConversion) * 2;
        public double Dxa => value;

        /// <summary>
        /// Implicit conversion to double
        /// </summary>
        /// <param name="dxa"></param>
        public static implicit operator double(Uom dxa)
        {
            return dxa.value;
        }
    }
}
