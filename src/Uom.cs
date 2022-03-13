namespace DXPlus
{
    /// <summary>
    /// In order to avoid floating point calculations and still maintain high precision DOCX uses some odd measurement units - Dxa, Half-points, and EMUs.
    /// This class represents the various measurements used by docx files.
    /// Dxa is the main unit and states values in twentieths of a point (1/1440 of an inch).
    /// Half-points are the main unit for font sizes, calculate the value by multiplying the ptSize by 2.
    /// Emu is English Metric Units. EMUs are equivalent to 1/360,000 of a centimeter, 1/914,400 inches or 1/12,700 points.
    /// </summary>
    public sealed class Uom
    {
        /// <summary>
        /// Conversion ratio for Dxa to points
        /// </summary>
        public const double DxaConversion = 20.0;

        /// <summary>
        /// Conversion ratio for inches to points
        /// </summary>
        public const double InchConversion = 1440.0;

        /// <summary>
        /// Conversion ratio for EMU to points
        /// </summary>
        public const double EmuConversion = 914400.0 / 96.0; // 91440/inch, 96px/inch

        /// <summary>
        /// Internally we represent this as dxa (1/20 of a 
        /// </summary>
        private readonly double value;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="value"></param>
        internal Uom(double value)
        {
            this.value = value;
        }

        /// <summary>
        /// Convert from inches
        /// </summary>
        /// <param name="inches">Inches</param>
        /// <returns>Uom</returns>
        public static Uom FromInches(double inches) => new(inches * InchConversion);

        /// <summary>
        /// Convert from points
        /// </summary>
        /// <param name="pt">Points</param>
        /// <returns>Uom</returns>
        public static Uom FromPoints(double pt) => new(pt * DxaConversion);

        /// <summary>
        /// Convert from Dxa
        /// </summary>
        /// <param name="value">DXA value</param>
        /// <returns>Uom</returns>
        public static Uom FromDxa(double value) => new(value);

        /// <summary>
        /// Convert from half-points
        /// </summary>
        /// <param name="hpt">Half-points</param>
        /// <returns>Uom</returns>
        public static Uom FromHalfPoints(double hpt) => FromPoints(hpt/2.0);

        /// <summary>
        /// To inches
        /// </summary>
        public double Inches => value / InchConversion;

        /// <summary>
        /// To points
        /// </summary>
        public double Points => value / DxaConversion;

        /// <summary>
        /// To half points
        /// </summary>
        public double HalfPoints => (value / DxaConversion) * 2;

        /// <summary>
        /// To dxa
        /// </summary>
        public double Dxa => value;

        /// <summary>
        /// Implicit conversion to double
        /// </summary>
        /// <param name="dxa"></param>
        public static implicit operator double(Uom dxa) => dxa.value;

        /// <summary>
        /// Implicit conversion from double
        /// </summary>
        /// <param name="value"></param>
        public static implicit operator Uom(double value) => new(value);
    }
}
