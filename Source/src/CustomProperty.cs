using System;

namespace DXPlus
{
    public class CustomProperty
    {
        private const string LPWSTR = "lpwstr";
        private const string I4 = "i4";
        private const string R8 = "r8";
        private const string FILETIME = "filetime";
        private const string BOOL = "bool";

        /// <summary>
        /// The name of this CustomProperty.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// The value of this CustomProperty.
        /// </summary>
        public object Value { get; }

        internal string Type { get; }

        internal CustomProperty(string name, string type, string value)
        {
            this.Name = name;
            this.Type = type;
            this.Value = type switch
            {
                LPWSTR => value,
                I4 => int.Parse(value, System.Globalization.CultureInfo.InvariantCulture),
                R8 => double.Parse(value, System.Globalization.CultureInfo.InvariantCulture),
                FILETIME => DateTime.Parse(value, System.Globalization.CultureInfo.InvariantCulture),
                BOOL => bool.Parse(value),
                _ => throw new Exception($"Invalid type {type} used for {nameof(CustomProperty)}")
            };
        }

        private CustomProperty(string name, string type, object value)
        {
            this.Name = name;
            this.Type = type;
            this.Value = value;
        }

        /// <summary>
        /// Create a new CustomProperty to hold a string.
        /// </summary>
        /// <param name="name">The name of this CustomProperty.</param>
        /// <param name="value">The value of this CustomProperty.</param>
        public CustomProperty(string name, string value) : this(name, LPWSTR, value as object)
        {
        }

        /// <summary>
        /// Create a new CustomProperty to hold an int.
        /// </summary>
        /// <param name="name">The name of this CustomProperty.</param>
        /// <param name="value">The value of this CustomProperty.</param>
        public CustomProperty(string name, int value) : this(name, I4, value)
        {
        }

        /// <summary>
        /// Create a new CustomProperty to hold a double.
        /// </summary>
        /// <param name="name">The name of this CustomProperty.</param>
        /// <param name="value">The value of this CustomProperty.</param>
        public CustomProperty(string name, double value) : this(name, R8, value)
        {
        }

        /// <summary>
        /// Create a new CustomProperty to hold a DateTime.
        /// </summary>
        /// <param name="name">The name of this CustomProperty.</param>
        /// <param name="value">The value of this CustomProperty.</param>
        public CustomProperty(string name, DateTime value) : this(name, FILETIME, value.ToUniversalTime())
        {
        }

        /// <summary>
        /// Create a new CustomProperty to hold a bool.
        /// </summary>
        /// <param name="name">The name of this CustomProperty.</param>
        /// <param name="value">The value of this CustomProperty.</param>
        public CustomProperty(string name, bool value) : this(name, BOOL, value)
        { 
        }
    }
}
