using System;

namespace DXPlus
{
    public class FormattedText : IComparable
    {
        public int Index { get; set; }
        public string Text { get; set; }
        public Formatting Formatting { get; set; }

        public int CompareTo(object obj)
        {
            var other = (FormattedText)obj;
            return other.Formatting == null || Formatting == null
                ? -1
                : Formatting.CompareTo(other.Formatting);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(this, obj))
            {
                return true;
            }

            if (obj is null)
            {
                return false;
            }

            return CompareTo(obj) == 0;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Formatting, Index, Text);
        }

        public static bool operator ==(FormattedText left, FormattedText right)
        {
            if (left is null)
            {
                return right is null;
            }

            return left.Equals(right);
        }

        public static bool operator !=(FormattedText left, FormattedText right)
        {
            return !(left == right);
        }

        public static bool operator <(FormattedText left, FormattedText right)
        {
            return left is null
                ? right is object
                : left.CompareTo(right) < 0;
        }

        public static bool operator <=(FormattedText left, FormattedText right)
        {
            return left is null || left.CompareTo(right) <= 0;
        }

        public static bool operator >(FormattedText left, FormattedText right)
        {
            return left is object && left.CompareTo(right) > 0;
        }

        public static bool operator >=(FormattedText left, FormattedText right)
        {
            return left is null ? right is null : left.CompareTo(right) >= 0;
        }
    }
}
