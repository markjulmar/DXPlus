using System;

namespace DXPlus
{
    public class FormattedText : IComparable
    {
        public int index;
        public string text;
        public Formatting formatting;

        public int CompareTo(object obj)
        {
            FormattedText other = (FormattedText) obj;
            FormattedText tf = this;

            return other.formatting == null || tf.formatting == null
                ? -1
                : tf.formatting.CompareTo(other.formatting);
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
            return HashCode.Combine(formatting);
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
