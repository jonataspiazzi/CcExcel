using CcExcel.Messages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace CcExcel
{
    public struct BaseAZ : IComparable<BaseAZ>
    {
        #region Fields

        private static readonly Regex _regexAz = new Regex(@"^[A-Za-z]+$", RegexOptions.Compiled);
        private readonly uint _value;

        #endregion

        #region Constructors

        public BaseAZ(uint value)
        {
            _value = value;
        }

        #endregion

        #region Methods

        public static BaseAZ Parse(string value)
        {
            if (!_regexAz.IsMatch(value))
            {
                throw new FormatException(string.Format(Texts.TheValue0CannotBeConvertedIn1, value, nameof(BaseAZ)));
            }

            var letters = value.Normalize().ToUpper().Reverse();

            uint plus = 1;
            uint total = 0;

            foreach (var letter in letters)
            {
                total += (uint)(plus * (letter - 'A' + 1));

                plus *= 26;
            }

            return new BaseAZ(total);
        }

        public int CompareTo(BaseAZ other)
        {
            return _value.CompareTo(other._value);
        }

        public override string ToString()
        {
            var array = new LinkedList<uint>();
            var myNumber = _value;

            while (myNumber > 26)
            {
                var value = myNumber % 26;

                if (value == 0)
                {
                    myNumber = myNumber / 26 - 1;
                    array.AddFirst(26);
                }
                else
                {
                    myNumber /= 26;
                    array.AddFirst(value);
                }
            }

            if (myNumber > 0)
            {
                array.AddFirst(myNumber);
            }

            return new string(array.Select(s => (char)('A' + s - 1)).ToArray());
        }

        public override int GetHashCode()
        {
            return _value.GetHashCode();
        }

        public override bool Equals(object obj)
        {
            return obj is BaseAZ && this == (BaseAZ)obj;
        }

        #endregion

        #region Operators

        #region Explicit Operators

        public static explicit operator BaseAZ(uint value)
        {
            return new BaseAZ(value);
        }

        public static explicit operator uint(BaseAZ value)
        {
            return value._value;
        }

        #endregion

        #region Comparision Operators

        public static bool operator ==(BaseAZ left, BaseAZ right)
        {
            return left._value == right._value;
        }

        public static bool operator !=(BaseAZ left, BaseAZ right)
        {
            return !(left == right);
        }

        public static bool operator >(BaseAZ left, BaseAZ right)
        {
            return left._value > right._value;
        }

        public static bool operator <(BaseAZ left, BaseAZ right)
        {
            return left._value < right._value;
        }

        public static bool operator >=(BaseAZ left, BaseAZ right)
        {
            return left > right || left == right;
        }

        public static bool operator <=(BaseAZ left, BaseAZ right)
        {
            return left < right || left == right;
        }

        #endregion

        #region Sum Operators

        public static BaseAZ operator +(BaseAZ left, BaseAZ right)
        {
            return new BaseAZ(left._value + right._value);
        }

        public static BaseAZ operator +(BaseAZ left, uint right)
        {
            return new BaseAZ(left._value + right);
        }

        public static BaseAZ operator +(uint left, BaseAZ right)
        {
            return new BaseAZ(right._value + left);
        }

        public static BaseAZ operator +(BaseAZ left, int right)
        {
            return new BaseAZ((uint)(left._value + right));
        }

        public static BaseAZ operator +(int left, BaseAZ right)
        {
            return new BaseAZ((uint)(right._value + left));
        }

        #endregion

        #region Subtraction Operators

        public static BaseAZ operator -(BaseAZ left, BaseAZ right)
        {
            return new BaseAZ(left._value - right._value);
        }

        public static BaseAZ operator -(BaseAZ left, uint right)
        {
            return new BaseAZ(left._value - right);
        }

        public static BaseAZ operator -(uint left, BaseAZ right)
        {
            return new BaseAZ(left - right._value);
        }

        public static BaseAZ operator -(BaseAZ left, int right)
        {
            return new BaseAZ((uint)(left._value - right));
        }

        public static BaseAZ operator -(int left, BaseAZ right)
        {
            return new BaseAZ((uint)(left - right._value));
        }

        #endregion

        #region Increment Operators

        public static BaseAZ operator ++(BaseAZ left)
        {
            return new BaseAZ(left._value + 1);
        }

        public static BaseAZ operator --(BaseAZ left)
        {
            return new BaseAZ(left._value - 1);
        }

        #endregion

        #endregion
    }
}
