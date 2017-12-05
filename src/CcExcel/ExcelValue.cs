using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace CcExcel
{
    public class ExcelValue
    {
        #region Non Public

        private string _value;
        private bool _useStringTable;
        private static readonly CultureInfo _enUS = new CultureInfo("en-US");

        internal ExcelValue(string value, bool useStringTable)
        {
            _value = value;
            _useStringTable = useStringTable;
        }

        #endregion

        #region Boolean

        public static implicit operator ExcelValue(bool value)
        {
            throw new NotImplementedException();
        }

        public static implicit operator ExcelValue(bool? value)
        {
            throw new NotImplementedException();
        }

        public static implicit operator bool(ExcelValue value)
        {
            throw new NotImplementedException();
        }

        public static implicit operator bool? (ExcelValue value)
        {
            throw new NotImplementedException();
        }

        #endregion

        #region DateTime

        public static implicit operator ExcelValue(DateTime value)
        {
            return new ExcelValue(value.ToString(), false);
        }

        public static implicit operator ExcelValue(DateTime? value)
        {
            return new ExcelValue(value?.ToString(), false);
        }

        public static implicit operator DateTime(ExcelValue value)
        {
            return DateTime.Parse(value._value, _enUS);
        }

        public static implicit operator DateTime? (ExcelValue value)
        {
            return value._value != null ? DateTime.Parse(value._value, _enUS) : (DateTime?)null;
        }

        #endregion

        #region TimeSpan

        public static implicit operator ExcelValue(TimeSpan value)
        {
            return new ExcelValue(value.ToString(), false);
        }

        public static implicit operator ExcelValue(TimeSpan? value)
        {
            return new ExcelValue(value?.ToString(), false);
        }

        public static implicit operator TimeSpan(ExcelValue value)
        {
            return TimeSpan.Parse(value._value, _enUS);
        }

        public static implicit operator TimeSpan? (ExcelValue value)
        {
            return value._value != null ? TimeSpan.Parse(value._value, _enUS) : (TimeSpan?)null;
        }

        #endregion

        #region Single

        public static implicit operator ExcelValue(float value)
        {
            return new ExcelValue(value.ToString(), false);
        }

        public static implicit operator ExcelValue(float? value)
        {
            return new ExcelValue(value?.ToString(), false);
        }

        public static implicit operator float(ExcelValue value)
        {
            return float.Parse(value._value, _enUS);
        }

        public static implicit operator float? (ExcelValue value)
        {
            return value._value != null ? float.Parse(value._value, _enUS) : (float?)null;
        }

        #endregion

        #region Double

        public static implicit operator ExcelValue(double value)
        {
            return new ExcelValue(value.ToString(), false);
        }

        public static implicit operator ExcelValue(double? value)
        {
            return new ExcelValue(value?.ToString(), false);
        }

        public static implicit operator double(ExcelValue value)
        {
            return double.Parse(value._value, _enUS);
        }

        public static implicit operator double? (ExcelValue value)
        {
            return value._value != null ? double.Parse(value._value, _enUS) : (double?)null;
        }

        #endregion

        #region Decimal

        public static implicit operator ExcelValue(decimal value)
        {
            return new ExcelValue(value.ToString(), false);
        }

        public static implicit operator ExcelValue(decimal? value)
        {
            return new ExcelValue(value?.ToString(), false);
        }

        public static implicit operator decimal(ExcelValue value)
        {
            return decimal.Parse(value._value, _enUS);
        }

        public static implicit operator decimal? (ExcelValue value)
        {
            return value._value != null ? decimal.Parse(value._value, _enUS) : (decimal?)null;
        }

        #endregion

        #region SByte

        public static implicit operator ExcelValue(sbyte value)
        {
            return new ExcelValue(value.ToString(), false);
        }

        public static implicit operator ExcelValue(sbyte? value)
        {
            return new ExcelValue(value?.ToString(), false);
        }

        public static implicit operator sbyte(ExcelValue value)
        {
            return sbyte.Parse(value._value, _enUS);
        }

        public static implicit operator sbyte? (ExcelValue value)
        {
            return value._value != null ? sbyte.Parse(value._value, _enUS) : (sbyte?)null;
        }

        #endregion

        #region Int16

        public static implicit operator ExcelValue(short value)
        {
            return new ExcelValue(value.ToString(), false);
        }

        public static implicit operator ExcelValue(short? value)
        {
            return new ExcelValue(value?.ToString(), false);
        }

        public static implicit operator short(ExcelValue value)
        {
            return short.Parse(value._value, _enUS);
        }

        public static implicit operator short? (ExcelValue value)
        {
            return value._value != null ? short.Parse(value._value, _enUS) : (short?)null;
        }

        #endregion

        #region Int32

        public static implicit operator ExcelValue(int value)
        {
            return new ExcelValue(value.ToString(), false);
        }

        public static implicit operator ExcelValue(int? value)
        {
            return new ExcelValue(value?.ToString(), false);
        }

        public static implicit operator int(ExcelValue value)
        {
            return int.Parse(value._value, _enUS);
        }

        public static implicit operator int? (ExcelValue value)
        {
            return value._value != null ? int.Parse(value._value, _enUS) : (int?)null;
        }

        #endregion

        #region Int64

        public static implicit operator ExcelValue(long value)
        {
            return new ExcelValue(value.ToString(), false);
        }

        public static implicit operator ExcelValue(long? value)
        {
            return new ExcelValue(value?.ToString(), false);
        }

        public static implicit operator long(ExcelValue value)
        {
            return long.Parse(value._value, _enUS);
        }

        public static implicit operator long? (ExcelValue value)
        {
            return value._value != null ? long.Parse(value._value, _enUS) : (long?)null;
        }

        #endregion

        #region Byte

        public static implicit operator ExcelValue(byte value)
        {
            return new ExcelValue(value.ToString(), false);
        }

        public static implicit operator ExcelValue(byte? value)
        {
            return new ExcelValue(value?.ToString(), false);
        }

        public static implicit operator byte(ExcelValue value)
        {
            return byte.Parse(value._value, _enUS);
        }

        public static implicit operator byte? (ExcelValue value)
        {
            return value._value != null ? byte.Parse(value._value, _enUS) : (byte?)null;
        }

        #endregion

        #region UInt16

        public static implicit operator ExcelValue(ushort value)
        {
            return new ExcelValue(value.ToString(), false);
        }

        public static implicit operator ExcelValue(ushort? value)
        {
            return new ExcelValue(value?.ToString(), false);
        }

        public static implicit operator ushort(ExcelValue value)
        {
            return ushort.Parse(value._value, _enUS);
        }

        public static implicit operator ushort? (ExcelValue value)
        {
            return value._value != null ? ushort.Parse(value._value, _enUS) : (ushort?)null;
        }

        #endregion

        #region UInt32

        public static implicit operator ExcelValue(uint value)
        {
            return new ExcelValue(value.ToString(), false);
        }

        public static implicit operator ExcelValue(uint? value)
        {
            return new ExcelValue(value?.ToString(), false);
        }

        public static implicit operator uint(ExcelValue value)
        {
            return uint.Parse(value._value, _enUS);
        }

        public static implicit operator uint? (ExcelValue value)
        {
            return value._value != null ? uint.Parse(value._value, _enUS) : (uint?)null;
        }

        #endregion

        #region UInt64

        public static implicit operator ExcelValue(ulong value)
        {
            return new ExcelValue(value.ToString(), false);
        }

        public static implicit operator ExcelValue(ulong? value)
        {
            return new ExcelValue(value?.ToString(), false);
        }

        public static implicit operator ulong(ExcelValue value)
        {
            return ulong.Parse(value._value, _enUS);
        }

        public static implicit operator ulong? (ExcelValue value)
        {
            return value._value != null ? ulong.Parse(value._value, _enUS) : (ulong?)null;
        }

        #endregion

        #region Char

        public static implicit operator ExcelValue(char value)
        {
            return new ExcelValue(value.ToString(), false);
        }

        public static implicit operator ExcelValue(char? value)
        {
            return new ExcelValue(value?.ToString(), false);
        }

        public static implicit operator char(ExcelValue value)
        {
            return value._value[0];
        }

        public static implicit operator char? (ExcelValue value)
        {
            return value._value != null ? value._value.Length == 0 ? (char?)null : value._value[0] : (char?)null;
        }

        #endregion

        #region String

        public static implicit operator ExcelValue(string value)
        {
            return new ExcelValue(value.ToString(), false);
        }

        public static implicit operator string(ExcelValue value)
        {
            return value._value;
        }
        
        #endregion
    }
}
