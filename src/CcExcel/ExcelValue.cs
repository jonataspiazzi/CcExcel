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
            return new ExcelValue(value ? "1" : "0", false);
        }

        public static implicit operator ExcelValue(bool? value)
        {
            return new ExcelValue(value == null ? null : value.Value ? "1" : "0", false);
        }

        public bool ToBoolean()
        {
            return _value == "1" ? true : _value == "0" ? false : bool.Parse(_value);
        }

        public bool? ToNullableBoolean()
        {
            return string.IsNullOrWhiteSpace(_value) ? (bool?)null : ToBoolean();
        }

        #endregion

        #region Byte

        public static implicit operator ExcelValue(byte value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static implicit operator ExcelValue(byte? value)
        {
            return new ExcelValue(value?.ToString(_enUS), false);
        }

        public byte ToByte()
        {
            return byte.Parse(_value, _enUS);
        }

        public byte? ToNullableByte()
        {
            return string.IsNullOrWhiteSpace(_value) ? (byte?)null : ToByte();
        }

        #endregion

        #region Char

        public static implicit operator ExcelValue(char value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static implicit operator ExcelValue(char? value)
        {
            return new ExcelValue(value?.ToString(_enUS), false);
        }

        public char ToChar()
        {
            return char.Parse(_value);
        }

        public char? ToNullableChar()
        {
            return string.IsNullOrWhiteSpace(_value) ? (char?)null : ToChar();
        }

        #endregion

        #region Double

        public static implicit operator ExcelValue(double value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static implicit operator ExcelValue(double? value)
        {
            return new ExcelValue(value?.ToString(_enUS), false);
        }

        public double ToDouble()
        {
            return double.Parse(_value, _enUS);
        }

        public double? ToNullableDouble()
        {
            return string.IsNullOrWhiteSpace(_value) ? (double?)null : ToDouble();
        }

        #endregion

        #region Int16

        public static implicit operator ExcelValue(short value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static implicit operator ExcelValue(short? value)
        {
            return new ExcelValue(value?.ToString(_enUS), false);
        }

        public short ToInt16()
        {
            return short.Parse(_value, _enUS);
        }

        public short? ToNullableInt16()
        {
            return string.IsNullOrWhiteSpace(_value) ? (short?)null : ToInt16();
        }

        #endregion

        #region Int32

        public static implicit operator ExcelValue(int value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static implicit operator ExcelValue(int? value)
        {
            return new ExcelValue(value?.ToString(_enUS), false);
        }

        public int ToInt32()
        {
            return int.Parse(_value, _enUS);
        }

        public int? ToNullableInt32()
        {
            return string.IsNullOrWhiteSpace(_value) ? (int?)null : ToInt32();
        }

        #endregion

        #region Int64

        public static implicit operator ExcelValue(long value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static implicit operator ExcelValue(long? value)
        {
            return new ExcelValue(value?.ToString(_enUS), false);
        }

        public long ToInt64()
        {
            return long.Parse(_value, _enUS);
        }

        public long? ToNullableInt64()
        {
            return string.IsNullOrWhiteSpace(_value) ? (long?)null : ToInt64();
        }

        #endregion

        #region SByte

        public static implicit operator ExcelValue(sbyte value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static implicit operator ExcelValue(sbyte? value)
        {
            return new ExcelValue(value?.ToString(_enUS), false);
        }

        public sbyte ToSByte()
        {
            return sbyte.Parse(_value, _enUS);
        }

        public sbyte? ToNullableSByte()
        {
            return string.IsNullOrWhiteSpace(_value) ? (sbyte?)null : ToSByte();
        }

        #endregion

        #region Single

        public static implicit operator ExcelValue(float value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static implicit operator ExcelValue(float? value)
        {
            return new ExcelValue(value?.ToString(_enUS), false);
        }

        public float ToSingle()
        {
            return float.Parse(_value, _enUS);
        }

        public float? ToNullableSingle()
        {
            return string.IsNullOrWhiteSpace(_value) ? (float?)null : ToSingle();
        }

        #endregion

        #region UInt16

        public static implicit operator ExcelValue(ushort value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static implicit operator ExcelValue(ushort? value)
        {
            return new ExcelValue(value?.ToString(_enUS), false);
        }

        public ushort ToUInt16()
        {
            return ushort.Parse(_value, _enUS);
        }

        public ushort? ToNullableUInt16()
        {
            return string.IsNullOrWhiteSpace(_value) ? (ushort?)null : ToUInt16();
        }

        #endregion

        #region UInt32

        public static implicit operator ExcelValue(uint value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static implicit operator ExcelValue(uint? value)
        {
            return new ExcelValue(value?.ToString(_enUS), false);
        }

        public uint ToUInt32()
        {
            return uint.Parse(_value, _enUS);
        }

        public uint? ToNullableUInt32()
        {
            return string.IsNullOrWhiteSpace(_value) ? (uint?)null : ToUInt32();
        }

        #endregion

        #region UInt64

        public static implicit operator ExcelValue(ulong value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static implicit operator ExcelValue(ulong? value)
        {
            return new ExcelValue(value?.ToString(_enUS), false);
        }

        public ulong ToUInt64()
        {
            return ulong.Parse(_value, _enUS);
        }

        public ulong? ToNullableUInt64()
        {
            return string.IsNullOrWhiteSpace(_value) ? (ulong?)null : ToUInt64();
        }

        #endregion

        #region Decimal

        public static implicit operator ExcelValue(decimal value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static implicit operator ExcelValue(decimal? value)
        {
            return new ExcelValue(value?.ToString(_enUS), false);
        }

        public decimal ToDecimal()
        {
            return decimal.Parse(_value, _enUS);
        }

        public decimal? ToNullableDecimal()
        {
            return string.IsNullOrWhiteSpace(_value) ? (decimal?)null : ToDecimal();
        }

        #endregion

        #region DateTime

        public static implicit operator ExcelValue(DateTime value)
        {
            return new ExcelValue(value.ToOADate().ToString(_enUS), false);
        }

        public static implicit operator ExcelValue(DateTime? value)
        {
            return new ExcelValue(value?.ToOADate().ToString(_enUS), false);
        }

        public DateTime ToDateTime()
        {
            return DateTime.FromOADate(double.Parse(_value, _enUS));
        }

        public DateTime? ToNullableDateTime()
        {
            return string.IsNullOrWhiteSpace(_value) ? (DateTime?)null : ToDateTime();
        }

        #endregion

        #region TimeSpan

        public static implicit operator ExcelValue(TimeSpan value)
        {
            return new ExcelValue((DateTime.FromOADate(0) + value).ToOADate().ToString(_enUS), false);
        }

        public static implicit operator ExcelValue(TimeSpan? value)
        {
            return new ExcelValue((DateTime.FromOADate(0) + value)?.ToOADate().ToString(_enUS), false);
        }

        public TimeSpan ToTimeSpan()
        {
            var date = DateTime.FromOADate(double.Parse(_value, _enUS));

            return date - date.Date;
        }

        public TimeSpan? ToNullableTimeSpan()
        {
            return string.IsNullOrWhiteSpace(_value) ? (TimeSpan?)null : ToTimeSpan();
        }

        #endregion

        #region String

        public static implicit operator ExcelValue(string value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public override string ToString()
        {
            return _value;
        }

        #endregion
    }
}
