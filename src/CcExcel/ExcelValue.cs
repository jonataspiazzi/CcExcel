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

        #region Public

        public bool IsEmpty => string.IsNullOrEmpty(_value);

        #endregion

        #region Parses

        #region Boolean

        public static ExcelValue FromBoolean(bool value)
        {
            return new ExcelValue(value ? "1" : "0", false);
        }

        public static ExcelValue FromNullableBoolean(bool? value)
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

        public static implicit operator ExcelValue(bool value)
        {
            return FromBoolean(value);
        }

        public static implicit operator ExcelValue(bool? value)
        {
            return FromNullableBoolean(value);
        }

        #endregion

        #region Byte

        public static ExcelValue FromByte(byte value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static ExcelValue FromNullableByte(byte? value)
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

        public static implicit operator ExcelValue(byte value)
        {
            return FromByte(value);
        }

        public static implicit operator ExcelValue(byte? value)
        {
            return FromNullableByte(value);
        }

        #endregion

        #region Char

        public static ExcelValue FromChar(char value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static ExcelValue FromNullableChar(char? value)
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

        public static implicit operator ExcelValue(char value)
        {
            return FromChar(value);
        }

        public static implicit operator ExcelValue(char? value)
        {
            return FromNullableChar(value);
        }

        #endregion

        #region Double

        public static ExcelValue FromDouble(double value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static ExcelValue FromNullableDouble(double? value)
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

        public static implicit operator ExcelValue(double value)
        {
            return FromDouble(value);
        }

        public static implicit operator ExcelValue(double? value)
        {
            return FromNullableDouble(value);
        }

        #endregion

        #region Int16

        public static ExcelValue FromInt16(short value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static ExcelValue FromNullableInt16(short? value)
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

        public static implicit operator ExcelValue(short value)
        {
            return FromInt16(value);
        }

        public static implicit operator ExcelValue(short? value)
        {
            return FromNullableInt16(value);
        }

        #endregion

        #region Int32

        public static ExcelValue FromInt32(int value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static ExcelValue FromNullableInt32(int? value)
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

        public static implicit operator ExcelValue(int value)
        {
            return FromInt32(value);
        }

        public static implicit operator ExcelValue(int? value)
        {
            return FromNullableInt32(value);
        }

        #endregion

        #region Int64

        public static ExcelValue FromInt64(long value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static ExcelValue FromNullableInt64(long? value)
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

        public static implicit operator ExcelValue(long value)
        {
            return FromInt64(value);
        }

        public static implicit operator ExcelValue(long? value)
        {
            return FromNullableInt64(value);
        }

        #endregion

        #region SByte

        public static ExcelValue FromSByte(sbyte value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static ExcelValue FromNullableSByte(sbyte? value)
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

        public static implicit operator ExcelValue(sbyte value)
        {
            return FromSByte(value);
        }

        public static implicit operator ExcelValue(sbyte? value)
        {
            return FromNullableSByte(value);
        }

        #endregion

        #region Single

        public static ExcelValue FromSingle(float value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static ExcelValue FromNullableSingle(float? value)
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

        public static implicit operator ExcelValue(float value)
        {
            return FromSingle(value);
        }

        public static implicit operator ExcelValue(float? value)
        {
            return FromNullableSingle(value);
        }

        #endregion

        #region UInt16

        public static ExcelValue FromUInt16(ushort value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static ExcelValue FromNullableUInt16(ushort? value)
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

        public static implicit operator ExcelValue(ushort value)
        {
            return FromUInt16(value);
        }

        public static implicit operator ExcelValue(ushort? value)
        {
            return FromNullableUInt16(value);
        }

        #endregion

        #region UInt32

        public static ExcelValue FromUInt32(uint value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static ExcelValue FromNullableUInt32(uint? value)
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

        public static implicit operator ExcelValue(uint value)
        {
            return FromUInt32(value);
        }

        public static implicit operator ExcelValue(uint? value)
        {
            return FromNullableUInt32(value);
        }

        #endregion

        #region UInt64

        public static ExcelValue FromUInt64(ulong value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static ExcelValue FromNullableUInt64(ulong? value)
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

        public static implicit operator ExcelValue(ulong value)
        {
            return FromUInt64(value);
        }

        public static implicit operator ExcelValue(ulong? value)
        {
            return FromNullableUInt64(value);
        }

        #endregion

        #region Decimal

        public static ExcelValue FromDecimal(decimal value)
        {
            return new ExcelValue(value.ToString(_enUS), false);
        }

        public static ExcelValue FromNullableDecimal(decimal? value)
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

        public static implicit operator ExcelValue(decimal value)
        {
            return FromDecimal(value);
        }

        public static implicit operator ExcelValue(decimal? value)
        {
            return FromNullableDecimal(value);
        }

        #endregion

        #region DateTime

        public static ExcelValue FromDateTime(DateTime value)
        {
            return new ExcelValue(value.ToOADate().ToString(_enUS), false);
        }

        public static ExcelValue FromNullableDateTime(DateTime? value)
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

        public static implicit operator ExcelValue(DateTime value)
        {
            return FromDateTime(value);
        }

        public static implicit operator ExcelValue(DateTime? value)
        {
            return FromNullableDateTime(value);
        }

        #endregion

        #region TimeSpan

        public static ExcelValue FromTimeSpan(TimeSpan value)
        {
            return new ExcelValue((DateTime.FromOADate(0) + value).ToOADate().ToString(_enUS), false);
        }

        public static ExcelValue FromNullableTimeSpan(TimeSpan? value)
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

        public static implicit operator ExcelValue(TimeSpan value)
        {
            return FromTimeSpan(value);
        }

        public static implicit operator ExcelValue(TimeSpan? value)
        {
            return FromNullableTimeSpan(value);
        }

        #endregion

        #region String

        public static ExcelValue FromString(string value)
        {
            return new ExcelValue(value.ToString(_enUS), true);
        }

        public override string ToString()
        {
            return _value;
        }

        public static implicit operator ExcelValue(string value)
        {
            return FromString(value);
        }

        #endregion

        #endregion
    }
}
