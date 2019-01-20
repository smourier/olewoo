using System;
using System.Runtime.InteropServices;

namespace olewoo.interop
{
    [StructLayout(LayoutKind.Explicit)]
    public struct VARIANT
    {
        [FieldOffset(0)] public TypeUnion _typeUnion;
        [FieldOffset(0)] public decimal _decimal;

        [StructLayout(LayoutKind.Sequential)]
        public struct TypeUnion
        {
            public ushort _vt;
            public ushort _wReserved1;
            public ushort _wReserved2;
            public ushort _wReserved3;
            public UnionTypes _unionTypes;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct Record
        {
            private IntPtr _record;
            private IntPtr _recordInfo;
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct UnionTypes
        {
            [FieldOffset(0)] internal sbyte _i1;
            [FieldOffset(0)] internal short _i2;
            [FieldOffset(0)] internal int _i4;
            [FieldOffset(0)] internal long _i8;
            [FieldOffset(0)] internal byte _ui1;
            [FieldOffset(0)] internal ushort _ui2;
            [FieldOffset(0)] internal uint _ui4;
            [FieldOffset(0)] internal ulong _ui8;
            [FieldOffset(0)] internal int _int;
            [FieldOffset(0)] internal uint _uint;
            [FieldOffset(0)] internal short _bool;
            [FieldOffset(0)] internal int _error;
            [FieldOffset(0)] internal float _r4;
            [FieldOffset(0)] internal double _r8;
            [FieldOffset(0)] internal long _cy;
            [FieldOffset(0)] internal double _date;
            [FieldOffset(0)] internal IntPtr _bstr;
            [FieldOffset(0)] internal IntPtr _unknown;
            [FieldOffset(0)] internal IntPtr _dispatch;
            [FieldOffset(0)] internal IntPtr _pvarVal;
            [FieldOffset(0)] internal IntPtr _byref;
            [FieldOffset(0)] internal Record _record;
        }
    }
}
