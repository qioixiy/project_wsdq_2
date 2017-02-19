using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExportExcel
{
    public class Myutility
    {
        public static byte[] ToHostEndian(byte[] src)
        {
            byte[] dest = new byte[src.Length];
            for (int i = src.Length - 1, j = 0; i >= 0; i--, j++)
            {
                dest[j] = src[i];
            }

            return dest;
        }
        public static bool InInt32Scope(Int32 value, Int32 min, Int32 max)
        {
            bool ret = true;
            if (value < min || value > max) {
                ret = false;
            }

            return ret;
        }
    }
}
