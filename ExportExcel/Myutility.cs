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

        public static string GetMajorVersionNumber() {
            string originVer = ExportExcel.Properties.Resources.Version;
            string ret = originVer;

            switch (originVer) {
                case "V1.11": ret = "V1.1"; break;
                case "V1.21": ret = "V1.2"; break;
                case "V1.31": ret = "V1.3"; break;
                default: ret = "V1.3"; break;
            }

            return ret;
        }
    }
}
