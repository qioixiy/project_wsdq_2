using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace ExportExcel
{
    public class EnergyData
    {
        public class EnergyDataRaw
        {
            // 0-15
            public byte[] day;//日 1byte
            public byte[] hour;//时 1byte
            public byte[] minuts;//分 1byte
            public byte[] second;//秒 1byte
            public byte[] power1;//正向电能 4byte
            public byte[] power2;//负向电能 4byte
            public byte[] powerAll;//总电能 4byte

            // 16-40
            public byte[] v; // 电压2个字节，2位小数（读取的数据除以100）单位KV
            public byte[] i; // 电流2个字节，1位小数（读取的数据除以10）单位A
            public byte[] powerFactor; // 功率因素 2个字节 3位小数（读取的数据除以1000）无单位
            public byte[] powerRealTime; //实时功率 2个字节 单位KW
            public byte[] v3rd; // 电压3次谐波含油率 2个字节 2位小数（读取的数据除以100）单位 %
            public byte[] v5rd; // 电压5次谐波含油率 2个字节 2位小数（读取的数据除以100）单位 %
            public byte[] v7rd; // 电压7次谐波含油率 2个字节 2位小数（读取的数据除以100）单位 %
            public byte[] v9rd; // 电压9次谐波含油率 2个字节 2位小数（读取的数据除以100）单位 %
            public byte[] i3rd; // 电流3次谐波含油率 2个字节 2位小数（读取的数据除以100）单位 %
            public byte[] i5rd; // 电流5次谐波含油率 2个字节 2位小数（读取的数据除以100）单位 %
            public byte[] i7rd; // 电流7次谐波含油率 2个字节 2位小数（读取的数据除以100）单位 %
            public byte[] i9rd; // 电流9次谐波含油率 2个字节 2位小数（读取的数据除以100）单位 %
            public byte[] Rosebowcar; // 升弓车厢 1个字节 =1表示3车升弓；=2表示6车升弓；=0表示未升弓；

            // 41-47
            public byte[] revese; // 预留
            // 多个字节先转换成大端模式

            public int getDay()
            {
                return Int32.Parse(BitConverter.ToString(day), System.Globalization.NumberStyles.HexNumber);
            }
            public int getHour()
            {
                return Int32.Parse(BitConverter.ToString(hour), System.Globalization.NumberStyles.HexNumber);
            }
            public int getMinuts()
            {
                return Int32.Parse(BitConverter.ToString(minuts), System.Globalization.NumberStyles.HexNumber);
            }
            public int getSecond()
            {
                return Int32.Parse(BitConverter.ToString(second), System.Globalization.NumberStyles.HexNumber);
            }
            public int getPower1()
            {
                return BitConverter.ToInt32(Myutility.ToHostEndian(power1), 0);
            }
            public int getPower2()
            {
                return BitConverter.ToInt32(Myutility.ToHostEndian(power2), 0);
            }
            public int getPowerAll()
            {
                return BitConverter.ToInt32(Myutility.ToHostEndian(powerAll), 0);
            }

            public float getV()
            {
                return BitConverter.ToInt16(Myutility.ToHostEndian(v), 0) / 100.0f;
            }
            public float getI()
            {
                return BitConverter.ToInt16(Myutility.ToHostEndian(i), 0) / 10.0f;
            }
            public float getPowerFactor()
            {
                return BitConverter.ToInt16(Myutility.ToHostEndian(powerFactor), 0) / 1000.0f;
            }
            public float getPowerRealTime()
            {
                return BitConverter.ToInt16(Myutility.ToHostEndian(powerRealTime), 0) / 1.0f;
            }
            public float getV3rd()
            {
                return BitConverter.ToInt16(Myutility.ToHostEndian(v3rd), 0) / 100.0f;
            }
            public float getV5rd()
            {
                return BitConverter.ToInt16(Myutility.ToHostEndian(v5rd), 0) / 100.0f;
            }
            public float getV7rd()
            {
                return BitConverter.ToInt16(Myutility.ToHostEndian(v7rd), 0) / 100.0f;
            }
            public float getV9rd()
            {
                return BitConverter.ToInt16(Myutility.ToHostEndian(v9rd), 0) / 100.0f;
            }
            public float getI3rd()
            {
                return BitConverter.ToInt16(Myutility.ToHostEndian(i3rd), 0) / 100.0f;
            }
            public float getI5rd()
            {
                return BitConverter.ToInt16(Myutility.ToHostEndian(i5rd), 0) / 100.0f;
            }
            public float getI7rd()
            {
                return BitConverter.ToInt16(Myutility.ToHostEndian(i7rd), 0) / 100.0f;
            }
            public float getI9rd()
            {
                return BitConverter.ToInt16(Myutility.ToHostEndian(i9rd), 0) / 100.0f;
            }
            public int getRosebowcar()
            {
                return Int32.Parse(BitConverter.ToString(Rosebowcar), System.Globalization.NumberStyles.HexNumber);
                return BitConverter.ToChar(Myutility.ToHostEndian(Rosebowcar), 0);
            }
        }

        public byte[] reserve = new byte[1];
        public byte[] carType = new byte[1];
        public byte[] carNum = new byte[2];
        public List<EnergyDataRaw> mEnergyDataRawList = new List<EnergyDataRaw>();

        public EnergyData(String filename)
        {
            ReadFromFile(filename);
        }

        bool ReadFromFile(String filename)
        {
            if (!File.Exists(filename))
            {
                MessageBox.Show("不存在此文件");
                return false;
            }
            else
            {
                try
                {
                    FileStream fs = new FileStream(filename, FileMode.Open);
                    BinaryReader br = new BinaryReader(fs);
                    try
                    {
                        while (true)
                        {
                            EnergyDataRaw _EnergyDataRaw = new EnergyDataRaw();

                            byte[] buffer = br.ReadBytes(48);

                            _EnergyDataRaw.day = SubByte(buffer, 0, 1);//日 1byte
                            _EnergyDataRaw.hour = SubByte(buffer, 1, 1);//时 1byte
                            _EnergyDataRaw.minuts = SubByte(buffer, 2, 1);//分 1byte
                            _EnergyDataRaw.second = SubByte(buffer, 3, 1);//秒 1byte
                            _EnergyDataRaw.power1 = SubByte(buffer, 4, 4);//正向电能 4byte
                            _EnergyDataRaw.power2 = SubByte(buffer, 8, 4);//负向电能 4byte
                            _EnergyDataRaw.powerAll = SubByte(buffer, 12, 4);//总电能 4byte

                            // 16-40
                            _EnergyDataRaw.v = SubByte(buffer, 16, 2); // 电压2个字节，2位小数（读取的数据除以100）单位KV
                            _EnergyDataRaw.i = SubByte(buffer, 18, 2); // 电流2个字节，1位小数（读取的数据除以10）单位A
                            _EnergyDataRaw.powerFactor = SubByte(buffer, 20, 2); // 功率因素 2个字节 3位小数（读取的数据除以1000）无单位
                            _EnergyDataRaw.powerRealTime = SubByte(buffer, 22, 2); //实时功率 2个字节 单位KW
                            _EnergyDataRaw.v3rd = SubByte(buffer, 24, 2); // 电压3次谐波含油率 2个字节 2位小数（读取的数据除以100）单位 %
                            _EnergyDataRaw.v5rd = SubByte(buffer, 26, 2); // 电压5次谐波含油率 2个字节 2位小数（读取的数据除以100）单位 %
                            _EnergyDataRaw.v7rd = SubByte(buffer, 28, 2); // 电压7次谐波含油率 2个字节 2位小数（读取的数据除以100）单位 %
                            _EnergyDataRaw.v9rd = SubByte(buffer, 30, 2); // 电压9次谐波含油率 2个字节 2位小数（读取的数据除以100）单位 %
                            _EnergyDataRaw.i3rd = SubByte(buffer, 32, 2); // 电流3次谐波含油率 2个字节 2位小数（读取的数据除以100）单位 %
                            _EnergyDataRaw.i5rd = SubByte(buffer, 34, 2); // 电流5次谐波含油率 2个字节 2位小数（读取的数据除以100）单位 %
                            _EnergyDataRaw.i7rd = SubByte(buffer, 36, 2); // 电流7次谐波含油率 2个字节 2位小数（读取的数据除以100）单位 %
                            _EnergyDataRaw.i9rd = SubByte(buffer, 38, 2); // 电流9次谐波含油率 2个字节 2位小数（读取的数据除以100）单位 %
                            _EnergyDataRaw.Rosebowcar = SubByte(buffer, 40, 1); // 升弓车厢 1个字节 =1表示3车升弓；=2表示6车升弓；=0表示未升弓；

                            if (_EnergyDataRaw.day == null
                                || _EnergyDataRaw.hour == null
                                || _EnergyDataRaw.minuts == null
                                || _EnergyDataRaw.second == null
                                || _EnergyDataRaw.power1 == null
                                || _EnergyDataRaw.power2 == null
                                || _EnergyDataRaw.powerAll == null
                                || _EnergyDataRaw.v == null
                                || _EnergyDataRaw.i == null
                                || _EnergyDataRaw.powerFactor == null
                                || _EnergyDataRaw.powerRealTime == null
                                || _EnergyDataRaw.v3rd == null
                                || _EnergyDataRaw.v5rd == null
                                || _EnergyDataRaw.v7rd == null
                                || _EnergyDataRaw.v9rd == null
                                || _EnergyDataRaw.i3rd == null
                                || _EnergyDataRaw.i5rd == null
                                || _EnergyDataRaw.i7rd == null
                                || _EnergyDataRaw.i9rd == null
                                || _EnergyDataRaw.day.Length != 1
                                || _EnergyDataRaw.hour.Length != 1
                                || _EnergyDataRaw.minuts.Length != 1
                                || _EnergyDataRaw.second.Length != 1
                                || _EnergyDataRaw.power1.Length != 4
                                || _EnergyDataRaw.power2.Length != 4
                                || _EnergyDataRaw.powerAll.Length != 4
                                || _EnergyDataRaw.v.Length != 2
                                || _EnergyDataRaw.i.Length != 2
                                || _EnergyDataRaw.powerFactor.Length != 2
                                || _EnergyDataRaw.powerRealTime.Length != 2
                                || _EnergyDataRaw.v3rd.Length != 2
                                || _EnergyDataRaw.v5rd.Length != 2
                                || _EnergyDataRaw.v7rd.Length != 2
                                || _EnergyDataRaw.v9rd.Length != 2
                                || _EnergyDataRaw.i3rd.Length != 2
                                || _EnergyDataRaw.i5rd.Length != 2
                                || _EnergyDataRaw.i7rd.Length != 2
                                || _EnergyDataRaw.i9rd.Length != 2)
                            {
                                _EnergyDataRaw = null;
                                break;
                            }

                            mEnergyDataRawList.Add(_EnergyDataRaw);
                        }
                    }
                    catch (Exception)
                    {
                        Console.WriteLine("读取结束！");
                    }
                    br.Close();
                    fs.Close();
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }

                return true;
            }
        }

        byte[] ToHostEndian(byte[] src)
        {
            byte[] dest = new byte[src.Length];
            for (int i = src.Length - 1, j = 0; i >= 0; i--, j++)
            {
                dest[j] = src[i];
            }

            return dest;
        }

        /// <summary>  
        /// 截取字节数组  
        /// </summary>  
        /// <param name="srcBytes">要截取的字节数组</param>  
        /// <param name="startIndex">开始截取位置的索引</param>  
        /// <param name="length">要截取的字节长度</param>  
        /// <returns>截取后的字节数组</returns>  
        public byte[] SubByte(byte[] srcBytes, int startIndex, int length)
        {
            System.IO.MemoryStream bufferStream = new System.IO.MemoryStream();
            byte[] returnByte = new byte[] { };
            if (srcBytes == null) { return returnByte; }
            if (startIndex < 0) { startIndex = 0; }
            if (startIndex < srcBytes.Length)
            {
                if (length < 1 || length > srcBytes.Length - startIndex) { length = srcBytes.Length - startIndex; }
                bufferStream.Write(srcBytes, startIndex, length);
                returnByte = bufferStream.ToArray();
                bufferStream.SetLength(0);
                bufferStream.Position = 0;
            }
            bufferStream.Close();
            bufferStream.Dispose();
            return returnByte;
        }
    }
}
