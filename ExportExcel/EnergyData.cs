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
        public List<EnergyDataRaw> mBackupEnergyDataRawList = new List<EnergyDataRaw>();

        public EnergyData(String filename)
        {
            mBackupEnergyDataRawList = new List<EnergyDataRaw>();
            ReadFromFile(filename);
        }

        int probeDataHeader(String filename)
        {
            int ret = 0;
            
            FileStream fs = new FileStream(filename, FileMode.Open);
            BinaryReader br = new BinaryReader(fs);

            while (true)
            {
                byte[] buffer = br.ReadBytes(16);

                EnergyDataRaw tEnergyDataRaw = new EnergyDataRaw();
                tEnergyDataRaw.day = SubByte(buffer, 0, 1);//日 1byte
                tEnergyDataRaw.hour = SubByte(buffer, 1, 1);//时 1byte
                tEnergyDataRaw.minuts = SubByte(buffer, 2, 1);//分 1byte
                tEnergyDataRaw.second = SubByte(buffer, 3, 1);//秒 1byte
                tEnergyDataRaw.power1 = SubByte(buffer, 4, 4);//正向电能 4byte
                tEnergyDataRaw.power2 = SubByte(buffer, 8, 4);//负向电能 4byte
                tEnergyDataRaw.powerAll = SubByte(buffer, 12, 4);//总电能 4byte

                //4.原先是我们存储解析是按照下图1这样的48个字节是一套存储信息，但是现实中由于FLASH存储存在按扇区擦除问题，
                //所以现实中有可能下载的数据00000000h被擦除了
                //这样V1.33就错开了，全解析错了，没能读到一条有用的信息，所以处理好开头很重要，
                //如何处理？一排16个字节，一条解析字段48个字节。所以打开一个文档先读16个字节，
                //第一个字节是日：日做判断条件，日大于等于1，小于等于31；
                //第二个字节是时：时做判断条件，时大于等于0，小于等于23；
                //第三个字节是分：分做判断条件，分大于等于0，小于等于59；
                //第四个字节是秒：秒做判断条件，秒大于等于0，小于等于59；
                //第五到八个字节是正向电能：
                //第九到十二个字节是反向电能：
                //第十三到十六个字节是总电能：
                //判断条件是:总电能=正向电能-反向电能 或者 总电能=反向电能-正向电能
                //如果这个几个条件满足，可以判断这个这16个字节是整个文档的开端，每48个字节为一条否则，丢弃读取下一16个字节判断开端。
                int day = tEnergyDataRaw.getDay();
                int hour = tEnergyDataRaw.getHour();
                int minuts = tEnergyDataRaw.getMinuts();
                int second = tEnergyDataRaw.getSecond();
                int powerall = tEnergyDataRaw.getPowerAll();
                int power1 = tEnergyDataRaw.getPower1();
                int power2 = tEnergyDataRaw.getPower2();
                if (day >= 1 && day <= 31
                    && hour >= 0 && hour <= 23
                    && minuts >= 0 && minuts <= 59
                    && second >= 0 && second <= 59
                    && powerall == (power1 - power2))
                {
                    break;
                }

                ret += 16;
            }

            br.Close();
            fs.Close();

            return ret;
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
                    int offset = probeDataHeader(filename);

                    FileStream fs = new FileStream(filename, FileMode.Open);
                    BinaryReader br = new BinaryReader(fs);
                    try
                    {
                        if (offset != 0)
                        {
                            byte[] discard = br.ReadBytes(offset);
                        }

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
