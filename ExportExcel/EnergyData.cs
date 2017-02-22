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
            public byte[] year;//年 2byte
            public byte[] mouth;//月 1byte
            public byte[] day;//日 1byte
            public byte[] power1;//正向电能 4byte
            public byte[] power2;//负向电能 4byte
            public byte[] powerAll;//总电能 4byte
        }
        
        public byte[] reserve = new byte[1];
        public byte[] carType = new byte[1];
        public byte[] carNum = new byte[2];
        public List<EnergyDataRaw> mEnergyDataRawList = new List<EnergyDataRaw>();

        public EnergyData(String filename)
        {
            ReadFromFile(filename);
        }

        static int count = 0;
        bool ValidData(EnergyDataRaw _EnergyDataRaw)
        {
            if (count++ > 300) {
                // for test
                //return false;
            } 
            bool ret = true;

            if (Myutility.GetMajorVersionNumber() == "V1.3")
            {
                string v0_0, v0_1, v0_2, v0_3;

                v0_0 = _EnergyDataRaw.year[0].ToString();
                v0_1 = _EnergyDataRaw.year[1].ToString();
                v0_2 = Int32.Parse(BitConverter.ToString(_EnergyDataRaw.mouth), System.Globalization.NumberStyles.HexNumber).ToString();
                v0_3 = Int32.Parse(BitConverter.ToString(_EnergyDataRaw.day), System.Globalization.NumberStyles.HexNumber).ToString();
                // 验证有效性： 大于31或者小于等于0；时：大于等于24；分大于等于60；秒大于等于60；就丢弃这16个字节。
                Int32 day = Int32.Parse(v0_0);
                Int32 hour = Int32.Parse(v0_1);
                Int32 minuts = Int32.Parse(v0_2);
                Int32 second = Int32.Parse(v0_3);

                if (!(Myutility.InInt32Scope(day, 1, 31)
                    && Myutility.InInt32Scope(hour, 0, 23)
                    && Myutility.InInt32Scope(minuts, 0, 59)
                    && Myutility.InInt32Scope(second, 0, 59)))
                {
                    Console.WriteLine("无效数据" + v0_0 + "日" + v0_1 + "时" + v0_2 + "分" + v0_3 + "秒");
                    ret = false;
                }
            }

            return ret;
        }
        bool ReadFromFile(String filename)
        {
            if (!File.Exists(filename))
            {
                MessageBox.Show("不存在此文件");
                return false;
            }
            else {
                try
                {
                    FileStream fs = new FileStream(filename, FileMode.Open);
                    BinaryReader br = new BinaryReader(fs);
                    try
                    {
                        if ((Myutility.GetMajorVersionNumber() != "V1.1")
                            && (Myutility.GetMajorVersionNumber() != "V1.3"))
                        {
                            reserve = br.ReadBytes(1);
                            carType = br.ReadBytes(1);
                            carNum = br.ReadBytes(2);
                            byte temp = carNum[0];
                            carNum[0] = carNum[1];
                            carNum[1] = temp;
                        }
                        while (true)
                        {
                            EnergyDataRaw _EnergyDataRaw = new EnergyDataRaw();
                            // 2 1 1 4 4 4
                            _EnergyDataRaw.year = br.ReadBytes(2);
                            _EnergyDataRaw.mouth = br.ReadBytes(1);
                            _EnergyDataRaw.day = br.ReadBytes(1);
                            _EnergyDataRaw.power1 = br.ReadBytes(4);
                            _EnergyDataRaw.power2 = br.ReadBytes(4);
                            _EnergyDataRaw.powerAll = br.ReadBytes(4);

                            if (_EnergyDataRaw.year == null
                                || _EnergyDataRaw.mouth == null
                                || _EnergyDataRaw.day == null
                                || _EnergyDataRaw.power1 == null
                                || _EnergyDataRaw.power2 == null
                                || _EnergyDataRaw.powerAll == null
                                || _EnergyDataRaw.year.Length == 0
                                || _EnergyDataRaw.mouth.Length == 0
                                || _EnergyDataRaw.day.Length == 0
                                || _EnergyDataRaw.power1.Length == 0
                                || _EnergyDataRaw.power2.Length == 0
                                || _EnergyDataRaw.powerAll.Length == 0)
                            {
                                _EnergyDataRaw = null;
                                break;
                            }
                            /*
                            Console.WriteLine("year:" + BitConverter.ToInt16(ToHostEndian(_EnergyDataRaw.year),0)
                                + " mouth:" + BitConverter.ToString(_EnergyDataRaw.mouth)
                                + " day:" + BitConverter.ToString(_EnergyDataRaw.day)
                                + " power1:" + BitConverter.ToInt32(ToHostEndian(_EnergyDataRaw.power1), 0)
                                + " power2:" + BitConverter.ToInt32(ToHostEndian(_EnergyDataRaw.power2), 0)
                                + " powerAll:" + BitConverter.ToInt32(ToHostEndian(_EnergyDataRaw.powerAll), 0));//*/
                           
                            if (ValidData(_EnergyDataRaw)) {
                                mEnergyDataRawList.Add(_EnergyDataRaw);
                            }
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
       
    }
}
