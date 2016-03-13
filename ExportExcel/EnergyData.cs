using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace ExportExcel
{
    class EnergyData
    {
        class EngrgyDataRaw
        {
            public byte[] year;//年 2byte
            public byte[] mouth;//月 1byte
            public byte[] day;//日 1byte
            public byte[] power1;//正向电能 4byte
            public byte[] power2;//负向电能 4byte
            public byte[] powerAll;//总电能 4byte
        }
        List<EngrgyDataRaw> mEngrgyDataRawList = new List<EngrgyDataRaw>();

        public EnergyData(String filename)
        {
            ReadFromFile(filename);
        }

        bool ReadFromFile(String filename)
        {
            if (!File.Exists(filename))
            {
                Console.WriteLine("不存在此文件");
                return false;
            }
            else {
                try
                {
                    FileStream fs = new FileStream(filename, FileMode.Open);
                    BinaryReader br = new BinaryReader(fs);
                    try
                    {
                        while (true)
                        {
                            EngrgyDataRaw _EngrgyDataRaw = new EngrgyDataRaw();
                            // 2 1 1 4 4 4
                            _EngrgyDataRaw.year = br.ReadBytes(2);
                            _EngrgyDataRaw.mouth = br.ReadBytes(1);
                            _EngrgyDataRaw.day = br.ReadBytes(1);
                            _EngrgyDataRaw.power1 = br.ReadBytes(4);
                            _EngrgyDataRaw.power2 = br.ReadBytes(4);
                            _EngrgyDataRaw.powerAll = br.ReadBytes(4);

                            if (_EngrgyDataRaw.year == null
                                || _EngrgyDataRaw.mouth == null
                                || _EngrgyDataRaw.day == null
                                || _EngrgyDataRaw.power1 == null
                                || _EngrgyDataRaw.power2 == null
                                || _EngrgyDataRaw.powerAll == null)
                            {
                                _EngrgyDataRaw = null;
                                break;
                            }
                            //*
                            Console.WriteLine("year:" + BitConverter.ToUInt16(ToHostEndian(_EngrgyDataRaw.year),0)
                                + " mouth:" + BitConverter.ToString(_EngrgyDataRaw.mouth)
                                + " day:" + BitConverter.ToString(_EngrgyDataRaw.day)
                                + " power1:" + BitConverter.ToUInt32(ToHostEndian(_EngrgyDataRaw.power1), 0)
                                + " power2:" + BitConverter.ToUInt32(ToHostEndian(_EngrgyDataRaw.power2), 0)
                                + " powerAll:" + BitConverter.ToUInt32(ToHostEndian(_EngrgyDataRaw.powerAll), 0));//*/
                            
                            mEngrgyDataRawList.Add(_EngrgyDataRaw);
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
            for (int i = src.Length-1, j = 0; i >= 0; i--, j++) {
                dest[j] = src[i];
            }

            return dest;
        }
    }
}
