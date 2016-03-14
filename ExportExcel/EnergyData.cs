using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace ExportExcel
{
    class EnergyData
    {
        class EnergyDataRaw
        {
            public byte[] year;//年 2byte
            public byte[] mouth;//月 1byte
            public byte[] day;//日 1byte
            public byte[] power1;//正向电能 4byte
            public byte[] power2;//负向电能 4byte
            public byte[] powerAll;//总电能 4byte
        }
        
        public List<EnergyDataRaw> mEnergyDataRawList = new List<EnergyDataRaw>();

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
                                || _EnergyDataRaw.powerAll == null)
                            {
                                _EnergyDataRaw = null;
                                break;
                            }
                            //*
                            Console.WriteLine("year:" + BitConverter.ToUInt16(ToHostEndian(_EnergyDataRaw.year),0)
                                + " mouth:" + BitConverter.ToString(_EnergyDataRaw.mouth)
                                + " day:" + BitConverter.ToString(_EnergyDataRaw.day)
                                + " power1:" + BitConverter.ToUInt32(ToHostEndian(_EnergyDataRaw.power1), 0)
                                + " power2:" + BitConverter.ToUInt32(ToHostEndian(_EnergyDataRaw.power2), 0)
                                + " powerAll:" + BitConverter.ToUInt32(ToHostEndian(_EnergyDataRaw.powerAll), 0));//*/
                            
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
            for (int i = src.Length-1, j = 0; i >= 0; i--, j++) {
                dest[j] = src[i];
            }

            return dest;
        }
    }
}
