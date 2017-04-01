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
            public byte[] hour;//时 1byte
            public byte[] minute;//分 1byte
            public byte[] consumePower;//消耗电能 4byte
            public byte[] revivePower;//再生电能 4byte          
            public DateTime EnergyDate 
            { 
                get
                {
                    return this.GetDateTimeFilter();
                }  
            }
            public int consumeEnergy
            {
                get
                {
                    return this.getConsumePower();
                }
            }

            public int reviveEgergy
            {
                get
                {
                    return this.getRevivePower();
                }
            }

            // 总消耗能量=消耗电能-再生电能
            public int totalEnergy
            {
                get
                {
                    return consumeEnergy - reviveEgergy;
                }
            }

            public int getYear()
            {
                int ret = BitConverter.ToInt16(Myutility.ToHostEndian(year), 0);
                return ret;
            }
            public int getMouth()
            {
                int ret = mouth[0];
                return ret;
            }
            public int getDay()
            {
                int ret = day[0];
                return ret;
            }
            public int getHour()
            {
                int ret = hour[0];
                return ret;
            }
            public int getMinute()
            {
                int ret = minute[0];
                return ret;
            }
            public string GetDateTime()        //时间用于显示
            {
                string year = this.getYear().ToString();
                string mouth = this.getMouth().ToString();
                string day = this.getDay().ToString();
                string hour = this.getHour().ToString();
                string minute = this.getMinute().ToString();
                string ret = year + "年" + mouth + "月" + day + "日" + hour + "时" + minute + "分";
                return ret;
            }
            public DateTime GetDateTimeFilter()     //用于处理过滤的时间
            {
                string year = this.getYear().ToString();
                string mouth = this.getMouth().ToString();
                string day = this.getDay().ToString();
                string hour = this.getHour().ToString();
                string minute = this.getMinute().ToString();
                string DateTimeFilter = year + "-" + mouth + "-" + day + " " + hour + ":" + minute + ":00";
                DateTime ret = Convert.ToDateTime(DateTimeFilter);
                return ret;
            }
            public int getConsumePower()
            {
                int ret = BitConverter.ToInt32(Myutility.ToHostEndian(consumePower), 0);
                return ret;
            }
            public int getRevivePower()
            {

                int ret = BitConverter.ToInt32(Myutility.ToHostEndian(revivePower), 0);
                return ret;
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

        bool ValidData(EnergyDataRaw _EnergyDataRaw)
        {
            bool ret = true;

            // 验证有效性： 大于31或者小于等于0；时：大于等于24；分大于等于60；秒大于等于60；就丢弃这16个字节。
            Int32 year = _EnergyDataRaw.getYear();
            Int32 mounth = _EnergyDataRaw.getMouth();
            Int32 day = _EnergyDataRaw.getDay();
            Int32 hour = _EnergyDataRaw.getHour();
            Int32 minute = _EnergyDataRaw.getMinute();

            if (!(Myutility.InInt32Scope(year, 2000, 2900)
                && Myutility.InInt32Scope(mounth, 1, 12)
                && Myutility.InInt32Scope(day, 0, 31)
                && Myutility.InInt32Scope(hour, 0, 23)
                && Myutility.InInt32Scope(minute, 0, 59)))
            {
                Console.WriteLine("无效数据： " + year + "年" + mounth + "月" + day + "日" + hour + "时" + minute + "分");
                ret = false;
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
                        while (true)
                        {
                            byte[] buffer = br.ReadBytes(16);
                            EnergyDataRaw _EnergyDataRaw = new EnergyDataRaw();
                            // 2 1 1 1 1 4 4
                            _EnergyDataRaw.year = Myutility.SubByte(buffer, 0, 2); //年 2byte
                            _EnergyDataRaw.mouth = Myutility.SubByte(buffer, 2, 1); //月 1byte
                            _EnergyDataRaw.day = Myutility.SubByte(buffer, 3, 1); //日 1byte
                            _EnergyDataRaw.hour = Myutility.SubByte(buffer, 4, 1); //时 1byte
                            _EnergyDataRaw.minute = Myutility.SubByte(buffer, 5, 1); //分 1byte
                            _EnergyDataRaw.consumePower = Myutility.SubByte(buffer, 6, 4); //消耗电能 4byte
                            _EnergyDataRaw.revivePower = Myutility.SubByte(buffer, 10, 4); //再生电能 4byte
                           
                            if (_EnergyDataRaw.year == null
                                || _EnergyDataRaw.mouth == null
                                || _EnergyDataRaw.day == null
                                || _EnergyDataRaw.hour == null
                                || _EnergyDataRaw.minute == null
                                || _EnergyDataRaw.consumePower == null
                                || _EnergyDataRaw.revivePower == null

                                || _EnergyDataRaw.year.Length != 2
                                || _EnergyDataRaw.mouth.Length != 1
                                || _EnergyDataRaw.day.Length != 1
                                || _EnergyDataRaw.hour.Length != 1
                                || _EnergyDataRaw.minute.Length != 1
                                || _EnergyDataRaw.consumePower.Length != 4
                                || _EnergyDataRaw.revivePower.Length != 4)
                            {
                                _EnergyDataRaw = null;
                                break;
                            }
                           
                            if (ValidData(_EnergyDataRaw)) {
                                mEnergyDataRawList.Add(_EnergyDataRaw);
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("读取异常结束！" + e.ToString());
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
    }
}
