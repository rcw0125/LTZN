using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
using System.Data.OracleClient;

namespace LTZN.XiaohaoGaolu
{
    internal class Xiaohao : IEditableObject, INotifyPropertyChanged
    {
        /// <summary>
        /// 当前值
        /// </summary>
        private XiaohaoData currentData;
        private XiaohaoData oldData;
        private DateTime riqi;
        private int gaolu;
        /// <summary>
        /// 更改标识
        /// </summary>
        /// <remarks>标志对象的属性值是否有更改</remarks>
        private bool change = false;
        public Xiaohao()
        {
            this.riqi = DateTime.Today;
            this.gaolu = 1;
            Init();
        }
        public Xiaohao(DateTime riqi,int gaolu)
        {
            this.riqi = riqi;
            this.gaolu = gaolu;
            Init();
           
        }
        /// <summary>
        /// 日期
        /// </summary>
        public DateTime Riqi
        {
            get
            {
                return this.riqi;
            }
        }

        /// <summary>
        /// 高炉
        /// </summary>
        public int Gaolu
        {
            get
            {
                return this.gaolu;
            }

        }

        /// <summary>
        /// 富氧
        /// </summary>
        public double? Fuyang
        {
            get
            {
                return this.currentData.fuyang ;
            }
            set
            {
                this.currentData.fuyang  = value;
                NotifyPropertyChanged("Fuyang");
            }
        }

        /// <summary>
        /// 焦丁
        /// </summary>
        public double? Jiaoding
        {
            get
            {
                return this.currentData.jiaoding ;
            }
            set
            {
                this.currentData.jiaoding  = value;
                NotifyPropertyChanged("Jiaoding");
            }
        }

        private double? _JiaodingShuifen = null;
        /// <summary>
        /// 焦丁水份
        /// </summary>
        public double? JiaodingShuifen
        {
            get
            {
                return _JiaodingShuifen;
            }
            set
            {
                this._JiaodingShuifen = value;
                NotifyPropertyChanged("JiaodingShuifen");
            }
        }

        /// <summary>
        /// 冷风流量
        /// </summary>
        public double? Lengfengliuliang
        {
            get
            {
                return this.currentData.lengfengliuliang ;
            }
            set
            {
                this.currentData.lengfengliuliang = value;
                NotifyPropertyChanged("Lengfengliuliang");
            }
        }

        /// <summary>
        /// 风温
        /// </summary>
        public double? Fengwen
        {
            get
            {
                return this.currentData.fengwen ;
            }
            set
            {
                this.currentData.fengwen  = value;
                NotifyPropertyChanged("Fengwen");
            }
        }

        /// <summary>
        /// 热风压力
        /// </summary>
        public double? Refengyali
        {
            get
            {
                return this.currentData.refengyali ;
            }
            set
            {
                this.currentData.refengyali  = value;
                NotifyPropertyChanged("Refengyali");
            }
        }

        /// <summary>
        /// 炉顶温度
        /// </summary>
        public double? Ludingwendu
        {
            get
            {
                return this.currentData.ludingwendu ;
            }
            set
            {
                this.currentData.ludingwendu  = value;
                NotifyPropertyChanged("Ludingwendu");
            }
        }

        /// <summary>
        /// 炉顶压力
        /// </summary>
        public double? Ludingyali
        {
            get
            {
                return this.currentData.ludingyali ;
            }
            set
            {
                this.currentData.ludingyali  = value;
                NotifyPropertyChanged("Ludingyali");
            }
        }

        /// <summary>
        /// 湿焦粉
        /// </summary>
        public double? Shijiaofen
        {
            get
            {
                return this.currentData.shijiaofen ;
            }
            set
            {
                this.currentData.shijiaofen  = value;
                NotifyPropertyChanged("Shijiaofen");
            }
        }

        private double? _JiaofenShuifen = null;
        /// <summary>
        /// 焦粉水份
        /// </summary>
        public double? JiaofenShuifen
        {
            get
            {
                return _JiaofenShuifen;
            }
            set
            {
                this._JiaofenShuifen = value;
                NotifyPropertyChanged("JiaofenShuifen");
            }
        }

        /// <summary>
        /// 分数
        /// </summary>
        public double? Fenshu
        {
            get
            {
                return this.currentData.fenshu ;
            }
            set
            {
                this.currentData.fenshu = value;
                NotifyPropertyChanged("Fenshu");
            }
        }

        /// <summary>
        /// 炉数
        /// </summary>
        public int? Lushu
        {
            get
            {
                return this.currentData.lushu ;
            }
            set
            {
                this.currentData.lushu  = value;
                NotifyPropertyChanged("Lushu");
            }
        }

        /// <summary>
        /// 二氧化碳
        /// </summary>
        public double? CO2
        {
            get
            {
                return this.currentData.co2 ;
            }
            set
            {
                this.currentData.co2  = value;
                NotifyPropertyChanged("CO2");
            }
        }

        /// <summary>
        /// 一氧化碳
        /// </summary>
        public double? CO
        {
            get
            {
                return this.currentData.co;
            }
            set
            {
                this.currentData.co  = value;
                NotifyPropertyChanged("CO");
            }
        }

        /// <summary>
        /// 高压操作时间
        /// </summary>
        public int? Gaoyashijian
        {
            get
            {
                return this.currentData.gaoyashijian;
            }
            set
            {
                this.currentData.gaoyashijian = value;
                NotifyPropertyChanged("Gaoyashijian");
            }
        }

        /// <summary>
        /// 废铁
        /// </summary>
        public double? Feitie
        {
            get
            {
                return this.currentData.feitie ;
            }
            set
            {
                this.currentData.feitie  = value;
                NotifyPropertyChanged("Feitie");
            }
        }

        /// <summary>
        /// 外购废铁
        /// </summary>
        public double? Waigoufeitie
        {
            get
            {
                return this.currentData.waigoufeitie ;
            }
            set
            {
                this.currentData.waigoufeitie  = value;
                NotifyPropertyChanged("Waigoufeitie");
            }
        }

        /// <summary>
        /// 料批
        /// </summary>
        public int? Liaopi
        {
            get
            {
                return this.currentData.liaopi ;
            }
            set
            {
                this.currentData.liaopi  = value;
                NotifyPropertyChanged("Liaopi");
            }
        }

        /// <summary>
        /// 深料批数
        /// </summary>
        public int? Shenliaopi
        {
            get
            {
                return this.currentData.shenliaopi ;
            }
            set
            {
                this.currentData.shenliaopi  = value;
                NotifyPropertyChanged("Shenliaopi");
            }
        }

        /// <summary>
        /// 悬料
        /// </summary>
        public int? Xuanliao
        {
            get
            {
                return this.currentData.xuanliao ;
            }
            set
            {
                this.currentData.xuanliao = value;
                NotifyPropertyChanged("Xuanliao");
            }
        }

        /// <summary>
        /// 坐料
        /// </summary>
        public int? Zuoliao
        {
            get
            {
                return this.currentData.zuoliao ;
            }
            set
            {
                this.currentData.zuoliao  = value;
                NotifyPropertyChanged("Zuoliao");
            }
        }

        /// <summary>
        /// 崩料
        /// </summary>
        /// <value></value>
        public int? Bengliao
        {
            get
            {
                return this.currentData.bengliao;
            }
            set
            {
                this.currentData.bengliao  = value;
                NotifyPropertyChanged("Bengliao");
            }
        }

        /// <summary>
        /// 风口大套
        /// </summary>
        public int? Fengkoudatao
        {
            get
            {
                return this.currentData.fengkoudatao ;
            }
            set
            {
                this.currentData.fengkoudatao  = value;
                NotifyPropertyChanged("Fengkoudatao");
            }
        }

        /// <summary>
        /// 风口中套
        /// </summary>
        public int? Fengkouzhongtao
        {
            get
            {
                return this.currentData.fengkouzhongtao ;
            }
            set
            {
                this.currentData.fengkouzhongtao  = value;
                NotifyPropertyChanged("Fengkouzhongtao");
            }
        }

        /// <summary>
        /// 风口小套
        /// </summary>
        public int? Fengkouxiaotao
        {
            get
            {
                return this.currentData.fengkouxiaotao ;
            }
            set
            {
                this.currentData.fengkouxiaotao  = value;
                NotifyPropertyChanged("Fengkouxiaotao");
            }
        }

        /// <summary>
        /// 渣口大套
        /// </summary>
        public int? Zhakoudatao
        {
            get
            {
                return this.currentData.zhakoudatao ;
            }
            set
            {
                this.currentData.zhakoudatao  = value;
                NotifyPropertyChanged("Zhakoudatao");
            }
        }

        /// <summary>
        /// 渣口中套
        /// </summary>
        public int? Zhakouzhongtao
        {
            get
            {
                return this.currentData.zhakouzhongtao ;
            }
            set
            {
                this.currentData.zhakouzhongtao  = value;
                NotifyPropertyChanged("Zhakouzhongtao");
            }
        }

        /// <summary>
        /// 渣口小套
        /// </summary>
        public int? Zhakouxiaotao
        {
            get
            {
                return this.currentData.zhakouxiaotao;
            }
            set
            {
                this.currentData.zhakouxiaotao = value;
                NotifyPropertyChanged("Zhakouxiaotao");
            }
        }

        /// <summary>
        /// 二氧化硅
        /// </summary>
        public double? SiO2
        {
            get
            {
                return this.currentData.sio2;
            }
            set
            {
                this.currentData.sio2 = value;
                NotifyPropertyChanged("SiO2");
            }
        }

        /// <summary>
        /// 氧化钙
        /// </summary>
        public double? CaO
        {
            get
            {
                return this.currentData.cao;
            }
            set
            {
                this.currentData.cao = value;
                NotifyPropertyChanged("CaO");
            }
        }

        /// <summary>
        /// 氧化铁
        /// </summary>
        public double? FeO
        {
            get
            {
                return this.currentData.feo;
            }
            set
            {
                this.currentData.feo = value;
                NotifyPropertyChanged("FeO");
            }
        }

        /// <summary>
        /// 氧化镁
        /// </summary>
        public double? MgO
        {
            get
            {
                return this.currentData.mgo;
            }
            set
            {
                this.currentData.mgo = value;
                NotifyPropertyChanged("MgO");
            }
        }

        /// <summary>
        /// 氧化锰
        /// </summary>
        public double? MnO
        {
            get
            {
                return this.currentData.mno;
            }
            set
            {
                this.currentData.mno = value;
                NotifyPropertyChanged("MnO");
            }
        }

        /// <summary>
        /// 三氧化二铝
        /// </summary>
        /// <value></value>
        public double? Al2O3
        {
            get
            {
               return  this.currentData.al2o3;
            }
            set
            {
                this.currentData.al2o3 = value;
                NotifyPropertyChanged("Al2O3");
            }
        }

        /// <summary>
        /// 硫
        /// </summary>
        public double? S
        {
            get
            {
                return this.currentData.s;
            }
            set
            {
                this.currentData.s = value;
                NotifyPropertyChanged("S");
            }
        }

        #region INotifyPropertyChanged 成员

        public event PropertyChangedEventHandler PropertyChanged;
      
        private void NotifyPropertyChanged(String info)
        {
            BeginEdit();
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(info));
            }
        }

        #endregion

        #region IEditableObject 成员

        public void BeginEdit()
        {
            if (!change)
            {
                oldData = currentData;
                change = true;
            }
        }

        public void CancelEdit()
        {
            if (change)
            {
                currentData = oldData;
                change = false;
            }
        }

        public void EndEdit()
        {
            if (change)
            {
                Save();
                oldData = currentData;
                change = false;
            }
        }

        #endregion

        public void Init()
        {

            OracleConnection cn = new OracleConnection();
            cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
            cn.Open();
            OracleCommand cmds1 = new OracleCommand("select AL2O3,BENGLIAO,CAO,CO,CO2,FEITIE"
                                                    + ",FENGKOUDATAO,FENGKOUXIAOTAO,FENGKOUZHONGTAO,FENGWEN,JIAODINGSHUIFEN,FEO,FUYANG,GAOYASHIJIAN,JIAODING"
                                                    + ",LENGFENGLIULIANG,LIAOPI,LUDINGWENDU,LUDINGYALI,LUSHU,MGO"
                                                    + ",MNO,REFENGYALI,S,SHENLIAOPI,SHIJIAOFEN,SIO2,WAIGOUFEITIE,XUANLIAO,ZHAKOUDATAO"
                                                    + ",ZHAKOUXIAOTAO,ZHAKOUZHONGTAO,ZUOLIAO,JIAOFENSHUIFEN from xiaohao where rq=:rq and gaolu=:gaolu", cn);
            cmds1.Parameters.Add(":rq", OracleType.DateTime).Value = Riqi;
            cmds1.Parameters.Add(":gaolu", OracleType.Int32).Value = Gaolu;

            OracleDataReader dr = cmds1.ExecuteReader();
            if (dr.HasRows)
            {
                dr.Read();
                Al2O3 = dr.IsDBNull(0) ? (double?)null : dr.GetDouble(0);
                Bengliao = dr.IsDBNull(1) ? (int?)null : dr.GetInt32(1);
                CaO = dr.IsDBNull(2) ? (double?)null : dr.GetDouble(2);
                CO = dr.IsDBNull(3) ? (double?)null : dr.GetDouble(3);
                CO2 = dr.IsDBNull(4) ? (double?)null : dr.GetDouble(4);
                Feitie = dr.IsDBNull(5) ? (double?)null : dr.GetDouble(5);
                Fengkoudatao = dr.IsDBNull(6) ? (int?)null : dr.GetInt32(6);
                Fengkouxiaotao = dr.IsDBNull(7) ? (int?)null : dr.GetInt32(7);
                Fengkouzhongtao = dr.IsDBNull(8) ? (int?)null : dr.GetInt32(8);
                Fengwen = dr.IsDBNull(9) ? (double?)null : dr.GetDouble(9);
                JiaodingShuifen = dr.IsDBNull(10) ? (double?)null : dr.GetDouble(10);
                FeO = dr.IsDBNull(11) ? (double?)null : dr.GetDouble(11);
                Fuyang = dr.IsDBNull(12) ? (double?)null : dr.GetDouble(12);
                Gaoyashijian = dr.IsDBNull(13) ? (int?)null : dr.GetInt32(13);
                Jiaoding = dr.IsDBNull(14) ? (double?)null : dr.GetDouble(14);
                Lengfengliuliang = dr.IsDBNull(15) ? (double?)null : dr.GetDouble(15);
                Liaopi = dr.IsDBNull(16) ? (int?)null : dr.GetInt32(16);
                Ludingwendu = dr.IsDBNull(17) ? (double?)null : dr.GetDouble(17);
                Ludingyali = dr.IsDBNull(18) ? (double?)null : dr.GetDouble(18);
                Lushu = dr.IsDBNull(19) ? (int?)null : dr.GetInt32(19);
                MgO = dr.IsDBNull(20) ? (double?)null : dr.GetDouble(20);

                MnO = dr.IsDBNull(21) ? (double?)null : dr.GetDouble(21);
                Refengyali = dr.IsDBNull(22) ? (double?)null : dr.GetDouble(22);
                S = dr.IsDBNull(23) ? (double?)null : dr.GetDouble(23);
                Shenliaopi  = dr.IsDBNull(24) ? (int?)null : dr.GetInt32(24);
                Shijiaofen = dr.IsDBNull(25) ? (double?)null : dr.GetDouble(25);
                SiO2 = dr.IsDBNull(26) ? (double?)null : dr.GetDouble(26);
                Waigoufeitie = dr.IsDBNull(27) ? (double?)null : dr.GetDouble(27);
                Xuanliao = dr.IsDBNull(28) ? (int?)null : dr.GetInt32(28);
                Zhakoudatao = dr.IsDBNull(29) ? (int?)null : dr.GetInt32(29);
                Zhakouxiaotao = dr.IsDBNull(30) ? (int?)null : dr.GetInt32(30);
                Zhakouzhongtao = dr.IsDBNull(31) ? (int?)null : dr.GetInt32(31);
                Zuoliao = dr.IsDBNull(32) ? (int?)null : dr.GetInt32(32);
                JiaofenShuifen=dr.IsDBNull(33) ? (double?)null : dr.GetDouble(33);
                dr.Close();
            }
            else
            {
                dr.Close();
                OracleCommand cmdi1 = new OracleCommand("insert into xiaohao(RQ,GAOLU) values(:rq,:gaolu)", cn);
                cmdi1.Parameters.Add(":rq", OracleType.DateTime).Value = Riqi;
                cmdi1.Parameters.Add(":gaolu", OracleType.Int32).Value = Gaolu;
                cmdi1.ExecuteNonQuery();
            }
            cn.Close();
            cn.Dispose();
     
        }

        public void Save()
        {
            OracleConnection cn = new OracleConnection();
            cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
            cn.Open();

            OracleCommand cmdu1 = new OracleCommand();
            cmdu1.Connection = cn;
   
            cmdu1.CommandText = "update xiaohao "
                              + "set AL2O3=:AL2O3,BENGLIAO=:BENGLIAO,CAO=:CAO,CO=:CO,CO2=:CO2,FEITIE=:FEITIE"
                              + ",FENGKOUDATAO=:FENGKOUDATAO,FENGKOUXIAOTAO=:FENGKOUXIAOTAO,FENGKOUZHONGTAO=:FENGKOUZHONGTAO,FENGWEN=:FENGWEN"
                              + ",JIAODINGSHUIFEN=:JIAODINGSHUIFEN,FEO=:FEO,FUYANG=:FUYANG,GAOYASHIJIAN=:GAOYASHIJIAN,JIAODING=:JIAODING,LENGFENGLIULIANG=:LENGFENGLIULIANG"
                              + ",LIAOPI=:LIAOPI,LUDINGWENDU=:LUDINGWENDU,LUDINGYALI=:LUDINGYALI,LUSHU=:LUSHU,MGO=:MGO,MNO=:MNO,REFENGYALI=:REFENGYALI "
                              + ",S=:S,SHENLIAOPI=:SHENLIAOPI,SHIJIAOFEN=:SHIJIAOFEN,XUANLIAO=:XUANLIAO,ZHAKOUDATAO=:ZHAKOUDATAO,ZHAKOUXIAOTAO=:ZHAKOUXIAOTAO "
                              + ",ZHAKOUZHONGTAO=:ZHAKOUZHONGTAO,ZUOLIAO=:ZUOLIAO,SIO2=:SIO2,WAIGOUFEITIE=:WAIGOUFEITIE,JIAOFENSHUIFEN=:JIAOFENSHUIFEN where RQ=:RQ and GAOLU=:GAOLU";
        

            //   cmdu1.CommandText = "update ddluci "
            //      + "set DGSJ=:DGSJ,DKSJ=:DKSJ "
            //       + " where ZDSJ=:ZDSJ and GAOLU=:GAOLU";
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":AL2O3", Al2O3));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":BENGLIAO", Bengliao));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":CAO", CaO));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":CO", CO));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":CO2", CO2));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":FEITIE", Feitie));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":FENGKOUDATAO", Fengkoudatao));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":FENGKOUXIAOTAO",Fengkouxiaotao));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":FENGKOUZHONGTAO", Fengkouzhongtao));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":FENGWEN", Fengwen));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":JIAODINGSHUIFEN",JiaodingShuifen));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":FEO", FeO));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":FUYANG", Fuyang));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":GAOYASHIJIAN", Gaoyashijian));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":JIAODING", Jiaoding));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":LENGFENGLIULIANG", Lengfengliuliang));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":LIAOPI", Liaopi));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":LUDINGWENDU", Ludingwendu));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":LUDINGYALI", Ludingyali));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":LUSHU", Lushu));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":MGO", MgO));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":MNO", MnO));

            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":REFENGYALI", Refengyali));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":S", S));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":SHENLIAOPI", Shenliaopi));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":SHIJIAOFEN", Shijiaofen));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":XUANLIAO", Xuanliao));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":ZHAKOUDATAO", Zhakoudatao));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":ZHAKOUXIAOTAO", Zhakouxiaotao));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":ZHAKOUZHONGTAO", Zhakouzhongtao));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":ZUOLIAO", Zuoliao));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":SIO2", SiO2));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":WAIGOUFEITIE", Waigoufeitie));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":JIAOFENSHUIFEN", JiaofenShuifen));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":RQ", Riqi));
            cmdu1.Parameters.Add(DBConvert.CreateOracleParameter(":GAOLU",Gaolu));
   

            cmdu1.ExecuteNonQuery();
            cn.Close();
            cn.Dispose();
        }
    }
}
