using System;
using System.Data;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Reflection;
using System.Data.OracleClient;
using LTZN;

namespace ZHCDB
{
    //消耗
    public partial class Rbxiaohao : DbEntity
    {
        private string _MC = "";
        private System.Nullable<double> _BAIBAN;
        private System.Nullable<double> _ZHONGBAN;
        private System.Nullable<double> _YEBAN;
        private System.Nullable<double> _LEIJI;
        private System.Nullable<int> _GAOLU;
        private string _BEIZHU = "";
        private System.Nullable<System.DateTime> _SJ = null;
        //ConStr 

        public string MC
        {
            get
            {
                return this._MC;
            }
            set
            {
                if (!_MC.Equals(value))
                {
                    _MC = value;
                    RaisePropertyChanged("MC", !isCal);
                }

            }
        }
        public System.Nullable<double> BAIBAN
        {
            get
            {
                return this._BAIBAN;
            }
            set
            {
                if (!_BAIBAN.Equals(value))
                {
                    if (formatString == "" || value == null)
                        _BAIBAN = value;
                    else
                        _BAIBAN = double.Parse(value.Value.ToString(this.formatString));
                    RaisePropertyChanged("BAIBAN", !isCal);
                }

            }
        }
        public System.Nullable<double> ZHONGBAN
        {
            get
            {
                return this._ZHONGBAN;
            }
            set
            {
                if (!_ZHONGBAN.Equals(value))
                {
                    if (formatString == "" || value == null)
                        _ZHONGBAN = value;
                    else
                        _ZHONGBAN = double.Parse(value.Value.ToString(this.formatString));
                    RaisePropertyChanged("ZHONGBAN", !isCal);
                }

            }
        }
        public System.Nullable<double> YEBAN
        {
            get
            {
                return this._YEBAN;
            }
            set
            {
                if (!_YEBAN.Equals(value))
                {
                    if (formatString == "" || value == null)
                        _YEBAN = value;
                    else
                        _YEBAN = double.Parse(value.Value.ToString(this.formatString));
                    RaisePropertyChanged("YEBAN", !isCal);
                }

            }
        }
        public System.Nullable<double> LEIJI
        {
            get
            {
                if (isCal)
                    return this._LEIJI;
                else
                    return double.Parse(((this.YEBAN ?? 0) + (this.BAIBAN ?? 0) + (this.ZHONGBAN ?? 0)).ToString(this.formatString));
            }
            set
            {
                if (!_LEIJI.Equals(value))
                {
                    if (formatString == "" || value == null)
                        _LEIJI = value;
                    else
                        _LEIJI = double.Parse(value.Value.ToString(this.formatString));
                    RaisePropertyChanged("LEIJI", !isCal);
                }

            }
        }
        public System.Nullable<int> GAOLU
        {
            get
            {
                return this._GAOLU;
            }
            set
            {
                if (!_GAOLU.Equals(value))
                {
                    _GAOLU = value;
                    RaisePropertyChanged("GAOLU", !isCal);
                }

            }
        }
        public string BEIZHU
        {
            get
            {
                return this._BEIZHU;
            }
            set
            {
                if (!_BEIZHU.Equals(value))
                {
                    _BEIZHU = value;
                    RaisePropertyChanged("BEIZHU", !isCal);
                }

            }
        }
        public System.Nullable<System.DateTime> SJ
        {
            get
            {
                return this._SJ;
            }
            set
            {
                if (!_SJ.Equals(value))
                {
                    _SJ = value;
                    RaisePropertyChanged("SJ", !isCal);
                }

            }
        }

       
        public override void DbInsert(OracleTransaction trans)
        {
            if (isCal) return;
            //InsertSql,InsertPara,AutoSql
            if (DataState == DataStateType.Add)
            {
                OracleConnection con = trans.Connection;
                OracleCommand insertCmd = new OracleCommand();
                insertCmd.Connection = con;
                insertCmd.Transaction = trans;
                insertCmd.CommandText = "INSERT INTO RBXIAOHAO(MC,BAIBAN,ZHONGBAN,YEBAN,LEIJI,GAOLU,BEIZHU,SJ) VALUES(:MC,:BAIBAN,:ZHONGBAN,:YEBAN,:LEIJI,:GAOLU,:BEIZHU,:SJ)";
                insertCmd.Parameters.Add(":MC", OracleType.VarChar, 200).Value = this.MC;
                insertCmd.Parameters.Add(":BAIBAN", OracleType.Double, 22).Value = (object)this.BAIBAN ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHONGBAN", OracleType.Double, 22).Value = (object)this.ZHONGBAN ?? DBNull.Value;
                insertCmd.Parameters.Add(":YEBAN", OracleType.Double, 22).Value = (object)this.YEBAN ?? DBNull.Value;
                insertCmd.Parameters.Add(":LEIJI", OracleType.Double, 22).Value = (object)this.LEIJI ?? DBNull.Value;
                insertCmd.Parameters.Add(":GAOLU", OracleType.Int32, 22).Value = (object)this.GAOLU ?? DBNull.Value;
                insertCmd.Parameters.Add(":BEIZHU", OracleType.VarChar, 200).Value = this.BEIZHU;
                insertCmd.Parameters.Add(":SJ", OracleType.DateTime, 7).Value = (object)this.SJ ?? DBNull.Value;

                insertCmd.ExecuteOracleNonQuery(out RowId);
            }
        }
        public override void DbDelete(OracleTransaction trans)
        {
            if (isCal) return;
            //DelSql,DelPara
            if (DataState == DataStateType.Delete)
            {
                OracleConnection con = trans.Connection;
                OracleCommand deleteCmd = new OracleCommand();
                deleteCmd.Connection = con;
                deleteCmd.Transaction = trans;
                deleteCmd.CommandText = "DELETE FROM RBXIAOHAO WHERE ROWID='" + RowId.Value + "'";

                deleteCmd.ExecuteNonQuery();
            }

        }
        public override void DbUpdate(OracleTransaction trans)
        {
            if (isCal) return;
            //UpdateSql,UpdatePara
            if (DataState == DataStateType.Update)
            {
                OracleConnection con = trans.Connection;
                OracleCommand updateCmd = new OracleCommand();
                updateCmd.Connection = con;
                updateCmd.Transaction = trans;
                updateCmd.CommandText = "UPDATE RBXIAOHAO SET MC=:MC,BAIBAN=:BAIBAN,ZHONGBAN=:ZHONGBAN,YEBAN=:YEBAN,LEIJI=:LEIJI,GAOLU=:GAOLU,BEIZHU=:BEIZHU,SJ=:SJ WHERE ROWID='" + RowId.Value + "'";
                updateCmd.Parameters.Add(":MC", OracleType.VarChar, 200).Value = this.MC;
                updateCmd.Parameters.Add(":BAIBAN", OracleType.Double, 22).Value = (object)this.BAIBAN ?? DBNull.Value;
                updateCmd.Parameters.Add(":ZHONGBAN", OracleType.Double, 22).Value = (object)this.ZHONGBAN ?? DBNull.Value;
                updateCmd.Parameters.Add(":YEBAN", OracleType.Double, 22).Value = (object)this.YEBAN ?? DBNull.Value;
                updateCmd.Parameters.Add(":LEIJI", OracleType.Double, 22).Value = (object)this.LEIJI ?? DBNull.Value;
                updateCmd.Parameters.Add(":GAOLU", OracleType.Int32, 22).Value = (object)this.GAOLU ?? DBNull.Value;
                updateCmd.Parameters.Add(":BEIZHU", OracleType.VarChar, 200).Value = this.BEIZHU;
                updateCmd.Parameters.Add(":SJ", OracleType.DateTime, 7).Value = (object)this.SJ ?? DBNull.Value;

                updateCmd.ExecuteNonQuery();
            }


        }

        public static Rbxiaohao operator +(Rbxiaohao r1, Rbxiaohao r2)
        {
            Rbxiaohao result = new Rbxiaohao();
            result.SetCal();
            if (r1 != null && r2 != null)
            {
                result.YEBAN = (r1.YEBAN ?? 0) + (r2.YEBAN ?? 0);
                result.BAIBAN = (r1.BAIBAN ?? 0) + (r2.BAIBAN ?? 0);
                result.ZHONGBAN = (r1.ZHONGBAN ?? 0) + (r2.ZHONGBAN ?? 0);
                result.LEIJI = (r1.LEIJI ?? 0) + (r2.LEIJI ?? 0);
            }
            return result;
        }
        public static Rbxiaohao operator -(Rbxiaohao r1, Rbxiaohao r2)
        {
            Rbxiaohao result = new Rbxiaohao();
            result.SetCal();
            if (r1 != null && r2 != null)
            {
                result.YEBAN = (r1.YEBAN ?? 0) - (r2.YEBAN ?? 0);
                result.BAIBAN = (r1.BAIBAN ?? 0) - (r2.BAIBAN ?? 0);
                result.ZHONGBAN = (r1.ZHONGBAN ?? 0) - (r2.ZHONGBAN ?? 0);
                result.LEIJI = (r1.LEIJI ?? 0) - (r2.LEIJI ?? 0);
            }
            return result;
        }
        public static Rbxiaohao operator *(Rbxiaohao r1, Rbxiaohao r2)
        {
            Rbxiaohao result = new Rbxiaohao();
            result.SetCal();
            if (r1 != null && r2 != null)
            {
                result.YEBAN = (r1.YEBAN ?? 0) * (r2.YEBAN ?? 0);
                result.BAIBAN = (r1.BAIBAN ?? 0) * (r2.BAIBAN ?? 0);
                result.ZHONGBAN = (r1.ZHONGBAN ?? 0) * (r2.ZHONGBAN ?? 0);
                result.LEIJI = (r1.LEIJI ?? 0) * (r2.LEIJI ?? 0);
            }
            return result;
        }
        public static Rbxiaohao operator /(Rbxiaohao r1, Rbxiaohao r2)
        {
            Rbxiaohao result = new Rbxiaohao();
            result.SetCal();
            if (r1 != null && r2 != null)
            {
                if (r2.YEBAN != null && r2.YEBAN != 0)
                    result.YEBAN = (r1.YEBAN ?? 0) / (r2.YEBAN ?? 0);
                else
                    result.YEBAN = null;

                if (r2.BAIBAN != null && r2.BAIBAN != 0)
                    result.BAIBAN = (r1.BAIBAN ?? 0) / (r2.BAIBAN ?? 0);
                else
                    result.BAIBAN = null;

                if (r2.ZHONGBAN != null && r2.ZHONGBAN != 0)
                    result.ZHONGBAN = (r1.ZHONGBAN ?? 0) / (r2.ZHONGBAN ?? 0);
                else
                    result.ZHONGBAN = null;

                if (r2.LEIJI != null && r2.LEIJI != 0)
                    result.LEIJI = (r1.LEIJI ?? 0) / (r2.LEIJI ?? 0);
                else
                    result.LEIJI = null;
            }
            return result;
        }

        public static Rbxiaohao operator +(Rbxiaohao r1, double r2)
        {
            Rbxiaohao result = new Rbxiaohao();
            result.SetCal();
            if (r1 != null)
            {
                result.YEBAN = (r1.YEBAN ?? 0) + r2;
                result.BAIBAN = (r1.BAIBAN ?? 0) + r2;
                result.ZHONGBAN = (r1.ZHONGBAN ?? 0) + r2;
                result.LEIJI = (r1.LEIJI ?? 0) + r2;
            }
            return result;
        }
        public static Rbxiaohao operator -(Rbxiaohao r1, double r2)
        {
            Rbxiaohao result = new Rbxiaohao();
            result.SetCal();
            if (r1 != null)
            {
                result.YEBAN = (r1.YEBAN ?? 0) - r2;
                result.BAIBAN = (r1.BAIBAN ?? 0) - r2;
                result.ZHONGBAN = (r1.ZHONGBAN ?? 0) - r2;
                result.LEIJI = (r1.LEIJI ?? 0) - r2;
            }
            return result;
        }
        public static Rbxiaohao operator *(Rbxiaohao r1, double r2)
        {
            Rbxiaohao result = new Rbxiaohao();
            result.SetCal();
            if (r1 != null)
            {
                result.YEBAN = (r1.YEBAN ?? 0) * r2;
                result.BAIBAN = (r1.BAIBAN ?? 0) * r2;
                result.ZHONGBAN = (r1.ZHONGBAN ?? 0) * r2;
                result.LEIJI = (r1.LEIJI ?? 0) * r2;
            }
            return result;
        }
        public static Rbxiaohao operator /(Rbxiaohao r1, double r2)
        {
            Rbxiaohao result = new Rbxiaohao();
            result.SetCal();
            if (r1 != null)
            {
                if (r2 > 0)
                    result.YEBAN = (r1.YEBAN ?? 0) / r2;
                else
                    result.YEBAN = null;

                if (r2 > 0)
                    result.BAIBAN = (r1.BAIBAN ?? 0) / r2;
                else
                    result.BAIBAN = null;

                if (r2 > 0)
                    result.ZHONGBAN = (r1.ZHONGBAN ?? 0) / r2;
                else
                    result.ZHONGBAN = null;

                if (r2 > 0)
                    result.LEIJI = (r1.LEIJI ?? 0) / r2;
                else
                    result.LEIJI = null;
            }
            return result;
        }

        public void GetVal(Rbxiaohao other)
        {
            if (other != null)
            {
                this.YEBAN = other.YEBAN;
                this.BAIBAN = other.BAIBAN;
                this.ZHONGBAN = other.ZHONGBAN;
                this.LEIJI = other.LEIJI;
            }
            else
            {
                this.YEBAN = null;
                this.BAIBAN = null;
                this.ZHONGBAN = null;
                this.LEIJI = null;
            }
        }

        private string formatString = "N2";
        public void SetFormat(string formatStr)
        {
            this.formatString = formatStr;
        }

        private bool isCal = false;
        public void SetCal()
        {
            isCal = true;
            this.DataState = DataStateType.UnChanged;
        }
    }
    public partial class RbxiaohaoTable : DbEntityTable<Rbxiaohao>
    {
        private List<string> AllTongji = new List<string>();

        public RbxiaohaoTable()
        {
            AllTongji.Add("机烧");
            AllTongji.Add("机烧比");
            AllTongji.Add("球团");
            AllTongji.Add("球团比");
            AllTongji.Add("块矿");
            AllTongji.Add("块矿比");
            AllTongji.Add("其它原料");
            AllTongji.Add("其它比");
            AllTongji.Add("焦炭");
            AllTongji.Add("焦粒");
            AllTongji.Add("兰炭");
            AllTongji.Add("煤粉");
            AllTongji.Add("当班产量");
            AllTongji.Add("毛焦比");
            AllTongji.Add("毛焦粒比");
            AllTongji.Add("煤比");
            AllTongji.Add("毛矿比");
            AllTongji.Add("综合毛焦比");
            AllTongji.Add("综合毛燃料比");
        }

        //Rbxiaohao,ltznConnectionString
        public void Load(int gaolu, DateTime riqi)
        {
            //ConStr, EntiyType, SelectSql, SelectPara, SelectParaVal

            List<string> TempAllTongji = new List<string>();
            TempAllTongji.AddRange(AllTongji);
            beginAdjustData = true;
            this.Clear();
            hideList.Clear();
            OracleConnection con = new OracleConnection(LTZN.Properties.Settings.Default.ltznConnectionString);
            con.Open();
            OracleCommand selectCmd = new OracleCommand();
            selectCmd.Connection = con;
            selectCmd.CommandText = "SELECT MC,BAIBAN,ZHONGBAN,YEBAN,LEIJI,GAOLU,BEIZHU,SJ,ROWID FROM RBXIAOHAO WHERE GAOLU=:GAOLU AND TRUNC(SJ)=:RQ ORDER BY SJ";
            selectCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = gaolu;
            selectCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = riqi;

            OracleDataReader dr = selectCmd.ExecuteReader();
            while (dr.Read())
            {
                Rbxiaohao item = new Rbxiaohao();
                item.BeginInit();
                if (dr.IsDBNull(0)) item.MC = ""; else item.MC = dr.GetString(0);
                if (dr.IsDBNull(1)) item.BAIBAN = null; else item.BAIBAN = dr.GetDouble(1);
                if (dr.IsDBNull(2)) item.ZHONGBAN = null; else item.ZHONGBAN = dr.GetDouble(2);
                if (dr.IsDBNull(3)) item.YEBAN = null; else item.YEBAN = dr.GetDouble(3);
                if (dr.IsDBNull(4)) item.LEIJI = null; else item.LEIJI = dr.GetDouble(4);
                if (dr.IsDBNull(5)) item.GAOLU = null; else item.GAOLU = dr.GetInt32(5);
                if (dr.IsDBNull(6)) item.BEIZHU = ""; else item.BEIZHU = dr.GetString(6);
                if (dr.IsDBNull(7)) item.SJ = null; else item.SJ = dr.GetDateTime(7);
                item.RowId = dr.GetString(8);

                item.EndInit();
                item.ClearDataState();
                this.Add(item);
                hideList.Add(item);
                if (TempAllTongji.Contains(item.MC))
                    TempAllTongji.Remove(item.MC);
            }
            dr.Close();
            con.Close();
            beginAdjustData = false;
            foreach (var mc in TempAllTongji)
            {
                Rbxiaohao item = new Rbxiaohao();
                item.MC = mc;
                item.GAOLU = gaolu;
                item.SJ = riqi;
                this.Add(item);
            }
            this.hideList.Sort(Compare);
            List<Rbxiaohao> list = this.Items as List<Rbxiaohao>;
            list.Sort(Compare);
            if (riqi <= DateTime.Today)
            {
                Rbxiaohao js = getByMc("机烧");
                Rbxiaohao qt = getByMc("球团");
                Rbxiaohao kk = getByMc("块矿");
                Rbxiaohao jt = getByMc("焦炭");
                Rbxiaohao jl = getByMc("焦粒");
                Rbxiaohao lt = getByMc("兰炭");
                Rbxiaohao mf = getByMc("煤粉");
                Rbxiaohao qtyl = getByMc("其它原料");
                Rbxiaohao cl = getByMc("当班产量");
                Rbxiaohao jsb = getByMc("机烧比");
                jsb.SetCal();
           
                Rbxiaohao qtb = getByMc("球团比");
                qtb.SetCal();
                Rbxiaohao kkb = getByMc("块矿比");
                 kkb.SetCal();
                Rbxiaohao qtylb = getByMc("其它比");
                qtylb.SetCal();
                Rbxiaohao mjb = getByMc("毛焦比");
                mjb.SetCal();
                mjb.SetFormat("N0");
                Rbxiaohao mjlb = getByMc("毛焦粒比");
                mjlb.SetCal();
                mjlb.SetFormat("N0");
                Rbxiaohao mb = getByMc("煤比");
                mb.SetCal();
                mb.SetFormat("N0");
                Rbxiaohao mkb = getByMc("毛矿比");
                mkb.SetCal();
                Rbxiaohao zhmjb = getByMc("综合毛焦比");
                zhmjb.SetCal();
                zhmjb.SetFormat("N0");
                Rbxiaohao zhmrlb = getByMc("综合毛燃料比");
                zhmrlb.SetCal();
                zhmrlb.SetFormat("N0");
                jsb.GetVal((js * 100) / (js + qt + kk + qtyl));
                qtb.GetVal((qt * 100) / (js + qt + kk + qtyl));
                kkb.GetVal((kk * 100) / (js + qt + kk + qtyl));
                qtylb.GetVal((qtyl * 100) / (js + qt + kk + qtyl));
                mkb.GetVal((js + qt + kk + qtyl) / 1000 / cl);
                mjb.GetVal(jt / cl);
                mjlb.GetVal(jl / cl);
                mb.GetVal(mf / cl);
                zhmjb.GetVal((jt + (jl + lt) * 0.9 + mf * 0.8) / cl);
                zhmrlb.GetVal((jt + jl + lt + mf) / cl);
                js.PropertyChanged += delegate(object sender, PropertyChangedEventArgs e)
                {
                    jsb.GetVal(js * 100 / (js + qt + kk + qtyl));
                    mkb.GetVal((js + qt + kk + qtyl)/1000 / cl);
                };
                qt.PropertyChanged += delegate(object sender, PropertyChangedEventArgs e)
                {
                    qtb.GetVal(qt * 100 / (js + qt + kk + qtyl));
                    mkb.GetVal((js + qt + kk + qtyl)/1000 / cl);
                };
                kk.PropertyChanged += delegate(object sender, PropertyChangedEventArgs e)
                {
                    kkb.GetVal(kk * 100 / (js + qt + kk + qtyl));
                    mkb.GetVal((js + qt + kk + qtyl)/1000 / cl);
                };
                qtyl.PropertyChanged += delegate(object sender, PropertyChangedEventArgs e)
                {

                    qtylb.GetVal(qtyl * 100 / (js + qt + kk + qtyl));
                    mkb.GetVal((js + qt + kk + qtyl)/1000 / cl);
                };

                jt.PropertyChanged += delegate(object sender, PropertyChangedEventArgs e)
                {
                    mjb.GetVal(jt / cl);
                    zhmjb.GetVal((jt + (jl + lt) * 0.9 + mf * 0.8) / cl);
                    zhmrlb.GetVal((jt + jl + lt + mf)/1000 / cl);
                };
                jl.PropertyChanged += delegate(object sender, PropertyChangedEventArgs e)
                {
                    mjlb.GetVal(jl / cl);
                    zhmjb.GetVal((jt + (jl + lt) * 0.9 + mf * 0.8) / cl);
                    zhmrlb.GetVal((jt + jl + lt + mf) / cl);
                };

                lt.PropertyChanged += delegate(object sender, PropertyChangedEventArgs e)
                {
                    zhmjb.GetVal((jt + (jl + lt) * 0.9 + mf * 0.8) / cl);
                    zhmrlb.GetVal((jt + jl + lt + mf) / cl);
                };

                mf.PropertyChanged += delegate(object sender, PropertyChangedEventArgs e)
                {
                    mb.GetVal(mf / cl);
                    zhmjb.GetVal((jt + (jl + lt) * 0.9 + mf * 0.8) / cl);
                    zhmrlb.GetVal((jt + jl + lt + mf) / cl);
                };
                cl.PropertyChanged += delegate(object sender, PropertyChangedEventArgs e)
                {
                    mkb.GetVal((js + qt + kk + qtyl)/1000 / cl);
                    mjb.GetVal(jt / cl);
                    mjlb.GetVal(jl / cl);
                    mb.GetVal(mf / cl);
                    zhmjb.GetVal((jt + (jl + lt) * 0.9 + mf * 0.8) / cl);
                    zhmrlb.GetVal((jt + jl + lt + mf) / cl);
                };

                if (cl != null)
                {
                    TagQuery tq = new TagQuery();
                    Dictionary<string, double> tags = tq.queryBanChanLiang(gaolu, riqi);
                    if (tags.ContainsKey("夜班产量")) cl.YEBAN = tags["夜班产量"];
                    if (tags.ContainsKey("白班产量")) cl.BAIBAN = tags["白班产量"];
                    if (tags.ContainsKey("中班产量")) cl.ZHONGBAN = tags["中班产量"];
                }
            }
            ListChangedEventArgs arg = new ListChangedEventArgs(ListChangedType.Reset, -1);
            this.OnListChanged(arg);
        }

        //private double? JisuanBili(double? p1, double? p2)
        //{
        //    double? target;
        //    if (p2 != null && p1 != null && p2 != 0)
        //        target = double.Parse((p1.Value / p2.Value).ToString("N0"));
        //    else
        //        target = null;
        //    return target;
        //}
        //private double? JisuanBili3(double? p1, double? p2)
        //{
        //    double? target;
        //    if (p2 != null && p1 != null && p2 != 0)
        //        target = double.Parse((p1.Value * 100 / p2.Value).ToString("N2"));
        //    else
        //        target = null;
        //    return target;
        //}
        //private double? JisuanBili2(double? p1, double? p2)
        //{
        //    double? target;
        //    if (p2 != null && p1 != null && p2 != 0)
        //        target = double.Parse((p1.Value / 1000 / p2.Value).ToString("N2"));
        //    else
        //        target = null;
        //    return target;
        //}

        //private void cal毛焦比(Rbxiaohao target, Rbxiaohao p1, Rbxiaohao p2)
        //{
        //    if (target == null || p1 == null || p2 == null) return;

        //    target.YEBAN = JisuanBili(p1.YEBAN, p2.YEBAN);
        //    target.BAIBAN = JisuanBili(p1.BAIBAN, p2.BAIBAN);
        //    target.ZHONGBAN = JisuanBili(p1.ZHONGBAN, p2.ZHONGBAN);
        //    target.LEIJI = JisuanBili(p1.LEIJI, p2.LEIJI);
        //}

        //private void cal毛焦粒比(Rbxiaohao target, Rbxiaohao p1, Rbxiaohao p2)
        //{

        //    if (target == null || p1 == null || p2 == null) return;

        //    target.YEBAN = JisuanBili(p1.YEBAN, p2.YEBAN);
        //    target.BAIBAN = JisuanBili(p1.BAIBAN, p2.BAIBAN);
        //    target.ZHONGBAN = JisuanBili(p1.ZHONGBAN, p2.ZHONGBAN);
        //    target.LEIJI = JisuanBili(p1.LEIJI, p2.LEIJI);
        //}

        //private void cal煤比(Rbxiaohao target, Rbxiaohao p1, Rbxiaohao p2)
        //{
        //    if (target == null || p1 == null || p2 == null) return;

        //    target.YEBAN = JisuanBili(p1.YEBAN, p2.YEBAN);
        //    target.BAIBAN = JisuanBili(p1.BAIBAN, p2.BAIBAN);
        //    target.ZHONGBAN = JisuanBili(p1.ZHONGBAN, p2.ZHONGBAN);
        //    target.LEIJI = JisuanBili(p1.LEIJI, p2.LEIJI);
        //}

        //private void cal毛矿比(Rbxiaohao target, Rbxiaohao p1, Rbxiaohao p2, Rbxiaohao p3, Rbxiaohao p4)
        //{
        //    if (target == null || p1 == null || p2 == null || p3 == null || p4 == null) return;
        //    target.YEBAN = JisuanBili2((p1.YEBAN ?? 0) + (p2.YEBAN ?? 0) + (p3.YEBAN ?? 0), p4.YEBAN);
        //    target.BAIBAN = JisuanBili2((p1.BAIBAN ?? 0) + (p2.BAIBAN ?? 0) + (p3.BAIBAN ?? 0), p4.BAIBAN);
        //    target.ZHONGBAN = JisuanBili2((p1.ZHONGBAN ?? 0) + (p2.ZHONGBAN ?? 0) + (p3.ZHONGBAN ?? 0), p4.ZHONGBAN);
        //    target.LEIJI = JisuanBili2((p1.LEIJI ?? 0) + (p2.LEIJI ?? 0) + (p3.LEIJI ?? 0), p4.LEIJI);
        //}
        //private void cal矿比(Rbxiaohao target, Rbxiaohao p1, Rbxiaohao p2, Rbxiaohao p3, Rbxiaohao p4, Rbxiaohao p5)
        //{
        //    if (target == null || p1 == null || p2 == null || p3 == null || p4 == null || p5 == null) return;
        //    target.YEBAN = JisuanBili3((p1.YEBAN ?? 0), ((p2.YEBAN ?? 0) + (p3.YEBAN ?? 0) + (p4.YEBAN ?? 0) + (p5.YEBAN ?? 0)));
        //    target.BAIBAN = JisuanBili3((p1.BAIBAN ?? 0), ((p2.BAIBAN ?? 0) + (p3.BAIBAN ?? 0) + (p4.BAIBAN ?? 0) + (p5.BAIBAN ?? 0)));
        //    target.ZHONGBAN = JisuanBili3((p1.ZHONGBAN ?? 0), ((p2.ZHONGBAN ?? 0) + (p3.ZHONGBAN ?? 0) + (p4.ZHONGBAN ?? 0) + (p5.ZHONGBAN ?? 0)));
        //    target.LEIJI = JisuanBili3((p1.LEIJI ?? 0), ((p2.LEIJI ?? 0) + (p3.LEIJI ?? 0) + (p4.LEIJI ?? 0) + (p5.LEIJI ?? 0)));
        //}

        //private void cal综合毛焦比(Rbxiaohao target, Rbxiaohao p1, Rbxiaohao p2, Rbxiaohao p3, Rbxiaohao p4)
        //{
        //    if (target == null || p1 == null || p2 == null || p3 == null || p4 == null) return;

        //    target.YEBAN = JisuanBili((p1.YEBAN ?? 0) + (p2.YEBAN * 0.9 ?? 0) + (p3.YEBAN * 0.8 ?? 0), p4.YEBAN);
        //    target.BAIBAN = JisuanBili((p1.BAIBAN ?? 0) + (p2.BAIBAN * 0.9 ?? 0) + (p3.BAIBAN * 0.8 ?? 0), p4.BAIBAN);
        //    target.ZHONGBAN = JisuanBili((p1.ZHONGBAN ?? 0) + (p2.ZHONGBAN * 0.9 ?? 0) + (p3.ZHONGBAN * 0.8 ?? 0), p4.ZHONGBAN);
        //    target.LEIJI = JisuanBili((p1.LEIJI ?? 0) + (p2.LEIJI * 0.9 ?? 0) + (p3.LEIJI * 0.8 ?? 0), p4.LEIJI);
        //}

        //private void cal综合毛燃料比(Rbxiaohao target, Rbxiaohao p1, Rbxiaohao p2, Rbxiaohao p3, Rbxiaohao p4)
        //{
        //    if (target == null || p1 == null || p2 == null || p3 == null || p4 == null) return;

        //    target.YEBAN = JisuanBili((p1.YEBAN ?? 0) + (p2.YEBAN ?? 0) + (p3.YEBAN ?? 0), p4.YEBAN);
        //    target.BAIBAN = JisuanBili((p1.BAIBAN ?? 0) + (p2.BAIBAN ?? 0) + (p3.BAIBAN ?? 0), p4.BAIBAN);
        //    target.ZHONGBAN = JisuanBili((p1.ZHONGBAN ?? 0) + (p2.ZHONGBAN ?? 0) + (p3.ZHONGBAN ?? 0), p4.ZHONGBAN);
        //    target.LEIJI = JisuanBili((p1.LEIJI ?? 0) + (p2.LEIJI ?? 0) + (p3.LEIJI ?? 0), p4.LEIJI);
        //}

        public Rbxiaohao getByMc(string mc)
        {
            Rbxiaohao result = null;
            foreach (var item in this)
            {
                if (item.MC == mc)
                {
                    result = item;
                    break;
                }

            }
            return result;
        }
        private int Compare(Rbxiaohao x, Rbxiaohao y)
        {
            int i1 = AllTongji.FindIndex(delegate(string strItem) { return strItem == x.MC; });
            int i2 = AllTongji.FindIndex(delegate(string strItem) { return strItem == y.MC; });
            return i1.CompareTo(i2);
            //x.MC

        }

    }
}
