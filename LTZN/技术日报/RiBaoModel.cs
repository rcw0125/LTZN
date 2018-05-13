using System;
using System.Collections.Generic;
using System.Text;
using LTZN.CalFramework2;
using System.Data.OracleClient;

namespace LTZN.技术日报
{
    public class RiBaoModel : MemTagModel
    {
        public RiBaoModel()
        {
            initTags();
            initRibaoView();
        }

        /// <summary>
        /// 初始化标签
        /// </summary>
        private void initTags()
        {
            MemTag tag1 = AddStringTag("报表日期", false);

            MemTag tag2 = AddStringTag("备注",true);
   

            AddNumericTag("累计.累计天数", "", null, null, true);

            #region 全铁产量
            AddNumericTag("本日.1高炉.全铁产量", "", 2, null, true);
            AddNumericTag("本日.4高炉.全铁产量", "", 2, null, true);
            AddNumericTag("本日.5高炉.全铁产量", "", 2, null, true);
            AddNumericTag("本日.6高炉.全铁产量", "", 2, null, true);
            AddNumericTag("本日.7高炉.全铁产量", "", 2, null, true);
            AddNumericTag("本日.8高炉.全铁产量", "", 2, null, true);

            AddNumericTag("本日.456高炉合计.全铁产量", "本日.4高炉.全铁产量+本日.5高炉.全铁产量+本日.6高炉.全铁产量", 2, null, false);
            AddNumericTag("本日.78高炉合计.全铁产量", "本日.7高炉.全铁产量+本日.8高炉.全铁产量", 2, null, false);
            AddNumericTag("本日.4_8高炉合计.全铁产量", "本日.4高炉.全铁产量+本日.5高炉.全铁产量+本日.6高炉.全铁产量+本日.7高炉.全铁产量+本日.8高炉.全铁产量", 2, null, false);
            AddNumericTag("本日.全铁产量", "本日.4高炉.全铁产量+本日.5高炉.全铁产量+本日.6高炉.全铁产量+本日.7高炉.全铁产量+本日.8高炉.全铁产量+本日.1高炉.全铁产量", 2, null, false);

            AddNumericTag("累计.1高炉.全铁产量", "昨日累计.1高炉.全铁产量+本日.1高炉.全铁产量", 2, null, true);
            AddNumericTag("累计.4高炉.全铁产量", "昨日累计.4高炉.全铁产量+本日.4高炉.全铁产量", 2, null, true);
            AddNumericTag("累计.5高炉.全铁产量", "昨日累计.5高炉.全铁产量+本日.5高炉.全铁产量", 2, null, true);
            AddNumericTag("累计.6高炉.全铁产量", "昨日累计.6高炉.全铁产量+本日.6高炉.全铁产量", 2, null, true);
            AddNumericTag("累计.7高炉.全铁产量", "昨日累计.7高炉.全铁产量+本日.7高炉.全铁产量", 2, null, true);
            AddNumericTag("累计.8高炉.全铁产量", "昨日累计.8高炉.全铁产量+本日.8高炉.全铁产量", 2, null, true);

            AddNumericTag("累计.456高炉合计.全铁产量", "累计.4高炉.全铁产量+累计.5高炉.全铁产量+累计.6高炉.全铁产量", 2, null, false);
            AddNumericTag("累计.78高炉合计.全铁产量", "累计.7高炉.全铁产量+累计.8高炉.全铁产量", 2, null, false);
            AddNumericTag("累计.4_8高炉合计.全铁产量", "累计.4高炉.全铁产量+累计.5高炉.全铁产量+累计.6高炉.全铁产量+累计.7高炉.全铁产量+累计.8高炉.全铁产量", 2, null, false);
            AddNumericTag("累计.全铁产量", "累计.4高炉.全铁产量+累计.5高炉.全铁产量+累计.6高炉.全铁产量+累计.7高炉.全铁产量+累计.8高炉.全铁产量+累计.1高炉.全铁产量", 2, null, false);

            AddNumericTag("昨日累计.1高炉.全铁产量", "", 2, null, false);
            AddNumericTag("昨日累计.4高炉.全铁产量", "", 2, null, false);
            AddNumericTag("昨日累计.5高炉.全铁产量", "", 2, null, false);
            AddNumericTag("昨日累计.6高炉.全铁产量", "", 2, null, false);
            AddNumericTag("昨日累计.7高炉.全铁产量", "", 2, null, false);
            AddNumericTag("昨日累计.8高炉.全铁产量", "", 2, null, false);
            #endregion

            #region 一级铁
            AddNumericTag("本日.1高炉.一级铁", "", 2, null, true);
            AddNumericTag("本日.4高炉.一级铁", "", 2, null, true);
            AddNumericTag("本日.5高炉.一级铁", "", 2, null, true);
            AddNumericTag("本日.6高炉.一级铁", "", 2, null, true);
            AddNumericTag("本日.7高炉.一级铁", "", 2, null, true);
            AddNumericTag("本日.8高炉.一级铁", "", 2, null, true);

            AddNumericTag("本日.456高炉合计.一级铁", "本日.4高炉.一级铁+本日.5高炉.一级铁+本日.6高炉.一级铁", 2, null, false);
            AddNumericTag("本日.78高炉合计.一级铁", "本日.7高炉.一级铁+本日.8高炉.一级铁", 2, null, false);
            AddNumericTag("本日.4_8高炉合计.一级铁", "本日.4高炉.一级铁+本日.5高炉.一级铁+本日.6高炉.一级铁+本日.7高炉.一级铁+本日.8高炉.一级铁", 2, null, false);
            AddNumericTag("本日.一级铁", "本日.4高炉.一级铁+本日.5高炉.一级铁+本日.6高炉.一级铁+本日.7高炉.一级铁+本日.8高炉.一级铁+本日.1高炉.一级铁", 2, null, false);

            AddNumericTag("累计.1高炉.一级铁", "昨日累计.1高炉.一级铁+本日.1高炉.一级铁", 2, null, true);
            AddNumericTag("累计.4高炉.一级铁", "昨日累计.4高炉.一级铁+本日.4高炉.一级铁", 2, null, true);
            AddNumericTag("累计.5高炉.一级铁", "昨日累计.5高炉.一级铁+本日.5高炉.一级铁", 2, null, true);
            AddNumericTag("累计.6高炉.一级铁", "昨日累计.6高炉.一级铁+本日.6高炉.一级铁", 2, null, true);
            AddNumericTag("累计.7高炉.一级铁", "昨日累计.7高炉.一级铁+本日.7高炉.一级铁", 2, null, true);
            AddNumericTag("累计.8高炉.一级铁", "昨日累计.8高炉.一级铁+本日.8高炉.一级铁", 2, null, true);

            AddNumericTag("累计.456高炉合计.一级铁", "累计.4高炉.一级铁+累计.5高炉.一级铁+累计.6高炉.一级铁", 2, null, false);
            AddNumericTag("累计.78高炉合计.一级铁", "累计.7高炉.一级铁+累计.8高炉.一级铁", 2, null, false);
            AddNumericTag("累计.4_8高炉合计.一级铁", "累计.4高炉.一级铁+累计.5高炉.一级铁+累计.6高炉.一级铁+累计.7高炉.一级铁+累计.8高炉.一级铁", 2, null, false);
            AddNumericTag("累计.一级铁", "累计.4高炉.一级铁+累计.5高炉.一级铁+累计.6高炉.一级铁+累计.7高炉.一级铁+累计.8高炉.一级铁+累计.1高炉.一级铁", 2, null, false);

            AddNumericTag("昨日累计.1高炉.一级铁", "", 2, null, false);
            AddNumericTag("昨日累计.4高炉.一级铁", "", 2, null, false);
            AddNumericTag("昨日累计.5高炉.一级铁", "", 2, null, false);
            AddNumericTag("昨日累计.6高炉.一级铁", "", 2, null, false);
            AddNumericTag("昨日累计.7高炉.一级铁", "", 2, null, false);
            AddNumericTag("昨日累计.8高炉.一级铁", "", 2, null, false);
            #endregion

            #region 一级品率
            int dec_一级品率 = 2;

            AddNumericTag("本日.1高炉.一级品率", "本日.1高炉.一级铁*100/本日.1高炉.全铁产量", dec_一级品率, null, false);
            AddNumericTag("本日.4高炉.一级品率", "本日.4高炉.一级铁*100/本日.4高炉.全铁产量", dec_一级品率, null, false);
            AddNumericTag("本日.5高炉.一级品率", "本日.5高炉.一级铁*100/本日.5高炉.全铁产量", dec_一级品率, null, false);
            AddNumericTag("本日.6高炉.一级品率", "本日.6高炉.一级铁*100/本日.6高炉.全铁产量", dec_一级品率, null, false);
            AddNumericTag("本日.7高炉.一级品率", "本日.7高炉.一级铁*100/本日.7高炉.全铁产量", dec_一级品率, null, false);
            AddNumericTag("本日.8高炉.一级品率", "本日.8高炉.一级铁*100/本日.8高炉.全铁产量", dec_一级品率, null, false);

            AddNumericTag("本日.456高炉合计.一级品率", "本日.456高炉合计.一级铁*100/本日.456高炉合计.全铁产量", dec_一级品率, null, false);
            AddNumericTag("本日.78高炉合计.一级品率", "本日.78高炉合计.一级铁*100/本日.78高炉合计.全铁产量", dec_一级品率, null, false);
            AddNumericTag("本日.4_8高炉合计.一级品率", "本日.4_8高炉合计.一级铁*100/本日.4_8高炉合计.全铁产量", dec_一级品率, null, false);
            AddNumericTag("本日.一级品率", "本日.一级铁*100/本日.全铁产量", 0, null, false);


            AddNumericTag("累计.1高炉.一级品率", "累计.1高炉.一级铁*100/累计.1高炉.全铁产量", dec_一级品率, null, false);
            AddNumericTag("累计.4高炉.一级品率", "累计.4高炉.一级铁*100/累计.4高炉.全铁产量", dec_一级品率, null, false);
            AddNumericTag("累计.5高炉.一级品率", "累计.5高炉.一级铁*100/累计.5高炉.全铁产量", dec_一级品率, null, false);
            AddNumericTag("累计.6高炉.一级品率", "累计.6高炉.一级铁*100/累计.6高炉.全铁产量", dec_一级品率, null, false);
            AddNumericTag("累计.7高炉.一级品率", "累计.7高炉.一级铁*100/累计.7高炉.全铁产量", dec_一级品率, null, false);
            AddNumericTag("累计.8高炉.一级品率", "累计.8高炉.一级铁*100/累计.8高炉.全铁产量", dec_一级品率, null, false);

            AddNumericTag("累计.456高炉合计.一级品率", "累计.456高炉合计.一级铁*100/累计.456高炉合计.全铁产量", dec_一级品率, null, false);
            AddNumericTag("累计.78高炉合计.一级品率", "累计.78高炉合计.一级铁*100/累计.78高炉合计.全铁产量", dec_一级品率, null, false);
            AddNumericTag("累计.4_8高炉合计.一级品率", "累计.4_8高炉合计.一级铁*100/累计.4_8高炉合计.全铁产量", dec_一级品率, null, false);
            AddNumericTag("累计.一级品率", "累计.一级铁*100/累计.全铁产量", dec_一级品率, null, false);

            #endregion

            #region 合格铁
            AddNumericTag("本日.1高炉.合格铁", "", 2, null, true);
            AddNumericTag("本日.4高炉.合格铁", "", 2, null, true);
            AddNumericTag("本日.5高炉.合格铁", "", 2, null, true);
            AddNumericTag("本日.6高炉.合格铁", "", 2, null, true);
            AddNumericTag("本日.7高炉.合格铁", "", 2, null, true);
            AddNumericTag("本日.8高炉.合格铁", "", 2, null, true);

            AddNumericTag("本日.456高炉合计.合格铁", "本日.4高炉.合格铁+本日.5高炉.合格铁+本日.6高炉.合格铁", 2, null, false);
            AddNumericTag("本日.78高炉合计.合格铁", "本日.7高炉.合格铁+本日.8高炉.合格铁", 2, null, false);
            AddNumericTag("本日.4_8高炉合计.合格铁", "本日.4高炉.合格铁+本日.5高炉.合格铁+本日.6高炉.合格铁+本日.7高炉.合格铁+本日.8高炉.合格铁", 2, null, false);
            AddNumericTag("本日.合格铁", "本日.4高炉.合格铁+本日.5高炉.合格铁+本日.6高炉.合格铁+本日.7高炉.合格铁+本日.8高炉.合格铁+本日.1高炉.合格铁", 2, null, false);

            AddNumericTag("累计.1高炉.合格铁", "昨日累计.1高炉.合格铁+本日.1高炉.合格铁", 2, null, true);
            AddNumericTag("累计.4高炉.合格铁", "昨日累计.4高炉.合格铁+本日.4高炉.合格铁", 2, null, true);
            AddNumericTag("累计.5高炉.合格铁", "昨日累计.5高炉.合格铁+本日.5高炉.合格铁", 2, null, true);
            AddNumericTag("累计.6高炉.合格铁", "昨日累计.6高炉.合格铁+本日.6高炉.合格铁", 2, null, true);
            AddNumericTag("累计.7高炉.合格铁", "昨日累计.7高炉.合格铁+本日.7高炉.合格铁", 2, null, true);
            AddNumericTag("累计.8高炉.合格铁", "昨日累计.8高炉.合格铁+本日.8高炉.合格铁", 2, null, true);

            AddNumericTag("累计.456高炉合计.合格铁", "累计.4高炉.合格铁+累计.5高炉.合格铁+累计.6高炉.合格铁", 2, null, false);
            AddNumericTag("累计.78高炉合计.合格铁", "累计.7高炉.合格铁+累计.8高炉.合格铁", 2, null, false);
            AddNumericTag("累计.4_8高炉合计.合格铁", "累计.4高炉.合格铁+累计.5高炉.合格铁+累计.6高炉.合格铁+累计.7高炉.合格铁+累计.8高炉.合格铁", 2, null, false);
            AddNumericTag("累计.合格铁", "累计.4高炉.合格铁+累计.5高炉.合格铁+累计.6高炉.合格铁+累计.7高炉.合格铁+累计.8高炉.合格铁+累计.1高炉.合格铁", 2, null, false);

            AddNumericTag("昨日累计.1高炉.合格铁", "", 2, null, false);
            AddNumericTag("昨日累计.4高炉.合格铁", "", 2, null, false);
            AddNumericTag("昨日累计.5高炉.合格铁", "", 2, null, false);
            AddNumericTag("昨日累计.6高炉.合格铁", "", 2, null, false);
            AddNumericTag("昨日累计.7高炉.合格铁", "", 2, null, false);
            AddNumericTag("昨日累计.8高炉.合格铁", "", 2, null, false);
            #endregion

            #region 合格率
            int dec_合格率 = 2;

            AddNumericTag("本日.1高炉.合格率", "本日.1高炉.合格铁*100/本日.1高炉.全铁产量", dec_合格率, null, false);
            AddNumericTag("本日.4高炉.合格率", "本日.4高炉.合格铁*100/本日.4高炉.全铁产量", dec_合格率, null, false);
            AddNumericTag("本日.5高炉.合格率", "本日.5高炉.合格铁*100/本日.5高炉.全铁产量", dec_合格率, null, false);
            AddNumericTag("本日.6高炉.合格率", "本日.6高炉.合格铁*100/本日.6高炉.全铁产量", dec_合格率, null, false);
            AddNumericTag("本日.7高炉.合格率", "本日.7高炉.合格铁*100/本日.7高炉.全铁产量", dec_合格率, null, false);
            AddNumericTag("本日.8高炉.合格率", "本日.8高炉.合格铁*100/本日.8高炉.全铁产量", dec_合格率, null, false);

            AddNumericTag("本日.456高炉合计.合格率", "本日.456高炉合计.合格铁*100/本日.456高炉合计.全铁产量", dec_合格率, null, false);
            AddNumericTag("本日.78高炉合计.合格率", "本日.78高炉合计.合格铁*100/本日.78高炉合计.全铁产量", dec_合格率, null, false);
            AddNumericTag("本日.4_8高炉合计.合格率", "本日.4_8高炉合计.合格铁*100/本日.4_8高炉合计.全铁产量", dec_合格率, null, false);
            AddNumericTag("本日.合格率", "本日.合格铁*100/本日.全铁产量", dec_合格率, null, false);


            AddNumericTag("累计.1高炉.合格率", "累计.1高炉.合格铁*100/累计.1高炉.全铁产量", dec_合格率, null, false);
            AddNumericTag("累计.4高炉.合格率", "累计.4高炉.合格铁*100/累计.4高炉.全铁产量", dec_合格率, null, false);
            AddNumericTag("累计.5高炉.合格率", "累计.5高炉.合格铁*100/累计.5高炉.全铁产量", dec_合格率, null, false);
            AddNumericTag("累计.6高炉.合格率", "累计.6高炉.合格铁*100/累计.6高炉.全铁产量", dec_合格率, null, false);
            AddNumericTag("累计.7高炉.合格率", "累计.7高炉.合格铁*100/累计.7高炉.全铁产量", dec_合格率, null, false);
            AddNumericTag("累计.8高炉.合格率", "累计.8高炉.合格铁*100/累计.8高炉.全铁产量", dec_合格率, null, false);

            AddNumericTag("累计.456高炉合计.合格率", "累计.456高炉合计.合格铁*100/累计.456高炉合计.全铁产量", dec_合格率, null, false);
            AddNumericTag("累计.78高炉合计.合格率", "累计.78高炉合计.合格铁*100/累计.78高炉合计.全铁产量", dec_合格率, null, false);
            AddNumericTag("累计.4_8高炉合计.合格率", "累计.4_8高炉合计.合格铁*100/累计.4_8高炉合计.全铁产量", dec_合格率, null, false);
            AddNumericTag("累计.合格率", "累计.合格铁*100/累计.全铁产量", dec_合格率, null, false);

            #endregion

            #region 号外铁

            AddNumericTag("本日.1高炉.号外铁", "本日.1高炉.全铁产量-本日.1高炉.合格铁", null, null, false);
            AddNumericTag("本日.4高炉.号外铁", "本日.4高炉.全铁产量-本日.4高炉.合格铁", null, null, false);
            AddNumericTag("本日.5高炉.号外铁", "本日.5高炉.全铁产量-本日.5高炉.合格铁", null, null, false);
            AddNumericTag("本日.6高炉.号外铁", "本日.6高炉.全铁产量-本日.6高炉.合格铁", null, null, false);
            AddNumericTag("本日.7高炉.号外铁", "本日.7高炉.全铁产量-本日.7高炉.合格铁", null, null, false);
            AddNumericTag("本日.8高炉.号外铁", "本日.8高炉.全铁产量-本日.8高炉.合格铁", null, null, false);
            AddNumericTag("本日.456高炉合计.号外铁", "本日.4高炉.号外铁+本日.5高炉.号外铁+本日.6高炉.号外铁", 2, null, false);
            AddNumericTag("本日.78高炉合计.号外铁", "本日.7高炉.号外铁+本日.8高炉.号外铁", 2, null, false);
            AddNumericTag("本日.4_8高炉合计.号外铁", "本日.4高炉.号外铁+本日.5高炉.号外铁+本日.6高炉.号外铁+本日.7高炉.号外铁+本日.8高炉.号外铁", 2, null, false);
            AddNumericTag("本日.号外铁", "本日.4高炉.号外铁+本日.5高炉.号外铁+本日.6高炉.号外铁+本日.7高炉.号外铁+本日.8高炉.号外铁+本日.1高炉.号外铁", 2, null, false);

            AddNumericTag("累计.1高炉.号外铁", "累计.1高炉.全铁产量-累计.1高炉.合格铁", null, null, false);
            AddNumericTag("累计.4高炉.号外铁", "累计.4高炉.全铁产量-累计.4高炉.合格铁", null, null, false);
            AddNumericTag("累计.5高炉.号外铁", "累计.5高炉.全铁产量-累计.5高炉.合格铁", null, null, false);
            AddNumericTag("累计.6高炉.号外铁", "累计.6高炉.全铁产量-累计.6高炉.合格铁", null, null, false);
            AddNumericTag("累计.7高炉.号外铁", "累计.7高炉.全铁产量-累计.7高炉.合格铁", null, null, false);
            AddNumericTag("累计.8高炉.号外铁", "累计.8高炉.全铁产量-累计.8高炉.合格铁", null, null, false);
            AddNumericTag("累计.456高炉合计.号外铁", "累计.4高炉.号外铁+累计.5高炉.号外铁+累计.6高炉.号外铁", 2, null, false);
            AddNumericTag("累计.78高炉合计.号外铁", "累计.7高炉.号外铁+累计.8高炉.号外铁", 2, null, false);
            AddNumericTag("累计.4_8高炉合计.号外铁", "累计.4高炉.号外铁+累计.5高炉.号外铁+累计.6高炉.号外铁+累计.7高炉.号外铁+累计.8高炉.号外铁", 2, null, false);
            AddNumericTag("累计.号外铁", "累计.4高炉.号外铁+累计.5高炉.号外铁+累计.6高炉.号外铁+累计.7高炉.号外铁+累计.8高炉.号外铁+累计.1高炉.号外铁", 2, null, false);


            #endregion

            #region 铸造铁
            AddNumericTag("本日.1高炉.铸造铁", "", 2, null, true);
            AddNumericTag("本日.4高炉.铸造铁", "", 2, null, true);
            AddNumericTag("本日.5高炉.铸造铁", "", 2, null, true);
            AddNumericTag("本日.6高炉.铸造铁", "", 2, null, true);
            AddNumericTag("本日.7高炉.铸造铁", "", 2, null, true);
            AddNumericTag("本日.8高炉.铸造铁", "", 2, null, true);

            AddNumericTag("本日.456高炉合计.铸造铁", "本日.4高炉.铸造铁+本日.5高炉.铸造铁+本日.6高炉.铸造铁", 2, null, false);
            AddNumericTag("本日.78高炉合计.铸造铁", "本日.7高炉.铸造铁+本日.8高炉.铸造铁", 2, null, false);
            AddNumericTag("本日.4_8高炉合计.铸造铁", "本日.4高炉.铸造铁+本日.5高炉.铸造铁+本日.6高炉.铸造铁+本日.7高炉.铸造铁+本日.8高炉.铸造铁", 2, null, false);
            AddNumericTag("本日.铸造铁", "本日.4高炉.铸造铁+本日.5高炉.铸造铁+本日.6高炉.铸造铁+本日.7高炉.铸造铁+本日.8高炉.铸造铁+本日.1高炉.铸造铁", 2, null, false);

            AddNumericTag("累计.1高炉.铸造铁", "昨日累计.1高炉.铸造铁+本日.1高炉.铸造铁", 2, null, true);
            AddNumericTag("累计.4高炉.铸造铁", "昨日累计.4高炉.铸造铁+本日.4高炉.铸造铁", 2, null, true);
            AddNumericTag("累计.5高炉.铸造铁", "昨日累计.5高炉.铸造铁+本日.5高炉.铸造铁", 2, null, true);
            AddNumericTag("累计.6高炉.铸造铁", "昨日累计.6高炉.铸造铁+本日.6高炉.铸造铁", 2, null, true);
            AddNumericTag("累计.7高炉.铸造铁", "昨日累计.7高炉.铸造铁+本日.7高炉.铸造铁", 2, null, true);
            AddNumericTag("累计.8高炉.铸造铁", "昨日累计.8高炉.铸造铁+本日.8高炉.铸造铁", 2, null, true);

            AddNumericTag("累计.456高炉合计.铸造铁", "累计.4高炉.铸造铁+累计.5高炉.铸造铁+累计.6高炉.铸造铁", 2, null, false);
            AddNumericTag("累计.78高炉合计.铸造铁", "累计.7高炉.铸造铁+累计.8高炉.铸造铁", 2, null, false);
            AddNumericTag("累计.4_8高炉合计.铸造铁", "累计.4高炉.铸造铁+累计.5高炉.铸造铁+累计.6高炉.铸造铁+累计.7高炉.铸造铁+累计.8高炉.铸造铁", 2, null, false);
            AddNumericTag("累计.铸造铁", "累计.4高炉.铸造铁+累计.5高炉.铸造铁+累计.6高炉.铸造铁+累计.7高炉.铸造铁+累计.8高炉.铸造铁+累计.1高炉.铸造铁", 2, null, false);

            AddNumericTag("昨日累计.1高炉.铸造铁", "", 2, null, false);
            AddNumericTag("昨日累计.4高炉.铸造铁", "", 2, null, false);
            AddNumericTag("昨日累计.5高炉.铸造铁", "", 2, null, false);
            AddNumericTag("昨日累计.6高炉.铸造铁", "", 2, null, false);
            AddNumericTag("昨日累计.7高炉.铸造铁", "", 2, null, false);
            AddNumericTag("昨日累计.8高炉.铸造铁", "", 2, null, false);
            #endregion

            #region 机烧
            AddNumericTag("本日.1高炉.机烧", "", 2, null, true);
            AddNumericTag("本日.4高炉.机烧", "", 2, null, true);
            AddNumericTag("本日.5高炉.机烧", "", 2, null, true);
            AddNumericTag("本日.6高炉.机烧", "", 2, null, true);
            AddNumericTag("本日.7高炉.机烧", "", 2, null, true);
            AddNumericTag("本日.8高炉.机烧", "", 2, null, true);

            AddNumericTag("本日.456高炉合计.机烧", "本日.4高炉.机烧+本日.5高炉.机烧+本日.6高炉.机烧", 2, null, false);
            AddNumericTag("本日.78高炉合计.机烧", "本日.7高炉.机烧+本日.8高炉.机烧", 2, null, false);
            AddNumericTag("本日.4_8高炉合计.机烧", "本日.4高炉.机烧+本日.5高炉.机烧+本日.6高炉.机烧+本日.7高炉.机烧+本日.8高炉.机烧", 2, null, false);
            AddNumericTag("本日.机烧", "本日.4高炉.机烧+本日.5高炉.机烧+本日.6高炉.机烧+本日.7高炉.机烧+本日.8高炉.机烧+本日.1高炉.机烧", 2, null, false);

            AddNumericTag("累计.1高炉.机烧", "昨日累计.1高炉.机烧+本日.1高炉.机烧", 2, null, true);
            AddNumericTag("累计.4高炉.机烧", "昨日累计.4高炉.机烧+本日.4高炉.机烧", 2, null, true);
            AddNumericTag("累计.5高炉.机烧", "昨日累计.5高炉.机烧+本日.5高炉.机烧", 2, null, true);
            AddNumericTag("累计.6高炉.机烧", "昨日累计.6高炉.机烧+本日.6高炉.机烧", 2, null, true);
            AddNumericTag("累计.7高炉.机烧", "昨日累计.7高炉.机烧+本日.7高炉.机烧", 2, null, true);
            AddNumericTag("累计.8高炉.机烧", "昨日累计.8高炉.机烧+本日.8高炉.机烧", 2, null, true);

            AddNumericTag("累计.456高炉合计.机烧", "累计.4高炉.机烧+累计.5高炉.机烧+累计.6高炉.机烧", 2, null, false);
            AddNumericTag("累计.78高炉合计.机烧", "累计.7高炉.机烧+累计.8高炉.机烧", 2, null, false);
            AddNumericTag("累计.4_8高炉合计.机烧", "累计.4高炉.机烧+累计.5高炉.机烧+累计.6高炉.机烧+累计.7高炉.机烧+累计.8高炉.机烧", 2, null, false);
            AddNumericTag("累计.机烧", "累计.4高炉.机烧+累计.5高炉.机烧+累计.6高炉.机烧+累计.7高炉.机烧+累计.8高炉.机烧+累计.1高炉.机烧", 2, null, false);

            AddNumericTag("昨日累计.1高炉.机烧", "", 2, null, false);
            AddNumericTag("昨日累计.4高炉.机烧", "", 2, null, false);
            AddNumericTag("昨日累计.5高炉.机烧", "", 2, null, false);
            AddNumericTag("昨日累计.6高炉.机烧", "", 2, null, false);
            AddNumericTag("昨日累计.7高炉.机烧", "", 2, null, false);
            AddNumericTag("昨日累计.8高炉.机烧", "", 2, null, false);
            #endregion

            #region 球团
            AddNumericTag("本日.1高炉.球团", "", 2, null, true);
            AddNumericTag("本日.4高炉.球团", "", 2, null, true);
            AddNumericTag("本日.5高炉.球团", "", 2, null, true);
            AddNumericTag("本日.6高炉.球团", "", 2, null, true);
            AddNumericTag("本日.7高炉.球团", "", 2, null, true);
            AddNumericTag("本日.8高炉.球团", "", 2, null, true);

            AddNumericTag("本日.456高炉合计.球团", "本日.4高炉.球团+本日.5高炉.球团+本日.6高炉.球团", 2, null, false);
            AddNumericTag("本日.78高炉合计.球团", "本日.7高炉.球团+本日.8高炉.球团", 2, null, false);
            AddNumericTag("本日.4_8高炉合计.球团", "本日.4高炉.球团+本日.5高炉.球团+本日.6高炉.球团+本日.7高炉.球团+本日.8高炉.球团", 2, null, false);
            AddNumericTag("本日.球团", "本日.4高炉.球团+本日.5高炉.球团+本日.6高炉.球团+本日.7高炉.球团+本日.8高炉.球团+本日.1高炉.球团", 2, null, false);

            AddNumericTag("累计.1高炉.球团", "昨日累计.1高炉.球团+本日.1高炉.球团", 2, null, true);
            AddNumericTag("累计.4高炉.球团", "昨日累计.4高炉.球团+本日.4高炉.球团", 2, null, true);
            AddNumericTag("累计.5高炉.球团", "昨日累计.5高炉.球团+本日.5高炉.球团", 2, null, true);
            AddNumericTag("累计.6高炉.球团", "昨日累计.6高炉.球团+本日.6高炉.球团", 2, null, true);
            AddNumericTag("累计.7高炉.球团", "昨日累计.7高炉.球团+本日.7高炉.球团", 2, null, true);
            AddNumericTag("累计.8高炉.球团", "昨日累计.8高炉.球团+本日.8高炉.球团", 2, null, true);

            AddNumericTag("累计.456高炉合计.球团", "累计.4高炉.球团+累计.5高炉.球团+累计.6高炉.球团", 2, null, false);
            AddNumericTag("累计.78高炉合计.球团", "累计.7高炉.球团+累计.8高炉.球团", 2, null, false);
            AddNumericTag("累计.4_8高炉合计.球团", "累计.4高炉.球团+累计.5高炉.球团+累计.6高炉.球团+累计.7高炉.球团+累计.8高炉.球团", 2, null, false);
            AddNumericTag("累计.球团", "累计.4高炉.球团+累计.5高炉.球团+累计.6高炉.球团+累计.7高炉.球团+累计.8高炉.球团+累计.1高炉.球团", 2, null, false);

            AddNumericTag("昨日累计.1高炉.球团", "", 2, null, false);
            AddNumericTag("昨日累计.4高炉.球团", "", 2, null, false);
            AddNumericTag("昨日累计.5高炉.球团", "", 2, null, false);
            AddNumericTag("昨日累计.6高炉.球团", "", 2, null, false);
            AddNumericTag("昨日累计.7高炉.球团", "", 2, null, false);
            AddNumericTag("昨日累计.8高炉.球团", "", 2, null, false);
            #endregion

            #region 其它原料
            AddNumericTag("本日.1高炉.其它原料", "", 2, null, true);
            AddNumericTag("本日.4高炉.其它原料", "", 2, null, true);
            AddNumericTag("本日.5高炉.其它原料", "", 2, null, true);
            AddNumericTag("本日.6高炉.其它原料", "", 2, null, true);
            AddNumericTag("本日.7高炉.其它原料", "", 2, null, true);
            AddNumericTag("本日.8高炉.其它原料", "", 2, null, true);

            AddNumericTag("本日.456高炉合计.其它原料", "本日.4高炉.其它原料+本日.5高炉.其它原料+本日.6高炉.其它原料", 2, null, false);
            AddNumericTag("本日.78高炉合计.其它原料", "本日.7高炉.其它原料+本日.8高炉.其它原料", 2, null, false);
            AddNumericTag("本日.4_8高炉合计.其它原料", "本日.4高炉.其它原料+本日.5高炉.其它原料+本日.6高炉.其它原料+本日.7高炉.其它原料+本日.8高炉.其它原料", 2, null, false);
            AddNumericTag("本日.其它原料", "本日.4高炉.其它原料+本日.5高炉.其它原料+本日.6高炉.其它原料+本日.7高炉.其它原料+本日.8高炉.其它原料+本日.1高炉.其它原料", 2, null, false);

            AddNumericTag("累计.1高炉.其它原料", "昨日累计.1高炉.其它原料+本日.1高炉.其它原料", 2, null, true);
            AddNumericTag("累计.4高炉.其它原料", "昨日累计.4高炉.其它原料+本日.4高炉.其它原料", 2, null, true);
            AddNumericTag("累计.5高炉.其它原料", "昨日累计.5高炉.其它原料+本日.5高炉.其它原料", 2, null, true);
            AddNumericTag("累计.6高炉.其它原料", "昨日累计.6高炉.其它原料+本日.6高炉.其它原料", 2, null, true);
            AddNumericTag("累计.7高炉.其它原料", "昨日累计.7高炉.其它原料+本日.7高炉.其它原料", 2, null, true);
            AddNumericTag("累计.8高炉.其它原料", "昨日累计.8高炉.其它原料+本日.8高炉.其它原料", 2, null, true);

            AddNumericTag("累计.456高炉合计.其它原料", "累计.4高炉.其它原料+累计.5高炉.其它原料+累计.6高炉.其它原料", 2, null, false);
            AddNumericTag("累计.78高炉合计.其它原料", "累计.7高炉.其它原料+累计.8高炉.其它原料", 2, null, false);
            AddNumericTag("累计.4_8高炉合计.其它原料", "累计.4高炉.其它原料+累计.5高炉.其它原料+累计.6高炉.其它原料+累计.7高炉.其它原料+累计.8高炉.其它原料", 2, null, false);
            AddNumericTag("累计.其它原料", "累计.4高炉.其它原料+累计.5高炉.其它原料+累计.6高炉.其它原料+累计.7高炉.其它原料+累计.8高炉.其它原料+累计.1高炉.其它原料", 2, null, false);

            AddNumericTag("昨日累计.1高炉.其它原料", "", 2, null, false);
            AddNumericTag("昨日累计.4高炉.其它原料", "", 2, null, false);
            AddNumericTag("昨日累计.5高炉.其它原料", "", 2, null, false);
            AddNumericTag("昨日累计.6高炉.其它原料", "", 2, null, false);
            AddNumericTag("昨日累计.7高炉.其它原料", "", 2, null, false);
            AddNumericTag("昨日累计.8高炉.其它原料", "", 2, null, false);
            #endregion

            #region 块矿水份
            AddNumericTag("本日.1高炉.块矿水份", "", null, null, true);
            AddNumericTag("本日.4高炉.块矿水份", "", null, null, true);
            AddNumericTag("本日.5高炉.块矿水份", "", null, null, true);
            AddNumericTag("本日.6高炉.块矿水份", "", null, null, true);
            AddNumericTag("本日.7高炉.块矿水份", "", null, null, true);
            AddNumericTag("本日.8高炉.块矿水份", "", null, null, true);
            #endregion

            #region 块矿
            AddNumericTag("本日.1高炉.块矿", "", 2, null, true);
            AddNumericTag("本日.4高炉.块矿", "", 2, null, true);
            AddNumericTag("本日.5高炉.块矿", "", 2, null, true);
            AddNumericTag("本日.6高炉.块矿", "", 2, null, true);
            AddNumericTag("本日.7高炉.块矿", "", 2, null, true);
            AddNumericTag("本日.8高炉.块矿", "", 2, null, true);

            AddNumericTag("本日.456高炉合计.块矿", "本日.4高炉.块矿+本日.5高炉.块矿+本日.6高炉.块矿", 2, null, false);
            AddNumericTag("本日.78高炉合计.块矿", "本日.7高炉.块矿+本日.8高炉.块矿", 2, null, false);
            AddNumericTag("本日.4_8高炉合计.块矿", "本日.4高炉.块矿+本日.5高炉.块矿+本日.6高炉.块矿+本日.7高炉.块矿+本日.8高炉.块矿", 2, null, false);
            AddNumericTag("本日.块矿", "本日.4高炉.块矿+本日.5高炉.块矿+本日.6高炉.块矿+本日.7高炉.块矿+本日.8高炉.块矿+本日.1高炉.块矿", 2, null, false);

            AddNumericTag("累计.1高炉.块矿", "昨日累计.1高炉.块矿+本日.1高炉.块矿", 2, null, true);
            AddNumericTag("累计.4高炉.块矿", "昨日累计.4高炉.块矿+本日.4高炉.块矿", 2, null, true);
            AddNumericTag("累计.5高炉.块矿", "昨日累计.5高炉.块矿+本日.5高炉.块矿", 2, null, true);
            AddNumericTag("累计.6高炉.块矿", "昨日累计.6高炉.块矿+本日.6高炉.块矿", 2, null, true);
            AddNumericTag("累计.7高炉.块矿", "昨日累计.7高炉.块矿+本日.7高炉.块矿", 2, null, true);
            AddNumericTag("累计.8高炉.块矿", "昨日累计.8高炉.块矿+本日.8高炉.块矿", 2, null, true);

            AddNumericTag("累计.456高炉合计.块矿", "累计.4高炉.块矿+累计.5高炉.块矿+累计.6高炉.块矿", 2, null, false);
            AddNumericTag("累计.78高炉合计.块矿", "累计.7高炉.块矿+累计.8高炉.块矿", 2, null, false);
            AddNumericTag("累计.4_8高炉合计.块矿", "累计.4高炉.块矿+累计.5高炉.块矿+累计.6高炉.块矿+累计.7高炉.块矿+累计.8高炉.块矿", 2, null, false);
            AddNumericTag("累计.块矿", "累计.4高炉.块矿+累计.5高炉.块矿+累计.6高炉.块矿+累计.7高炉.块矿+累计.8高炉.块矿+累计.1高炉.块矿", 2, null, false);

            AddNumericTag("昨日累计.1高炉.块矿", "", 2, null, false);
            AddNumericTag("昨日累计.4高炉.块矿", "", 2, null, false);
            AddNumericTag("昨日累计.5高炉.块矿", "", 2, null, false);
            AddNumericTag("昨日累计.6高炉.块矿", "", 2, null, false);
            AddNumericTag("昨日累计.7高炉.块矿", "", 2, null, false);
            AddNumericTag("昨日累计.8高炉.块矿", "", 2, null, false);
            #endregion

            #region 原料矿总耗
            AddNumericTag("本日.1高炉.原料矿总耗", "本日.1高炉.机烧+本日.1高炉.球团+本日.1高炉.块矿+本日.1高炉.其它原料", null, null, false);
            AddNumericTag("本日.4高炉.原料矿总耗", "本日.4高炉.机烧+本日.4高炉.球团+本日.4高炉.块矿+本日.4高炉.其它原料", null, null, false);
            AddNumericTag("本日.5高炉.原料矿总耗", "本日.5高炉.机烧+本日.5高炉.球团+本日.5高炉.块矿+本日.5高炉.其它原料", null, null, false);
            AddNumericTag("本日.6高炉.原料矿总耗", "本日.6高炉.机烧+本日.6高炉.球团+本日.6高炉.块矿+本日.6高炉.其它原料", null, null, false);
            AddNumericTag("本日.7高炉.原料矿总耗", "本日.7高炉.机烧+本日.7高炉.球团+本日.7高炉.块矿+本日.7高炉.其它原料", null, null, false);
            AddNumericTag("本日.8高炉.原料矿总耗", "本日.8高炉.机烧+本日.8高炉.球团+本日.8高炉.块矿+本日.8高炉.其它原料", null, null, false);

            AddNumericTag("本日.456高炉合计.原料矿总耗", "本日.4高炉.原料矿总耗+本日.5高炉.原料矿总耗+本日.6高炉.原料矿总耗", null, null, false);
            AddNumericTag("本日.78高炉合计.原料矿总耗", "本日.7高炉.原料矿总耗+本日.8高炉.原料矿总耗", null, null, false);
            AddNumericTag("本日.4_8高炉合计.原料矿总耗", "本日.4高炉.原料矿总耗+本日.5高炉.原料矿总耗+本日.6高炉.原料矿总耗+本日.7高炉.原料矿总耗+本日.8高炉.原料矿总耗", null, null, false);
            AddNumericTag("本日.原料矿总耗", "本日.4高炉.原料矿总耗+本日.5高炉.原料矿总耗+本日.6高炉.原料矿总耗+本日.7高炉.原料矿总耗+本日.8高炉.原料矿总耗+本日.1高炉.原料矿总耗", null, null, false);

            AddNumericTag("累计.1高炉.原料矿总耗", "累计.1高炉.机烧+累计.1高炉.球团+累计.1高炉.块矿+累计.1高炉.其它原料", null, null, false);
            AddNumericTag("累计.4高炉.原料矿总耗", "累计.4高炉.机烧+累计.4高炉.球团+累计.4高炉.块矿+累计.4高炉.其它原料", null, null, false);
            AddNumericTag("累计.5高炉.原料矿总耗", "累计.5高炉.机烧+累计.5高炉.球团+累计.5高炉.块矿+累计.5高炉.其它原料", null, null, false);
            AddNumericTag("累计.6高炉.原料矿总耗", "累计.6高炉.机烧+累计.6高炉.球团+累计.6高炉.块矿+累计.6高炉.其它原料", null, null, false);
            AddNumericTag("累计.7高炉.原料矿总耗", "累计.7高炉.机烧+累计.7高炉.球团+累计.7高炉.块矿+累计.7高炉.其它原料", null, null, false);
            AddNumericTag("累计.8高炉.原料矿总耗", "累计.8高炉.机烧+累计.8高炉.球团+累计.8高炉.块矿+累计.8高炉.其它原料", null, null, false);

            AddNumericTag("累计.456高炉合计.原料矿总耗", "累计.4高炉.原料矿总耗+累计.5高炉.原料矿总耗+累计.6高炉.原料矿总耗", null, null, false);
            AddNumericTag("累计.78高炉合计.原料矿总耗", "累计.7高炉.原料矿总耗+累计.8高炉.原料矿总耗", null, null, false);
            AddNumericTag("累计.4_8高炉合计.原料矿总耗", "累计.4高炉.原料矿总耗+累计.5高炉.原料矿总耗+累计.6高炉.原料矿总耗+累计.7高炉.原料矿总耗+累计.8高炉.原料矿总耗", null, null, false);
            AddNumericTag("累计.原料矿总耗", "累计.4高炉.原料矿总耗+累计.5高炉.原料矿总耗+累计.6高炉.原料矿总耗+累计.7高炉.原料矿总耗+累计.8高炉.原料矿总耗+累计.1高炉.原料矿总耗", null, null, false);

            #endregion

            #region 原料矿单耗

            AddNumericTag("本日.1高炉.原料矿单耗", "本日.1高炉.原料矿总耗*1000/本日.1高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.4高炉.原料矿单耗", "本日.4高炉.原料矿总耗*1000/本日.4高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.5高炉.原料矿单耗", "本日.5高炉.原料矿总耗*1000/本日.5高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.6高炉.原料矿单耗", "本日.6高炉.原料矿总耗*1000/本日.6高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.7高炉.原料矿单耗", "本日.7高炉.原料矿总耗*1000/本日.7高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.8高炉.原料矿单耗", "本日.8高炉.原料矿总耗*1000/本日.8高炉.全铁产量", 0, null, false);

            AddNumericTag("本日.456高炉合计.原料矿单耗", "本日.456高炉合计.原料矿总耗*1000/本日.456高炉合计.全铁产量", 0, null, false);
            AddNumericTag("本日.78高炉合计.原料矿单耗", "本日.78高炉合计.原料矿总耗*1000/本日.78高炉合计.全铁产量", 0, null, false);
            AddNumericTag("本日.4_8高炉合计.原料矿单耗", "本日.4_8高炉合计.原料矿总耗*1000/本日.4_8高炉合计.全铁产量", 0, null, false);
            AddNumericTag("本日.原料矿单耗", "本日.原料矿总耗*1000/本日.全铁产量", 0, null, false);


            AddNumericTag("累计.1高炉.原料矿单耗", "累计.1高炉.原料矿总耗*1000/累计.1高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.4高炉.原料矿单耗", "累计.4高炉.原料矿总耗*1000/累计.4高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.5高炉.原料矿单耗", "累计.5高炉.原料矿总耗*1000/累计.5高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.6高炉.原料矿单耗", "累计.6高炉.原料矿总耗*1000/累计.6高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.7高炉.原料矿单耗", "累计.7高炉.原料矿总耗*1000/累计.7高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.8高炉.原料矿单耗", "累计.8高炉.原料矿总耗*1000/累计.8高炉.全铁产量", 0, null, false);

            AddNumericTag("累计.456高炉合计.原料矿单耗", "累计.456高炉合计.原料矿总耗*1000/累计.456高炉合计.全铁产量", 0, null, false);
            AddNumericTag("累计.78高炉合计.原料矿单耗", "累计.78高炉合计.原料矿总耗*1000/累计.78高炉合计.全铁产量", 0, null, false);
            AddNumericTag("累计.4_8高炉合计.原料矿单耗", "累计.4_8高炉合计.原料矿总耗*1000/累计.4_8高炉合计.全铁产量", 0, null, false);
            AddNumericTag("累计.原料矿单耗", "累计.原料矿总耗*1000/累计.全铁产量", 0, null, false);

            #endregion


            //AddTag("本日.1高炉.废铁", "", null, null, true);
            //AddTag("本日.1高炉.废铁单耗", "本日.1高炉.废铁*1000/本日.1高炉.全铁产量", 0, null, true);

            #region 废铁
            int? dec_废铁 = 0;
            AddNumericTag("本日.1高炉.废铁", "", dec_废铁, null, true);
            AddNumericTag("本日.4高炉.废铁", "", dec_废铁, null, true);
            AddNumericTag("本日.5高炉.废铁", "", dec_废铁, null, true);
            AddNumericTag("本日.6高炉.废铁", "", dec_废铁, null, true);
            AddNumericTag("本日.7高炉.废铁", "", dec_废铁, null, true);
            AddNumericTag("本日.8高炉.废铁", "", dec_废铁, null, true);

            AddNumericTag("本日.456高炉合计.废铁", "本日.4高炉.废铁+本日.5高炉.废铁+本日.6高炉.废铁", dec_废铁, null, false);
            AddNumericTag("本日.78高炉合计.废铁", "本日.7高炉.废铁+本日.8高炉.废铁", dec_废铁, null, false);
            AddNumericTag("本日.4_8高炉合计.废铁", "本日.4高炉.废铁+本日.5高炉.废铁+本日.6高炉.废铁+本日.7高炉.废铁+本日.8高炉.废铁", dec_废铁, null, false);
            AddNumericTag("本日.废铁", "本日.4高炉.废铁+本日.5高炉.废铁+本日.6高炉.废铁+本日.7高炉.废铁+本日.8高炉.废铁+本日.1高炉.废铁", dec_废铁, null, false);

            AddNumericTag("累计.1高炉.废铁", "昨日累计.1高炉.废铁+本日.1高炉.废铁", dec_废铁, null, true);
            AddNumericTag("累计.4高炉.废铁", "昨日累计.4高炉.废铁+本日.4高炉.废铁", dec_废铁, null, true);
            AddNumericTag("累计.5高炉.废铁", "昨日累计.5高炉.废铁+本日.5高炉.废铁", dec_废铁, null, true);
            AddNumericTag("累计.6高炉.废铁", "昨日累计.6高炉.废铁+本日.6高炉.废铁", dec_废铁, null, true);
            AddNumericTag("累计.7高炉.废铁", "昨日累计.7高炉.废铁+本日.7高炉.废铁", dec_废铁, null, true);
            AddNumericTag("累计.8高炉.废铁", "昨日累计.8高炉.废铁+本日.8高炉.废铁", dec_废铁, null, true);

            AddNumericTag("累计.456高炉合计.废铁", "累计.4高炉.废铁+累计.5高炉.废铁+累计.6高炉.废铁", dec_废铁, null, false);
            AddNumericTag("累计.78高炉合计.废铁", "累计.7高炉.废铁+累计.8高炉.废铁", dec_废铁, null, false);
            AddNumericTag("累计.4_8高炉合计.废铁", "累计.4高炉.废铁+累计.5高炉.废铁+累计.6高炉.废铁+累计.7高炉.废铁+累计.8高炉.废铁", dec_废铁, null, false);
            AddNumericTag("累计.废铁", "累计.4高炉.废铁+累计.5高炉.废铁+累计.6高炉.废铁+累计.7高炉.废铁+累计.8高炉.废铁+累计.1高炉.废铁", dec_废铁, null, false);

            AddNumericTag("昨日累计.1高炉.废铁", "", dec_废铁, null, false);
            AddNumericTag("昨日累计.4高炉.废铁", "", dec_废铁, null, false);
            AddNumericTag("昨日累计.5高炉.废铁", "", dec_废铁, null, false);
            AddNumericTag("昨日累计.6高炉.废铁", "", dec_废铁, null, false);
            AddNumericTag("昨日累计.7高炉.废铁", "", dec_废铁, null, false);
            AddNumericTag("昨日累计.8高炉.废铁", "", dec_废铁, null, false);
            #endregion

            #region 废铁单耗
            int? dec_废铁单耗 = 0;
            AddNumericTag("本日.1高炉.废铁单耗", "本日.1高炉.废铁*1000/本日.1高炉.全铁产量", dec_废铁单耗, null, false);
            AddNumericTag("本日.4高炉.废铁单耗", "本日.4高炉.废铁*1000/本日.4高炉.全铁产量", dec_废铁单耗, null, false);
            AddNumericTag("本日.5高炉.废铁单耗", "本日.5高炉.废铁*1000/本日.5高炉.全铁产量", dec_废铁单耗, null, false);
            AddNumericTag("本日.6高炉.废铁单耗", "本日.6高炉.废铁*1000/本日.6高炉.全铁产量", dec_废铁单耗, null, false);
            AddNumericTag("本日.7高炉.废铁单耗", "本日.7高炉.废铁*1000/本日.7高炉.全铁产量", dec_废铁单耗, null, false);
            AddNumericTag("本日.8高炉.废铁单耗", "本日.8高炉.废铁*1000/本日.8高炉.全铁产量", dec_废铁单耗, null, false);

            AddNumericTag("本日.456高炉合计.废铁单耗", "本日.456高炉合计.废铁*1000/本日.456高炉合计.全铁产量", dec_废铁单耗, null, false);
            AddNumericTag("本日.78高炉合计.废铁单耗", "本日.78高炉合计.废铁*1000/本日.78高炉合计.全铁产量", dec_废铁单耗, null, false);
            AddNumericTag("本日.4_8高炉合计.废铁单耗", "本日.4_8高炉合计.废铁*1000/本日.4_8高炉合计.全铁产量", dec_废铁单耗, null, false);
            AddNumericTag("本日.废铁单耗", "本日.废铁*1000/本日.全铁产量", dec_废铁单耗, null, false);


            AddNumericTag("累计.1高炉.废铁单耗", "累计.1高炉.废铁*1000/累计.1高炉.全铁产量", dec_废铁单耗, null, false);
            AddNumericTag("累计.4高炉.废铁单耗", "累计.4高炉.废铁*1000/累计.4高炉.全铁产量", dec_废铁单耗, null, false);
            AddNumericTag("累计.5高炉.废铁单耗", "累计.5高炉.废铁*1000/累计.5高炉.全铁产量", dec_废铁单耗, null, false);
            AddNumericTag("累计.6高炉.废铁单耗", "累计.6高炉.废铁*1000/累计.6高炉.全铁产量", dec_废铁单耗, null, false);
            AddNumericTag("累计.7高炉.废铁单耗", "累计.7高炉.废铁*1000/累计.7高炉.全铁产量", dec_废铁单耗, null, false);
            AddNumericTag("累计.8高炉.废铁单耗", "累计.8高炉.废铁*1000/累计.8高炉.全铁产量", dec_废铁单耗, null, false);

            AddNumericTag("累计.456高炉合计.废铁单耗", "累计.456高炉合计.废铁*1000/累计.456高炉合计.全铁产量", dec_废铁单耗, null, false);
            AddNumericTag("累计.78高炉合计.废铁单耗", "累计.78高炉合计.废铁*1000/累计.78高炉合计.全铁产量", dec_废铁单耗, null, false);
            AddNumericTag("累计.4_8高炉合计.废铁单耗", "累计.4_8高炉合计.废铁*1000/累计.4_8高炉合计.全铁产量", dec_废铁单耗, null, false);
            AddNumericTag("累计.废铁单耗", "累计.废铁*1000/累计.全铁产量", dec_废铁单耗, null, false);

            #endregion

            #region TRT发电
            AddNumericTag("本日.1高炉.TRT发电", "", null, null, true);
            AddNumericTag("本日.4高炉.TRT发电", "", null, null, true);
            AddNumericTag("本日.5高炉.TRT发电", "", null, null, true);
            AddNumericTag("本日.6高炉.TRT发电", "", null, null, true);
            AddNumericTag("本日.7高炉.TRT发电", "", null, null, true);
            AddNumericTag("本日.8高炉.TRT发电", "", null, null, true);

            AddNumericTag("本日.456高炉合计.TRT发电", "本日.4高炉.TRT发电+本日.5高炉.TRT发电+本日.6高炉.TRT发电", null, null, false);
            AddNumericTag("本日.78高炉合计.TRT发电", "本日.7高炉.TRT发电+本日.8高炉.TRT发电", null, null, false);
            AddNumericTag("本日.4_8高炉合计.TRT发电", "本日.4高炉.TRT发电+本日.5高炉.TRT发电+本日.6高炉.TRT发电+本日.7高炉.TRT发电+本日.8高炉.TRT发电", null, null, false);
            AddNumericTag("本日.TRT发电", "本日.4高炉.TRT发电+本日.5高炉.TRT发电+本日.6高炉.TRT发电+本日.7高炉.TRT发电+本日.8高炉.TRT发电+本日.1高炉.TRT发电", null, null, false);

            AddNumericTag("累计.1高炉.TRT发电", "昨日累计.1高炉.TRT发电+本日.1高炉.TRT发电", null, null, true);
            AddNumericTag("累计.4高炉.TRT发电", "昨日累计.4高炉.TRT发电+本日.4高炉.TRT发电", null, null, true);
            AddNumericTag("累计.5高炉.TRT发电", "昨日累计.5高炉.TRT发电+本日.5高炉.TRT发电", null, null, true);
            AddNumericTag("累计.6高炉.TRT发电", "昨日累计.6高炉.TRT发电+本日.6高炉.TRT发电", null, null, true);
            AddNumericTag("累计.7高炉.TRT发电", "昨日累计.7高炉.TRT发电+本日.7高炉.TRT发电", null, null, true);
            AddNumericTag("累计.8高炉.TRT发电", "昨日累计.8高炉.TRT发电+本日.8高炉.TRT发电", null, null, true);

            AddNumericTag("累计.456高炉合计.TRT发电", "累计.4高炉.TRT发电+累计.5高炉.TRT发电+累计.6高炉.TRT发电", null, null, false);
            AddNumericTag("累计.78高炉合计.TRT发电", "累计.7高炉.TRT发电+累计.8高炉.TRT发电", null, null, false);
            AddNumericTag("累计.4_8高炉合计.TRT发电", "累计.4高炉.TRT发电+累计.5高炉.TRT发电+累计.6高炉.TRT发电+累计.7高炉.TRT发电+累计.8高炉.TRT发电", null, null, false);
            AddNumericTag("累计.TRT发电", "累计.4高炉.TRT发电+累计.5高炉.TRT发电+累计.6高炉.TRT发电+累计.7高炉.TRT发电+累计.8高炉.TRT发电+累计.1高炉.TRT发电", null, null, false);

            AddNumericTag("昨日累计.1高炉.TRT发电", "", null, null, false);
            AddNumericTag("昨日累计.4高炉.TRT发电", "", null, null, false);
            AddNumericTag("昨日累计.5高炉.TRT发电", "", null, null, false);
            AddNumericTag("昨日累计.6高炉.TRT发电", "", null, null, false);
            AddNumericTag("昨日累计.7高炉.TRT发电", "", null, null, false);
            AddNumericTag("昨日累计.8高炉.TRT发电", "", null, null, false);
            #endregion

            #region 返矿
            AddNumericTag("本日.1高炉.返矿", "", null, null, true);
            AddNumericTag("本日.4高炉.返矿", "", null, null, true);
            AddNumericTag("本日.5高炉.返矿", "", null, null, true);
            AddNumericTag("本日.6高炉.返矿", "", null, null, true);
            AddNumericTag("本日.7高炉.返矿", "", null, null, true);
            AddNumericTag("本日.8高炉.返矿", "", null, null, true);

            AddNumericTag("本日.456高炉合计.返矿", "本日.4高炉.返矿+本日.5高炉.返矿+本日.6高炉.返矿", null, null, false);
            AddNumericTag("本日.78高炉合计.返矿", "本日.7高炉.返矿+本日.8高炉.返矿", null, null, false);
            AddNumericTag("本日.4_8高炉合计.返矿", "本日.4高炉.返矿+本日.5高炉.返矿+本日.6高炉.返矿+本日.7高炉.返矿+本日.8高炉.返矿", null, null, false);
            AddNumericTag("本日.返矿", "本日.4高炉.返矿+本日.5高炉.返矿+本日.6高炉.返矿+本日.7高炉.返矿+本日.8高炉.返矿+本日.1高炉.返矿", null, null, false);

            AddNumericTag("累计.1高炉.返矿", "昨日累计.1高炉.返矿+本日.1高炉.返矿", null, null, true);
            AddNumericTag("累计.4高炉.返矿", "昨日累计.4高炉.返矿+本日.4高炉.返矿", null, null, true);
            AddNumericTag("累计.5高炉.返矿", "昨日累计.5高炉.返矿+本日.5高炉.返矿", null, null, true);
            AddNumericTag("累计.6高炉.返矿", "昨日累计.6高炉.返矿+本日.6高炉.返矿", null, null, true);
            AddNumericTag("累计.7高炉.返矿", "昨日累计.7高炉.返矿+本日.7高炉.返矿", null, null, true);
            AddNumericTag("累计.8高炉.返矿", "昨日累计.8高炉.返矿+本日.8高炉.返矿", null, null, true);

            AddNumericTag("累计.456高炉合计.返矿", "累计.4高炉.返矿+累计.5高炉.返矿+累计.6高炉.返矿", null, null, false);
            AddNumericTag("累计.78高炉合计.返矿", "累计.7高炉.返矿+累计.8高炉.返矿", null, null, false);
            AddNumericTag("累计.4_8高炉合计.返矿", "累计.4高炉.返矿+累计.5高炉.返矿+累计.6高炉.返矿+累计.7高炉.返矿+累计.8高炉.返矿", null, null, false);
            AddNumericTag("累计.返矿", "累计.4高炉.返矿+累计.5高炉.返矿+累计.6高炉.返矿+累计.7高炉.返矿+累计.8高炉.返矿+累计.1高炉.返矿", null, null, false);

            AddNumericTag("昨日累计.1高炉.返矿", "", null, null, false);
            AddNumericTag("昨日累计.4高炉.返矿", "", null, null, false);
            AddNumericTag("昨日累计.5高炉.返矿", "", null, null, false);
            AddNumericTag("昨日累计.6高炉.返矿", "", null, null, false);
            AddNumericTag("昨日累计.7高炉.返矿", "", null, null, false);
            AddNumericTag("昨日累计.8高炉.返矿", "", null, null, false);
            #endregion

            #region 矿槽
            AddNumericTag("本日.1高炉.矿槽", "", null, null, true);
            AddNumericTag("本日.4高炉.矿槽", "", null, null, true);
            AddNumericTag("本日.5高炉.矿槽", "", null, null, true);
            AddNumericTag("本日.6高炉.矿槽", "", null, null, true);
            AddNumericTag("本日.7高炉.矿槽", "", null, null, true);
            AddNumericTag("本日.8高炉.矿槽", "", null, null, true);

            AddNumericTag("本日.456高炉合计.矿槽", "本日.4高炉.矿槽+本日.5高炉.矿槽+本日.6高炉.矿槽", null, null, false);
            AddNumericTag("本日.78高炉合计.矿槽", "本日.7高炉.矿槽+本日.8高炉.矿槽", null, null, false);
            AddNumericTag("本日.4_8高炉合计.矿槽", "本日.4高炉.矿槽+本日.5高炉.矿槽+本日.6高炉.矿槽+本日.7高炉.矿槽+本日.8高炉.矿槽", null, null, false);
            AddNumericTag("本日.矿槽", "本日.4高炉.矿槽+本日.5高炉.矿槽+本日.6高炉.矿槽+本日.7高炉.矿槽+本日.8高炉.矿槽+本日.1高炉.矿槽", null, null, false);

            AddNumericTag("累计.1高炉.矿槽", "昨日累计.1高炉.矿槽+本日.1高炉.矿槽", null, null, true);
            AddNumericTag("累计.4高炉.矿槽", "昨日累计.4高炉.矿槽+本日.4高炉.矿槽", null, null, true);
            AddNumericTag("累计.5高炉.矿槽", "昨日累计.5高炉.矿槽+本日.5高炉.矿槽", null, null, true);
            AddNumericTag("累计.6高炉.矿槽", "昨日累计.6高炉.矿槽+本日.6高炉.矿槽", null, null, true);
            AddNumericTag("累计.7高炉.矿槽", "昨日累计.7高炉.矿槽+本日.7高炉.矿槽", null, null, true);
            AddNumericTag("累计.8高炉.矿槽", "昨日累计.8高炉.矿槽+本日.8高炉.矿槽", null, null, true);

            AddNumericTag("累计.456高炉合计.矿槽", "累计.4高炉.矿槽+累计.5高炉.矿槽+累计.6高炉.矿槽", null, null, false);
            AddNumericTag("累计.78高炉合计.矿槽", "累计.7高炉.矿槽+累计.8高炉.矿槽", null, null, false);
            AddNumericTag("累计.4_8高炉合计.矿槽", "累计.4高炉.矿槽+累计.5高炉.矿槽+累计.6高炉.矿槽+累计.7高炉.矿槽+累计.8高炉.矿槽", null, null, false);
            AddNumericTag("累计.矿槽", "累计.4高炉.矿槽+累计.5高炉.矿槽+累计.6高炉.矿槽+累计.7高炉.矿槽+累计.8高炉.矿槽+累计.1高炉.矿槽", null, null, false);

            AddNumericTag("昨日累计.1高炉.矿槽", "", null, null, false);
            AddNumericTag("昨日累计.4高炉.矿槽", "", null, null, false);
            AddNumericTag("昨日累计.5高炉.矿槽", "", null, null, false);
            AddNumericTag("昨日累计.6高炉.矿槽", "", null, null, false);
            AddNumericTag("昨日累计.7高炉.矿槽", "", null, null, false);
            AddNumericTag("昨日累计.8高炉.矿槽", "", null, null, false);
            #endregion

            #region 焦炭水份
            AddNumericTag("本日.1高炉.焦炭水份", "", null, null, true);
            AddNumericTag("本日.4高炉.焦炭水份", "", null, null, true);
            AddNumericTag("本日.5高炉.焦炭水份", "", null, null, true);
            AddNumericTag("本日.6高炉.焦炭水份", "", null, null, true);
            AddNumericTag("本日.7高炉.焦炭水份", "", null, null, true);
            AddNumericTag("本日.8高炉.焦炭水份", "", null, null, true);
            #endregion

            #region 焦炭
            AddNumericTag("本日.1高炉.焦炭", "", 2, null, true);
            AddNumericTag("本日.4高炉.焦炭", "", 2, null, true);
            AddNumericTag("本日.5高炉.焦炭", "", 2, null, true);
            AddNumericTag("本日.6高炉.焦炭", "", 2, null, true);
            AddNumericTag("本日.7高炉.焦炭", "", 2, null, true);
            AddNumericTag("本日.8高炉.焦炭", "", 2, null, true);

            AddNumericTag("本日.456高炉合计.焦炭", "本日.4高炉.焦炭+本日.5高炉.焦炭+本日.6高炉.焦炭", 2, null, false);
            AddNumericTag("本日.78高炉合计.焦炭", "本日.7高炉.焦炭+本日.8高炉.焦炭", 2, null, false);
            AddNumericTag("本日.4_8高炉合计.焦炭", "本日.4高炉.焦炭+本日.5高炉.焦炭+本日.6高炉.焦炭+本日.7高炉.焦炭+本日.8高炉.焦炭", 2, null, false);
            AddNumericTag("本日.焦炭", "本日.4高炉.焦炭+本日.5高炉.焦炭+本日.6高炉.焦炭+本日.7高炉.焦炭+本日.8高炉.焦炭+本日.1高炉.焦炭", 2, null, false);

            AddNumericTag("累计.1高炉.焦炭", "昨日累计.1高炉.焦炭+本日.1高炉.焦炭", 2, null, true);
            AddNumericTag("累计.4高炉.焦炭", "昨日累计.4高炉.焦炭+本日.4高炉.焦炭", 2, null, true);
            AddNumericTag("累计.5高炉.焦炭", "昨日累计.5高炉.焦炭+本日.5高炉.焦炭", 2, null, true);
            AddNumericTag("累计.6高炉.焦炭", "昨日累计.6高炉.焦炭+本日.6高炉.焦炭", 2, null, true);
            AddNumericTag("累计.7高炉.焦炭", "昨日累计.7高炉.焦炭+本日.7高炉.焦炭", 2, null, true);
            AddNumericTag("累计.8高炉.焦炭", "昨日累计.8高炉.焦炭+本日.8高炉.焦炭", 2, null, true);

            AddNumericTag("累计.456高炉合计.焦炭", "累计.4高炉.焦炭+累计.5高炉.焦炭+累计.6高炉.焦炭", 2, null, false);
            AddNumericTag("累计.78高炉合计.焦炭", "累计.7高炉.焦炭+累计.8高炉.焦炭", 2, null, false);
            AddNumericTag("累计.4_8高炉合计.焦炭", "累计.4高炉.焦炭+累计.5高炉.焦炭+累计.6高炉.焦炭+累计.7高炉.焦炭+累计.8高炉.焦炭", 2, null, false);
            AddNumericTag("累计.焦炭", "累计.4高炉.焦炭+累计.5高炉.焦炭+累计.6高炉.焦炭+累计.7高炉.焦炭+累计.8高炉.焦炭+累计.1高炉.焦炭", 2, null, false);

            AddNumericTag("昨日累计.1高炉.焦炭", "", 2, null, false);
            AddNumericTag("昨日累计.4高炉.焦炭", "", 2, null, false);
            AddNumericTag("昨日累计.5高炉.焦炭", "", 2, null, false);
            AddNumericTag("昨日累计.6高炉.焦炭", "", 2, null, false);
            AddNumericTag("昨日累计.7高炉.焦炭", "", 2, null, false);
            AddNumericTag("昨日累计.8高炉.焦炭", "", 2, null, false);
            #endregion

            #region 焦粒水份
            AddNumericTag("本日.1高炉.焦粒水份", "", null, null, true);
            AddNumericTag("本日.4高炉.焦粒水份", "", null, null, true);
            AddNumericTag("本日.5高炉.焦粒水份", "", null, null, true);
            AddNumericTag("本日.6高炉.焦粒水份", "", null, null, true);
            AddNumericTag("本日.7高炉.焦粒水份", "", null, null, true);
            AddNumericTag("本日.8高炉.焦粒水份", "", null, null, true);
            #endregion

            #region 焦粒
            AddNumericTag("本日.1高炉.焦粒", "", 2, null, true);
            AddNumericTag("本日.4高炉.焦粒", "", 2, null, true);
            AddNumericTag("本日.5高炉.焦粒", "", 2, null, true);
            AddNumericTag("本日.6高炉.焦粒", "", 2, null, true);
            AddNumericTag("本日.7高炉.焦粒", "", 2, null, true);
            AddNumericTag("本日.8高炉.焦粒", "", 2, null, true);

            AddNumericTag("本日.456高炉合计.焦粒", "本日.4高炉.焦粒+本日.5高炉.焦粒+本日.6高炉.焦粒", 2, null, false);
            AddNumericTag("本日.78高炉合计.焦粒", "本日.7高炉.焦粒+本日.8高炉.焦粒", 2, null, false);
            AddNumericTag("本日.4_8高炉合计.焦粒", "本日.4高炉.焦粒+本日.5高炉.焦粒+本日.6高炉.焦粒+本日.7高炉.焦粒+本日.8高炉.焦粒", 2, null, false);
            AddNumericTag("本日.焦粒", "本日.4高炉.焦粒+本日.5高炉.焦粒+本日.6高炉.焦粒+本日.7高炉.焦粒+本日.8高炉.焦粒+本日.1高炉.焦粒", 2, null, false);

            AddNumericTag("累计.1高炉.焦粒", "昨日累计.1高炉.焦粒+本日.1高炉.焦粒", 2, null, true);
            AddNumericTag("累计.4高炉.焦粒", "昨日累计.4高炉.焦粒+本日.4高炉.焦粒", 2, null, true);
            AddNumericTag("累计.5高炉.焦粒", "昨日累计.5高炉.焦粒+本日.5高炉.焦粒", 2, null, true);
            AddNumericTag("累计.6高炉.焦粒", "昨日累计.6高炉.焦粒+本日.6高炉.焦粒", 2, null, true);
            AddNumericTag("累计.7高炉.焦粒", "昨日累计.7高炉.焦粒+本日.7高炉.焦粒", 2, null, true);
            AddNumericTag("累计.8高炉.焦粒", "昨日累计.8高炉.焦粒+本日.8高炉.焦粒", 2, null, true);

            AddNumericTag("累计.456高炉合计.焦粒", "累计.4高炉.焦粒+累计.5高炉.焦粒+累计.6高炉.焦粒", 2, null, false);
            AddNumericTag("累计.78高炉合计.焦粒", "累计.7高炉.焦粒+累计.8高炉.焦粒", 2, null, false);
            AddNumericTag("累计.4_8高炉合计.焦粒", "累计.4高炉.焦粒+累计.5高炉.焦粒+累计.6高炉.焦粒+累计.7高炉.焦粒+累计.8高炉.焦粒", 2, null, false);
            AddNumericTag("累计.焦粒", "累计.4高炉.焦粒+累计.5高炉.焦粒+累计.6高炉.焦粒+累计.7高炉.焦粒+累计.8高炉.焦粒+累计.1高炉.焦粒", 2, null, false);

            AddNumericTag("昨日累计.1高炉.焦粒", "", 2, null, false);
            AddNumericTag("昨日累计.4高炉.焦粒", "", 2, null, false);
            AddNumericTag("昨日累计.5高炉.焦粒", "", 2, null, false);
            AddNumericTag("昨日累计.6高炉.焦粒", "", 2, null, false);
            AddNumericTag("昨日累计.7高炉.焦粒", "", 2, null, false);
            AddNumericTag("昨日累计.8高炉.焦粒", "", 2, null, false);
            #endregion

            #region 焦粉水份
            AddNumericTag("本日.1高炉.焦粉水份", "", null, null, true);
            AddNumericTag("本日.4高炉.焦粉水份", "", null, null, true);
            AddNumericTag("本日.5高炉.焦粉水份", "", null, null, true);
            AddNumericTag("本日.6高炉.焦粉水份", "", null, null, true);
            AddNumericTag("本日.7高炉.焦粉水份", "", null, null, true);
            AddNumericTag("本日.8高炉.焦粉水份", "", null, null, true);
            #endregion

            #region 焦粉
            AddNumericTag("本日.1高炉.焦粉", "", 2, null, true);
            AddNumericTag("本日.4高炉.焦粉", "", 2, null, true);
            AddNumericTag("本日.5高炉.焦粉", "", 2, null, true);
            AddNumericTag("本日.6高炉.焦粉", "", 2, null, true);
            AddNumericTag("本日.7高炉.焦粉", "", 2, null, true);
            AddNumericTag("本日.8高炉.焦粉", "", 2, null, true);

            AddNumericTag("本日.456高炉合计.焦粉", "本日.4高炉.焦粉+本日.5高炉.焦粉+本日.6高炉.焦粉", 2, null, false);
            AddNumericTag("本日.78高炉合计.焦粉", "本日.7高炉.焦粉+本日.8高炉.焦粉", 2, null, false);
            AddNumericTag("本日.4_8高炉合计.焦粉", "本日.4高炉.焦粉+本日.5高炉.焦粉+本日.6高炉.焦粉+本日.7高炉.焦粉+本日.8高炉.焦粉", 2, null, false);
            AddNumericTag("本日.焦粉", "本日.4高炉.焦粉+本日.5高炉.焦粉+本日.6高炉.焦粉+本日.7高炉.焦粉+本日.8高炉.焦粉+本日.1高炉.焦粉", 2, null, false);

            AddNumericTag("累计.1高炉.焦粉", "昨日累计.1高炉.焦粉+本日.1高炉.焦粉", 2, null, true);
            AddNumericTag("累计.4高炉.焦粉", "昨日累计.4高炉.焦粉+本日.4高炉.焦粉", 2, null, true);
            AddNumericTag("累计.5高炉.焦粉", "昨日累计.5高炉.焦粉+本日.5高炉.焦粉", 2, null, true);
            AddNumericTag("累计.6高炉.焦粉", "昨日累计.6高炉.焦粉+本日.6高炉.焦粉", 2, null, true);
            AddNumericTag("累计.7高炉.焦粉", "昨日累计.7高炉.焦粉+本日.7高炉.焦粉", 2, null, true);
            AddNumericTag("累计.8高炉.焦粉", "昨日累计.8高炉.焦粉+本日.8高炉.焦粉", 2, null, true);

            AddNumericTag("累计.456高炉合计.焦粉", "累计.4高炉.焦粉+累计.5高炉.焦粉+累计.6高炉.焦粉", 2, null, false);
            AddNumericTag("累计.78高炉合计.焦粉", "累计.7高炉.焦粉+累计.8高炉.焦粉", 2, null, false);
            AddNumericTag("累计.4_8高炉合计.焦粉", "累计.4高炉.焦粉+累计.5高炉.焦粉+累计.6高炉.焦粉+累计.7高炉.焦粉+累计.8高炉.焦粉", 2, null, false);
            AddNumericTag("累计.焦粉", "累计.4高炉.焦粉+累计.5高炉.焦粉+累计.6高炉.焦粉+累计.7高炉.焦粉+累计.8高炉.焦粉+累计.1高炉.焦粉", 2, null, false);

            AddNumericTag("昨日累计.1高炉.焦粉", "", 2, null, false);
            AddNumericTag("昨日累计.4高炉.焦粉", "", 2, null, false);
            AddNumericTag("昨日累计.5高炉.焦粉", "", 2, null, false);
            AddNumericTag("昨日累计.6高炉.焦粉", "", 2, null, false);
            AddNumericTag("昨日累计.7高炉.焦粉", "", 2, null, false);
            AddNumericTag("昨日累计.8高炉.焦粉", "", 2, null, false);
            #endregion

            #region 兰炭水份
            AddNumericTag("本日.1高炉.兰炭水份", "", null, null, true);
            AddNumericTag("本日.4高炉.兰炭水份", "", null, null, true);
            AddNumericTag("本日.5高炉.兰炭水份", "", null, null, true);
            AddNumericTag("本日.6高炉.兰炭水份", "", null, null, true);
            AddNumericTag("本日.7高炉.兰炭水份", "", null, null, true);
            AddNumericTag("本日.8高炉.兰炭水份", "", null, null, true);
            #endregion

            #region 兰炭
            AddNumericTag("本日.1高炉.兰炭", "", 2, null, true);
            AddNumericTag("本日.4高炉.兰炭", "", 2, null, true);
            AddNumericTag("本日.5高炉.兰炭", "", 2, null, true);
            AddNumericTag("本日.6高炉.兰炭", "", 2, null, true);
            AddNumericTag("本日.7高炉.兰炭", "", 2, null, true);
            AddNumericTag("本日.8高炉.兰炭", "", 2, null, true);

            AddNumericTag("本日.456高炉合计.兰炭", "本日.4高炉.兰炭+本日.5高炉.兰炭+本日.6高炉.兰炭", 2, null, false);
            AddNumericTag("本日.78高炉合计.兰炭", "本日.7高炉.兰炭+本日.8高炉.兰炭", 2, null, false);
            AddNumericTag("本日.4_8高炉合计.兰炭", "本日.4高炉.兰炭+本日.5高炉.兰炭+本日.6高炉.兰炭+本日.7高炉.兰炭+本日.8高炉.兰炭", 2, null, false);
            AddNumericTag("本日.兰炭", "本日.4高炉.兰炭+本日.5高炉.兰炭+本日.6高炉.兰炭+本日.7高炉.兰炭+本日.8高炉.兰炭+本日.1高炉.兰炭", 2, null, false);

            AddNumericTag("累计.1高炉.兰炭", "昨日累计.1高炉.兰炭+本日.1高炉.兰炭", 2, null, true);
            AddNumericTag("累计.4高炉.兰炭", "昨日累计.4高炉.兰炭+本日.4高炉.兰炭", 2, null, true);
            AddNumericTag("累计.5高炉.兰炭", "昨日累计.5高炉.兰炭+本日.5高炉.兰炭", 2, null, true);
            AddNumericTag("累计.6高炉.兰炭", "昨日累计.6高炉.兰炭+本日.6高炉.兰炭", 2, null, true);
            AddNumericTag("累计.7高炉.兰炭", "昨日累计.7高炉.兰炭+本日.7高炉.兰炭", 2, null, true);
            AddNumericTag("累计.8高炉.兰炭", "昨日累计.8高炉.兰炭+本日.8高炉.兰炭", 2, null, true);

            AddNumericTag("累计.456高炉合计.兰炭", "累计.4高炉.兰炭+累计.5高炉.兰炭+累计.6高炉.兰炭", 2, null, false);
            AddNumericTag("累计.78高炉合计.兰炭", "累计.7高炉.兰炭+累计.8高炉.兰炭", 2, null, false);
            AddNumericTag("累计.4_8高炉合计.兰炭", "累计.4高炉.兰炭+累计.5高炉.兰炭+累计.6高炉.兰炭+累计.7高炉.兰炭+累计.8高炉.兰炭", 2, null, false);
            AddNumericTag("累计.兰炭", "累计.4高炉.兰炭+累计.5高炉.兰炭+累计.6高炉.兰炭+累计.7高炉.兰炭+累计.8高炉.兰炭+累计.1高炉.兰炭", 2, null, false);

            AddNumericTag("昨日累计.1高炉.兰炭", "", 2, null, false);
            AddNumericTag("昨日累计.4高炉.兰炭", "", 2, null, false);
            AddNumericTag("昨日累计.5高炉.兰炭", "", 2, null, false);
            AddNumericTag("昨日累计.6高炉.兰炭", "", 2, null, false);
            AddNumericTag("昨日累计.7高炉.兰炭", "", 2, null, false);
            AddNumericTag("昨日累计.8高炉.兰炭", "", 2, null, false);
            #endregion

            #region 煤粉水份
            AddNumericTag("本日.1高炉.煤粉水份", "", null, null, true);
            AddNumericTag("本日.4高炉.煤粉水份", "", null, null, true);
            AddNumericTag("本日.5高炉.煤粉水份", "", null, null, true);
            AddNumericTag("本日.6高炉.煤粉水份", "", null, null, true);
            AddNumericTag("本日.7高炉.煤粉水份", "", null, null, true);
            AddNumericTag("本日.8高炉.煤粉水份", "", null, null, true);
            #endregion

            #region 煤粉
            AddNumericTag("本日.1高炉.煤粉", "", 2, null, true);
            AddNumericTag("本日.4高炉.煤粉", "", 2, null, true);
            AddNumericTag("本日.5高炉.煤粉", "", 2, null, true);
            AddNumericTag("本日.6高炉.煤粉", "", 2, null, true);
            AddNumericTag("本日.7高炉.煤粉", "", 2, null, true);
            AddNumericTag("本日.8高炉.煤粉", "", 2, null, true);

            AddNumericTag("本日.456高炉合计.煤粉", "本日.4高炉.煤粉+本日.5高炉.煤粉+本日.6高炉.煤粉", 2, null, false);
            AddNumericTag("本日.78高炉合计.煤粉", "本日.7高炉.煤粉+本日.8高炉.煤粉", 2, null, false);
            AddNumericTag("本日.4_8高炉合计.煤粉", "本日.4高炉.煤粉+本日.5高炉.煤粉+本日.6高炉.煤粉+本日.7高炉.煤粉+本日.8高炉.煤粉", 2, null, false);
            AddNumericTag("本日.煤粉", "本日.4高炉.煤粉+本日.5高炉.煤粉+本日.6高炉.煤粉+本日.7高炉.煤粉+本日.8高炉.煤粉+本日.1高炉.煤粉", 2, null, false);

            AddNumericTag("累计.1高炉.煤粉", "昨日累计.1高炉.煤粉+本日.1高炉.煤粉", 2, null, true);
            AddNumericTag("累计.4高炉.煤粉", "昨日累计.4高炉.煤粉+本日.4高炉.煤粉", 2, null, true);
            AddNumericTag("累计.5高炉.煤粉", "昨日累计.5高炉.煤粉+本日.5高炉.煤粉", 2, null, true);
            AddNumericTag("累计.6高炉.煤粉", "昨日累计.6高炉.煤粉+本日.6高炉.煤粉", 2, null, true);
            AddNumericTag("累计.7高炉.煤粉", "昨日累计.7高炉.煤粉+本日.7高炉.煤粉", 2, null, true);
            AddNumericTag("累计.8高炉.煤粉", "昨日累计.8高炉.煤粉+本日.8高炉.煤粉", 2, null, true);

            AddNumericTag("累计.456高炉合计.煤粉", "累计.4高炉.煤粉+累计.5高炉.煤粉+累计.6高炉.煤粉", 2, null, false);
            AddNumericTag("累计.78高炉合计.煤粉", "累计.7高炉.煤粉+累计.8高炉.煤粉", 2, null, false);
            AddNumericTag("累计.4_8高炉合计.煤粉", "累计.4高炉.煤粉+累计.5高炉.煤粉+累计.6高炉.煤粉+累计.7高炉.煤粉+累计.8高炉.煤粉", 2, null, false);
            AddNumericTag("累计.煤粉", "累计.4高炉.煤粉+累计.5高炉.煤粉+累计.6高炉.煤粉+累计.7高炉.煤粉+累计.8高炉.煤粉+累计.1高炉.煤粉", 2, null, false);

            AddNumericTag("昨日累计.1高炉.煤粉", "", 2, null, false);
            AddNumericTag("昨日累计.4高炉.煤粉", "", 2, null, false);
            AddNumericTag("昨日累计.5高炉.煤粉", "", 2, null, false);
            AddNumericTag("昨日累计.6高炉.煤粉", "", 2, null, false);
            AddNumericTag("昨日累计.7高炉.煤粉", "", 2, null, false);
            AddNumericTag("昨日累计.8高炉.煤粉", "", 2, null, false);
            #endregion

            #region 煤粉单耗

            AddNumericTag("本日.1高炉.煤粉单耗", "本日.1高炉.煤粉*1000/本日.1高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.4高炉.煤粉单耗", "本日.4高炉.煤粉*1000/本日.4高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.5高炉.煤粉单耗", "本日.5高炉.煤粉*1000/本日.5高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.6高炉.煤粉单耗", "本日.6高炉.煤粉*1000/本日.6高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.7高炉.煤粉单耗", "本日.7高炉.煤粉*1000/本日.7高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.8高炉.煤粉单耗", "本日.8高炉.煤粉*1000/本日.8高炉.全铁产量", 0, null, false);

            AddNumericTag("本日.456高炉合计.煤粉单耗", "本日.456高炉合计.煤粉*1000/本日.456高炉合计.全铁产量", 0, null, false);
            AddNumericTag("本日.78高炉合计.煤粉单耗", "本日.78高炉合计.煤粉*1000/本日.78高炉合计.全铁产量", 0, null, false);
            AddNumericTag("本日.4_8高炉合计.煤粉单耗", "本日.4_8高炉合计.煤粉*1000/本日.4_8高炉合计.全铁产量", 0, null, false);
            AddNumericTag("本日.煤粉单耗", "本日.煤粉*1000/本日.全铁产量", 0, null, false);


            AddNumericTag("累计.1高炉.煤粉单耗", "累计.1高炉.煤粉*1000/累计.1高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.4高炉.煤粉单耗", "累计.4高炉.煤粉*1000/累计.4高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.5高炉.煤粉单耗", "累计.5高炉.煤粉*1000/累计.5高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.6高炉.煤粉单耗", "累计.6高炉.煤粉*1000/累计.6高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.7高炉.煤粉单耗", "累计.7高炉.煤粉*1000/累计.7高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.8高炉.煤粉单耗", "累计.8高炉.煤粉*1000/累计.8高炉.全铁产量", 0, null, false);

            AddNumericTag("累计.456高炉合计.煤粉单耗", "累计.456高炉合计.煤粉*1000/累计.456高炉合计.全铁产量", 0, null, false);
            AddNumericTag("累计.78高炉合计.煤粉单耗", "累计.78高炉合计.煤粉*1000/累计.78高炉合计.全铁产量", 0, null, false);
            AddNumericTag("累计.4_8高炉合计.煤粉单耗", "累计.4_8高炉合计.煤粉*1000/累计.4_8高炉合计.全铁产量", 0, null, false);
            AddNumericTag("累计.煤粉单耗", "累计.煤粉*1000/累计.全铁产量", 0, null, false);

            #endregion


            #region 干焦丁
            int? dec_干焦丁 = 0;
            AddNumericTag("本日.1高炉.干焦丁", "本日.1高炉.焦粒*(100-本日.1高炉.焦粒水份)/100", dec_干焦丁, null, true);
            AddNumericTag("本日.4高炉.干焦丁", "本日.4高炉.焦粒*(100-本日.4高炉.焦粒水份)/100", dec_干焦丁, null, true);
            AddNumericTag("本日.5高炉.干焦丁", "本日.5高炉.焦粒*(100-本日.5高炉.焦粒水份)/100", dec_干焦丁, null, true);
            AddNumericTag("本日.6高炉.干焦丁", "本日.6高炉.焦粒*(100-本日.6高炉.焦粒水份)/100", dec_干焦丁, null, true);
            AddNumericTag("本日.7高炉.干焦丁", "本日.7高炉.焦粒*(100-本日.7高炉.焦粒水份)/100", dec_干焦丁, null, true);
            AddNumericTag("本日.8高炉.干焦丁", "本日.8高炉.焦粒*(100-本日.8高炉.焦粒水份)/100", dec_干焦丁, null, true);

            AddNumericTag("本日.456高炉合计.干焦丁", "本日.4高炉.干焦丁+本日.5高炉.干焦丁+本日.6高炉.干焦丁", dec_干焦丁, null, false);
            AddNumericTag("本日.78高炉合计.干焦丁", "本日.7高炉.干焦丁+本日.8高炉.干焦丁", dec_干焦丁, null, false);
            AddNumericTag("本日.4_8高炉合计.干焦丁", "本日.4高炉.干焦丁+本日.5高炉.干焦丁+本日.6高炉.干焦丁+本日.7高炉.干焦丁+本日.8高炉.干焦丁", dec_干焦丁, null, false);
            AddNumericTag("本日.干焦丁", "本日.4高炉.干焦丁+本日.5高炉.干焦丁+本日.6高炉.干焦丁+本日.7高炉.干焦丁+本日.8高炉.干焦丁+本日.1高炉.干焦丁", dec_干焦丁, null, false);

            AddNumericTag("累计.1高炉.干焦丁", "昨日累计.1高炉.干焦丁+本日.1高炉.干焦丁", dec_干焦丁, null, true);
            AddNumericTag("累计.4高炉.干焦丁", "昨日累计.4高炉.干焦丁+本日.4高炉.干焦丁", dec_干焦丁, null, true);
            AddNumericTag("累计.5高炉.干焦丁", "昨日累计.5高炉.干焦丁+本日.5高炉.干焦丁", dec_干焦丁, null, true);
            AddNumericTag("累计.6高炉.干焦丁", "昨日累计.6高炉.干焦丁+本日.6高炉.干焦丁", dec_干焦丁, null, true);
            AddNumericTag("累计.7高炉.干焦丁", "昨日累计.7高炉.干焦丁+本日.7高炉.干焦丁", dec_干焦丁, null, true);
            AddNumericTag("累计.8高炉.干焦丁", "昨日累计.8高炉.干焦丁+本日.8高炉.干焦丁", dec_干焦丁, null, true);

            AddNumericTag("累计.456高炉合计.干焦丁", "累计.4高炉.干焦丁+累计.5高炉.干焦丁+累计.6高炉.干焦丁", dec_干焦丁, null, false);
            AddNumericTag("累计.78高炉合计.干焦丁", "累计.7高炉.干焦丁+累计.8高炉.干焦丁", dec_干焦丁, null, false);
            AddNumericTag("累计.4_8高炉合计.干焦丁", "累计.4高炉.干焦丁+累计.5高炉.干焦丁+累计.6高炉.干焦丁+累计.7高炉.干焦丁+累计.8高炉.干焦丁", dec_干焦丁, null, false);
            AddNumericTag("累计.干焦丁", "累计.4高炉.干焦丁+累计.5高炉.干焦丁+累计.6高炉.干焦丁+累计.7高炉.干焦丁+累计.8高炉.干焦丁+累计.1高炉.干焦丁", dec_干焦丁, null, false);

            AddNumericTag("昨日累计.1高炉.干焦丁", "", dec_干焦丁, null, false);
            AddNumericTag("昨日累计.4高炉.干焦丁", "", dec_干焦丁, null, false);
            AddNumericTag("昨日累计.5高炉.干焦丁", "", dec_干焦丁, null, false);
            AddNumericTag("昨日累计.6高炉.干焦丁", "", dec_干焦丁, null, false);
            AddNumericTag("昨日累计.7高炉.干焦丁", "", dec_干焦丁, null, false);
            AddNumericTag("昨日累计.8高炉.干焦丁", "", dec_干焦丁, null, false);
            #endregion

            //湿焦丁比
            #region 焦丁比

            AddNumericTag("本日.1高炉.焦丁比", "(本日.1高炉.焦粒+本日.1高炉.兰炭)*1000/本日.1高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.4高炉.焦丁比", "(本日.4高炉.焦粒+本日.4高炉.兰炭)*1000/本日.4高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.5高炉.焦丁比", "(本日.5高炉.焦粒+本日.5高炉.兰炭)*1000/本日.5高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.6高炉.焦丁比", "(本日.6高炉.焦粒+本日.6高炉.兰炭)*1000/本日.6高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.7高炉.焦丁比", "(本日.7高炉.焦粒+本日.7高炉.兰炭)*1000/本日.7高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.8高炉.焦丁比", "(本日.8高炉.焦粒+本日.8高炉.兰炭)*1000/本日.8高炉.全铁产量", 0, null, false);

            AddNumericTag("本日.456高炉合计.焦丁比", "(本日.456高炉合计.焦粒+本日.456高炉合计.兰炭)*1000/本日.456高炉合计.全铁产量", 0, null, false);
            AddNumericTag("本日.78高炉合计.焦丁比", "(本日.78高炉合计.焦粒+本日.78高炉合计.兰炭)*1000/本日.78高炉合计.全铁产量", 0, null, false);
            AddNumericTag("本日.4_8高炉合计.焦丁比", "(本日.4_8高炉合计.焦粒+本日.4_8高炉合计.兰炭)*1000/本日.4_8高炉合计.全铁产量", 0, null, false);
            AddNumericTag("本日.焦丁比", "(本日.焦粒+本日.兰炭)*1000/本日.全铁产量", 0, null, false);


            AddNumericTag("累计.1高炉.焦丁比", "(累计.1高炉.焦粒+累计.1高炉.兰炭)*1000/累计.1高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.4高炉.焦丁比", "(累计.4高炉.焦粒+累计.4高炉.兰炭)*1000/累计.4高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.5高炉.焦丁比", "(累计.5高炉.焦粒+累计.5高炉.兰炭)*1000/累计.5高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.6高炉.焦丁比", "(累计.6高炉.焦粒+累计.6高炉.兰炭)*1000/累计.6高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.7高炉.焦丁比", "(累计.7高炉.焦粒+累计.7高炉.兰炭)*1000/累计.7高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.8高炉.焦丁比", "(累计.8高炉.焦粒+累计.8高炉.兰炭)*1000/累计.8高炉.全铁产量", 0, null, false);

            AddNumericTag("累计.456高炉合计.焦丁比", "(累计.456高炉合计.焦粒+累计.456高炉合计.兰炭)*1000/累计.456高炉合计.全铁产量", 0, null, false);
            AddNumericTag("累计.78高炉合计.焦丁比", "(累计.78高炉合计.焦粒+累计.78高炉合计.兰炭)*1000/累计.78高炉合计.全铁产量", 0, null, false);
            AddNumericTag("累计.4_8高炉合计.焦丁比", "(累计.4_8高炉合计.焦粒+累计.4_8高炉合计.兰炭)*1000/累计.4_8高炉合计.全铁产量", 0, null, false);
            AddNumericTag("累计.焦丁比", "(累计.焦粒+累计.兰炭)*1000/累计.全铁产量", 0, null, false);

            #endregion

            //燃料比
            #region 燃料、燃料比

            AddNumericTag("本日.1高炉.燃料", "本日.1高炉.焦炭+本日.1高炉.焦粒+本日.1高炉.煤粉+本日.1高炉.兰炭", null, null, false);
            AddNumericTag("本日.1高炉.燃料比", "本日.1高炉.燃料*1000/本日.1高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.4高炉.燃料", "本日.4高炉.焦炭+本日.4高炉.焦粒+本日.4高炉.煤粉+本日.4高炉.兰炭", null, null, false);
            AddNumericTag("本日.4高炉.燃料比", "本日.4高炉.燃料*1000/本日.4高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.5高炉.燃料", "本日.5高炉.焦炭+本日.5高炉.焦粒+本日.5高炉.煤粉+本日.5高炉.兰炭", null, null, false);
            AddNumericTag("本日.5高炉.燃料比", "本日.5高炉.燃料*1000/本日.5高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.6高炉.燃料", "本日.6高炉.焦炭+本日.6高炉.焦粒+本日.6高炉.煤粉+本日.6高炉.兰炭", null, null, false);
            AddNumericTag("本日.6高炉.燃料比", "本日.6高炉.燃料*1000/本日.6高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.7高炉.燃料", "本日.7高炉.焦炭+本日.7高炉.焦粒+本日.7高炉.煤粉+本日.7高炉.兰炭", null, null, false);
            AddNumericTag("本日.7高炉.燃料比", "本日.7高炉.燃料*1000/本日.7高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.8高炉.燃料", "本日.8高炉.焦炭+本日.8高炉.焦粒+本日.8高炉.煤粉+本日.8高炉.兰炭", null, null, false);
            AddNumericTag("本日.8高炉.燃料比", "本日.8高炉.燃料*1000/本日.8高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.456高炉合计.燃料", "本日.4高炉.燃料+本日.5高炉.燃料+本日.6高炉.燃料", null, null, false);
            AddNumericTag("本日.456高炉合计.燃料比", "本日.456高炉合计.燃料*1000/本日.456高炉合计.全铁产量", 0, null, false);
            AddNumericTag("本日.78高炉合计.燃料", "本日.7高炉.燃料+本日.8高炉.燃料", null, null, false);
            AddNumericTag("本日.78高炉合计.燃料比", "本日.78高炉合计.燃料*1000/本日.78高炉合计.全铁产量", 0, null, false);
            AddNumericTag("本日.4_8高炉合计.燃料", "本日.4高炉.燃料+本日.5高炉.燃料+本日.6高炉.燃料+本日.7高炉.燃料+本日.8高炉.燃料", null, null, false);
            AddNumericTag("本日.4_8高炉合计.燃料比", "本日.4_8高炉合计.燃料*1000/本日.4_8高炉合计.全铁产量", 0, null, false);
            AddNumericTag("本日.燃料", "本日.4高炉.燃料+本日.5高炉.燃料+本日.6高炉.燃料+本日.7高炉.燃料+本日.8高炉.燃料+本日.1高炉.燃料", null, null, false);
            AddNumericTag("本日.燃料比", "本日.燃料*1000/本日.全铁产量", 0, null, false);

            AddNumericTag("累计.1高炉.燃料", "累计.1高炉.焦炭+累计.1高炉.焦粒+累计.1高炉.煤粉+累计.1高炉.兰炭", null, null, false);
            AddNumericTag("累计.1高炉.燃料比", "累计.1高炉.燃料*1000/累计.1高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.4高炉.燃料", "累计.4高炉.焦炭+累计.4高炉.焦粒+累计.4高炉.煤粉+累计.4高炉.兰炭", null, null, false);
            AddNumericTag("累计.4高炉.燃料比", "累计.4高炉.燃料*1000/累计.4高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.5高炉.燃料", "累计.5高炉.焦炭+累计.5高炉.焦粒+累计.5高炉.煤粉+累计.5高炉.兰炭", null, null, false);
            AddNumericTag("累计.5高炉.燃料比", "累计.5高炉.燃料*1000/累计.5高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.6高炉.燃料", "累计.6高炉.焦炭+累计.6高炉.焦粒+累计.6高炉.煤粉+累计.6高炉.兰炭", null, null, false);
            AddNumericTag("累计.6高炉.燃料比", "累计.6高炉.燃料*1000/累计.6高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.7高炉.燃料", "累计.7高炉.焦炭+累计.7高炉.焦粒+累计.7高炉.煤粉+累计.7高炉.兰炭", null, null, false);
            AddNumericTag("累计.7高炉.燃料比", "累计.7高炉.燃料*1000/累计.7高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.8高炉.燃料", "累计.8高炉.焦炭+累计.8高炉.焦粒+累计.8高炉.煤粉+累计.8高炉.兰炭", null, null, false);
            AddNumericTag("累计.8高炉.燃料比", "累计.8高炉.燃料*1000/累计.8高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.456高炉合计.燃料", "累计.4高炉.燃料+累计.5高炉.燃料+累计.6高炉.燃料", null, null, false);
            AddNumericTag("累计.456高炉合计.燃料比", "累计.456高炉合计.燃料*1000/累计.456高炉合计.全铁产量", 0, null, false);
            AddNumericTag("累计.78高炉合计.燃料", "累计.7高炉.燃料+累计.8高炉.燃料", null, null, false);
            AddNumericTag("累计.78高炉合计.燃料比", "累计.78高炉合计.燃料*1000/累计.78高炉合计.全铁产量", 0, null, false);
            AddNumericTag("累计.4_8高炉合计.燃料", "累计.4高炉.燃料+累计.5高炉.燃料+累计.6高炉.燃料+累计.7高炉.燃料+累计.8高炉.燃料", null, null, false);
            AddNumericTag("累计.4_8高炉合计.燃料比", "累计.4_8高炉合计.燃料*1000/累计.4_8高炉合计.全铁产量", 0, null, false);
            AddNumericTag("累计.燃料", "累计.4高炉.燃料+累计.5高炉.燃料+累计.6高炉.燃料+累计.7高炉.燃料+累计.8高炉.燃料+累计.1高炉.燃料", null, null, false);
            AddNumericTag("累计.燃料比", "累计.燃料*1000/累计.全铁产量", 0, null, false);

            #endregion

            //入炉焦比
            #region 入炉焦炭、入炉焦炭单耗(入炉焦比)
            int? dec_入炉焦炭单耗 = 0;
            AddNumericTag("本日.1高炉.入炉焦炭", "本日.1高炉.焦炭*(100-本日.1高炉.焦炭水份)/100", null, null, true);
            AddNumericTag("本日.1高炉.入炉焦炭单耗", "本日.1高炉.入炉焦炭*1000/本日.1高炉.全铁产量", dec_入炉焦炭单耗, null, false);
            AddNumericTag("本日.4高炉.入炉焦炭", "本日.4高炉.焦炭*(100-本日.4高炉.焦炭水份)/100", null, null, true);
            AddNumericTag("本日.4高炉.入炉焦炭单耗", "本日.4高炉.入炉焦炭*1000/本日.4高炉.全铁产量", dec_入炉焦炭单耗, null, false);
            AddNumericTag("本日.5高炉.入炉焦炭", "本日.5高炉.焦炭*(100-本日.5高炉.焦炭水份)/100", null, null, true);
            AddNumericTag("本日.5高炉.入炉焦炭单耗", "本日.5高炉.入炉焦炭*1000/本日.5高炉.全铁产量", dec_入炉焦炭单耗, null, false);
            AddNumericTag("本日.6高炉.入炉焦炭", "本日.6高炉.焦炭*(100-本日.6高炉.焦炭水份)/100", null, null, true);
            AddNumericTag("本日.6高炉.入炉焦炭单耗", "本日.6高炉.入炉焦炭*1000/本日.6高炉.全铁产量", dec_入炉焦炭单耗, null, false);
            AddNumericTag("本日.7高炉.入炉焦炭", "本日.7高炉.焦炭*(100-本日.7高炉.焦炭水份)/100", null, null, true);
            AddNumericTag("本日.7高炉.入炉焦炭单耗", "本日.7高炉.入炉焦炭*1000/本日.7高炉.全铁产量", dec_入炉焦炭单耗, null, false);
            AddNumericTag("本日.8高炉.入炉焦炭", "本日.8高炉.焦炭*(100-本日.8高炉.焦炭水份)/100", null, null, true);
            AddNumericTag("本日.8高炉.入炉焦炭单耗", "本日.8高炉.入炉焦炭*1000/本日.8高炉.全铁产量", dec_入炉焦炭单耗, null, false);
            AddNumericTag("本日.456高炉合计.入炉焦炭", "本日.4高炉.入炉焦炭+本日.5高炉.入炉焦炭+本日.6高炉.入炉焦炭", null, null, false);
            AddNumericTag("本日.456高炉合计.入炉焦炭单耗", "本日.456高炉合计.入炉焦炭*1000/本日.456高炉合计.全铁产量", dec_入炉焦炭单耗, null, false);
            AddNumericTag("本日.78高炉合计.入炉焦炭", "本日.7高炉.入炉焦炭+本日.8高炉.入炉焦炭", null, null, false);
            AddNumericTag("本日.78高炉合计.入炉焦炭单耗", "本日.78高炉合计.入炉焦炭*1000/本日.78高炉合计.全铁产量", dec_入炉焦炭单耗, null, false);
            AddNumericTag("本日.4_8高炉合计.入炉焦炭", "本日.4高炉.入炉焦炭+本日.5高炉.入炉焦炭+本日.6高炉.入炉焦炭+本日.7高炉.入炉焦炭+本日.8高炉.入炉焦炭", null, null, false);
            AddNumericTag("本日.4_8高炉合计.入炉焦炭单耗", "本日.4_8高炉合计.入炉焦炭*1000/本日.4_8高炉合计.全铁产量", dec_入炉焦炭单耗, null, false);
            AddNumericTag("本日.入炉焦炭", "本日.4高炉.入炉焦炭+本日.5高炉.入炉焦炭+本日.6高炉.入炉焦炭+本日.7高炉.入炉焦炭+本日.8高炉.入炉焦炭+本日.1高炉.入炉焦炭", null, null, false);
            AddNumericTag("本日.入炉焦炭单耗", "本日.入炉焦炭*1000/本日.全铁产量", dec_入炉焦炭单耗, null, false);

            AddNumericTag("累计.1高炉.入炉焦炭", "昨日累计.1高炉.焦炭+本日.1高炉.入炉焦炭", null, null, true);
            AddNumericTag("累计.1高炉.入炉焦炭单耗", "累计.1高炉.入炉焦炭*1000/累计.1高炉.全铁产量", dec_入炉焦炭单耗, null, false);
            AddNumericTag("累计.4高炉.入炉焦炭", "昨日累计.4高炉.焦炭+本日.4高炉.入炉焦炭", null, null, true);
            AddNumericTag("累计.4高炉.入炉焦炭单耗", "累计.4高炉.入炉焦炭*1000/累计.4高炉.全铁产量", dec_入炉焦炭单耗, null, false);
            AddNumericTag("累计.5高炉.入炉焦炭", "昨日累计.5高炉.焦炭+本日.5高炉.入炉焦炭", null, null, true);
            AddNumericTag("累计.5高炉.入炉焦炭单耗", "累计.5高炉.入炉焦炭*1000/累计.5高炉.全铁产量", dec_入炉焦炭单耗, null, false);
            AddNumericTag("累计.6高炉.入炉焦炭", "昨日累计.6高炉.焦炭+本日.6高炉.入炉焦炭", null, null, true);
            AddNumericTag("累计.6高炉.入炉焦炭单耗", "累计.6高炉.入炉焦炭*1000/累计.6高炉.全铁产量", dec_入炉焦炭单耗, null, false);
            AddNumericTag("累计.7高炉.入炉焦炭", "昨日累计.7高炉.焦炭+本日.7高炉.入炉焦炭", null, null, true);
            AddNumericTag("累计.7高炉.入炉焦炭单耗", "累计.7高炉.入炉焦炭*1000/累计.7高炉.全铁产量", dec_入炉焦炭单耗, null, false);
            AddNumericTag("累计.8高炉.入炉焦炭", "昨日累计.8高炉.焦炭+本日.8高炉.入炉焦炭", null, null, true);
            AddNumericTag("累计.8高炉.入炉焦炭单耗", "累计.8高炉.入炉焦炭*1000/累计.8高炉.全铁产量", dec_入炉焦炭单耗, null, false);
            AddNumericTag("累计.456高炉合计.入炉焦炭", "累计.4高炉.入炉焦炭+累计.5高炉.入炉焦炭+累计.6高炉.入炉焦炭", null, null, false);
            AddNumericTag("累计.456高炉合计.入炉焦炭单耗", "累计.456高炉合计.入炉焦炭*1000/累计.456高炉合计.全铁产量", dec_入炉焦炭单耗, null, false);
            AddNumericTag("累计.78高炉合计.入炉焦炭", "累计.7高炉.入炉焦炭+累计.8高炉.入炉焦炭", null, null, false);
            AddNumericTag("累计.78高炉合计.入炉焦炭单耗", "累计.78高炉合计.入炉焦炭*1000/累计.78高炉合计.全铁产量", dec_入炉焦炭单耗, null, false);
            AddNumericTag("累计.4_8高炉合计.入炉焦炭", "累计.4高炉.入炉焦炭+累计.5高炉.入炉焦炭+累计.6高炉.入炉焦炭+累计.7高炉.入炉焦炭+累计.8高炉.入炉焦炭", null, null, false);
            AddNumericTag("累计.4_8高炉合计.入炉焦炭单耗", "累计.4_8高炉合计.入炉焦炭*1000/累计.4_8高炉合计.全铁产量", dec_入炉焦炭单耗, null, false);
            AddNumericTag("累计.入炉焦炭", "累计.4高炉.入炉焦炭+累计.5高炉.入炉焦炭+累计.6高炉.入炉焦炭+累计.7高炉.入炉焦炭+累计.8高炉.入炉焦炭+累计.1高炉.入炉焦炭", null, null, false);
            AddNumericTag("累计.入炉焦炭单耗", "累计.入炉焦炭*1000/累计.全铁产量", dec_入炉焦炭单耗, null, false);

            AddNumericTag("昨日累计.1高炉.入炉焦炭", "", null, null, false);
            AddNumericTag("昨日累计.4高炉.入炉焦炭", "", null, null, false);
            AddNumericTag("昨日累计.5高炉.入炉焦炭", "", null, null, false);
            AddNumericTag("昨日累计.6高炉.入炉焦炭", "", null, null, false);
            AddNumericTag("昨日累计.7高炉.入炉焦炭", "", null, null, false);
            AddNumericTag("昨日累计.8高炉.入炉焦炭", "", null, null, false);
            #endregion

            #region 综合焦炭、综合焦炭单耗
            int? dec_综合焦比 = 0;
            AddNumericTag("本日.1高炉.综合焦炭", "本日.1高炉.焦炭+0.9*本日.1高炉.焦粉+0.8*本日.1高炉.煤粉", null, null, true);
            AddNumericTag("本日.1高炉.综合焦炭单耗", "本日.1高炉.综合焦炭*1000/本日.1高炉.全铁产量", dec_综合焦比, null, false);
            AddNumericTag("本日.4高炉.综合焦炭", "本日.4高炉.焦炭+0.9*本日.4高炉.焦粉+0.8*本日.4高炉.煤粉", null, null, true);
            AddNumericTag("本日.4高炉.综合焦炭单耗", "本日.4高炉.综合焦炭*1000/本日.4高炉.全铁产量", dec_综合焦比, null, false);
            AddNumericTag("本日.5高炉.综合焦炭", "本日.5高炉.焦炭+0.9*本日.5高炉.焦粉+0.8*本日.5高炉.煤粉", null, null, true);
            AddNumericTag("本日.5高炉.综合焦炭单耗", "本日.5高炉.综合焦炭*1000/本日.5高炉.全铁产量", dec_综合焦比, null, false);
            AddNumericTag("本日.6高炉.综合焦炭", "本日.6高炉.焦炭+0.9*本日.6高炉.焦粉+0.8*本日.6高炉.煤粉", null, null, true);
            AddNumericTag("本日.6高炉.综合焦炭单耗", "本日.6高炉.综合焦炭*1000/本日.6高炉.全铁产量", dec_综合焦比, null, false);
            AddNumericTag("本日.7高炉.综合焦炭", "本日.7高炉.焦炭+0.9*本日.7高炉.焦粉+0.8*本日.7高炉.煤粉", null, null, true);
            AddNumericTag("本日.7高炉.综合焦炭单耗", "本日.7高炉.综合焦炭*1000/本日.7高炉.全铁产量", dec_综合焦比, null, false);
            AddNumericTag("本日.8高炉.综合焦炭", "本日.8高炉.焦炭+0.9*本日.8高炉.焦粉+0.8*本日.8高炉.煤粉", null, null, true);
            AddNumericTag("本日.8高炉.综合焦炭单耗", "本日.8高炉.综合焦炭*1000/本日.8高炉.全铁产量", dec_综合焦比, null, false);
            AddNumericTag("本日.456高炉合计.综合焦炭", "本日.4高炉.综合焦炭+本日.5高炉.综合焦炭+本日.6高炉.综合焦炭", null, null, false);
            AddNumericTag("本日.456高炉合计.综合焦炭单耗", "本日.456高炉合计.综合焦炭*1000/本日.456高炉合计.全铁产量", dec_综合焦比, null, false);
            AddNumericTag("本日.78高炉合计.综合焦炭", "本日.7高炉.综合焦炭+本日.8高炉.综合焦炭", null, null, false);
            AddNumericTag("本日.78高炉合计.综合焦炭单耗", "本日.78高炉合计.综合焦炭*1000/本日.78高炉合计.全铁产量", dec_综合焦比, null, false);
            AddNumericTag("本日.4_8高炉合计.综合焦炭", "本日.4高炉.综合焦炭+本日.5高炉.综合焦炭+本日.6高炉.综合焦炭+本日.7高炉.综合焦炭+本日.8高炉.综合焦炭", null, null, false);
            AddNumericTag("本日.4_8高炉合计.综合焦炭单耗", "本日.4_8高炉合计.综合焦炭*1000/本日.4_8高炉合计.全铁产量", dec_综合焦比, null, false);
            AddNumericTag("本日.综合焦炭", "本日.4高炉.综合焦炭+本日.5高炉.综合焦炭+本日.6高炉.综合焦炭+本日.7高炉.综合焦炭+本日.8高炉.综合焦炭+本日.1高炉.综合焦炭", null, null, false);
            AddNumericTag("本日.综合焦炭单耗", "本日.综合焦炭*1000/本日.全铁产量", dec_综合焦比, null, false);

            AddNumericTag("累计.1高炉.综合焦炭", "昨日累计.1高炉.焦炭+本日.1高炉.综合焦炭", null, null, true);
            AddNumericTag("累计.1高炉.综合焦炭单耗", "累计.1高炉.综合焦炭*1000/累计.1高炉.全铁产量", dec_综合焦比, null, false);
            AddNumericTag("累计.4高炉.综合焦炭", "昨日累计.4高炉.焦炭+本日.4高炉.综合焦炭", null, null, true);
            AddNumericTag("累计.4高炉.综合焦炭单耗", "累计.4高炉.综合焦炭*1000/累计.4高炉.全铁产量", dec_综合焦比, null, false);
            AddNumericTag("累计.5高炉.综合焦炭", "昨日累计.5高炉.焦炭+本日.5高炉.综合焦炭", null, null, true);
            AddNumericTag("累计.5高炉.综合焦炭单耗", "累计.5高炉.综合焦炭*1000/累计.5高炉.全铁产量", dec_综合焦比, null, false);
            AddNumericTag("累计.6高炉.综合焦炭", "昨日累计.6高炉.焦炭+本日.6高炉.综合焦炭", null, null, true);
            AddNumericTag("累计.6高炉.综合焦炭单耗", "累计.6高炉.综合焦炭*1000/累计.6高炉.全铁产量", dec_综合焦比, null, false);
            AddNumericTag("累计.7高炉.综合焦炭", "昨日累计.7高炉.焦炭+本日.7高炉.综合焦炭", null, null, true);
            AddNumericTag("累计.7高炉.综合焦炭单耗", "累计.7高炉.综合焦炭*1000/累计.7高炉.全铁产量", dec_综合焦比, null, false);
            AddNumericTag("累计.8高炉.综合焦炭", "昨日累计.8高炉.焦炭+本日.8高炉.综合焦炭", null, null, true);
            AddNumericTag("累计.8高炉.综合焦炭单耗", "累计.8高炉.综合焦炭*1000/累计.8高炉.全铁产量", dec_综合焦比, null, false);
            AddNumericTag("累计.456高炉合计.综合焦炭", "累计.4高炉.综合焦炭+累计.5高炉.综合焦炭+累计.6高炉.综合焦炭", null, null, false);
            AddNumericTag("累计.456高炉合计.综合焦炭单耗", "累计.456高炉合计.综合焦炭*1000/累计.456高炉合计.全铁产量", dec_综合焦比, null, false);
            AddNumericTag("累计.78高炉合计.综合焦炭", "累计.7高炉.综合焦炭+累计.8高炉.综合焦炭", null, null, false);
            AddNumericTag("累计.78高炉合计.综合焦炭单耗", "累计.78高炉合计.综合焦炭*1000/累计.78高炉合计.全铁产量", dec_综合焦比, null, false);
            AddNumericTag("累计.4_8高炉合计.综合焦炭", "累计.4高炉.综合焦炭+累计.5高炉.综合焦炭+累计.6高炉.综合焦炭+累计.7高炉.综合焦炭+累计.8高炉.综合焦炭", null, null, false);
            AddNumericTag("累计.4_8高炉合计.综合焦炭单耗", "累计.4_8高炉合计.综合焦炭*1000/累计.4_8高炉合计.全铁产量", dec_综合焦比, null, false);
            AddNumericTag("累计.综合焦炭", "累计.4高炉.综合焦炭+累计.5高炉.综合焦炭+累计.6高炉.综合焦炭+累计.7高炉.综合焦炭+累计.8高炉.综合焦炭+累计.1高炉.综合焦炭", null, null, false);
            AddNumericTag("累计.综合焦炭单耗", "累计.综合焦炭*1000/累计.全铁产量", dec_综合焦比, null, false);

            AddNumericTag("昨日累计.1高炉.综合焦炭", "", null, null, false);
            AddNumericTag("昨日累计.4高炉.综合焦炭", "", null, null, false);
            AddNumericTag("昨日累计.5高炉.综合焦炭", "", null, null, false);
            AddNumericTag("昨日累计.6高炉.综合焦炭", "", null, null, false);
            AddNumericTag("昨日累计.7高炉.综合焦炭", "", null, null, false);
            AddNumericTag("昨日累计.8高炉.综合焦炭", "", null, null, false);

            #endregion

            #region 焦炭负荷

            AddNumericTag("本日.1高炉.焦炭负荷", "本日.1高炉.原料矿总耗/本日.1高炉.综合焦炭", 2, null, false);
            AddNumericTag("本日.4高炉.焦炭负荷", "本日.4高炉.原料矿总耗/本日.4高炉.综合焦炭", 2, null, false);
            AddNumericTag("本日.5高炉.焦炭负荷", "本日.5高炉.原料矿总耗/本日.5高炉.综合焦炭", 2, null, false);
            AddNumericTag("本日.6高炉.焦炭负荷", "本日.6高炉.原料矿总耗/本日.6高炉.综合焦炭", 2, null, false);
            AddNumericTag("本日.7高炉.焦炭负荷", "本日.7高炉.原料矿总耗/本日.7高炉.综合焦炭", 2, null, false);
            AddNumericTag("本日.8高炉.焦炭负荷", "本日.8高炉.原料矿总耗/本日.8高炉.综合焦炭", 2, null, false);
            AddNumericTag("本日.456高炉合计.焦炭负荷", "本日.456高炉合计.原料矿总耗/本日.456高炉合计.综合焦炭", 2, null, false);
            AddNumericTag("本日.78高炉合计.焦炭负荷", "本日.78高炉合计.原料矿总耗/本日.78高炉合计.综合焦炭", 2, null, false);
            AddNumericTag("本日.4_8高炉合计.焦炭负荷", "本日.4_8高炉合计.原料矿总耗/本日.4_8高炉合计.综合焦炭", 2, null, false);
            AddNumericTag("本日.焦炭负荷", "本日.原料矿总耗/本日.综合焦炭", 2, null, false);

            AddNumericTag("累计.1高炉.焦炭负荷", "累计.1高炉.原料矿总耗/累计.1高炉.综合焦炭", 2, null, false);
            AddNumericTag("累计.4高炉.焦炭负荷", "累计.4高炉.原料矿总耗/累计.4高炉.综合焦炭", 2, null, false);
            AddNumericTag("累计.5高炉.焦炭负荷", "累计.5高炉.原料矿总耗/累计.5高炉.综合焦炭", 2, null, false);
            AddNumericTag("累计.6高炉.焦炭负荷", "累计.6高炉.原料矿总耗/累计.6高炉.综合焦炭", 2, null, false);
            AddNumericTag("累计.7高炉.焦炭负荷", "累计.7高炉.原料矿总耗/累计.7高炉.综合焦炭", 2, null, false);
            AddNumericTag("累计.8高炉.焦炭负荷", "累计.8高炉.原料矿总耗/累计.8高炉.综合焦炭", 2, null, false);
            AddNumericTag("累计.456高炉合计.焦炭负荷", "累计.456高炉合计.原料矿总耗/累计.456高炉合计.综合焦炭", 2, null, false);
            AddNumericTag("累计.78高炉合计.焦炭负荷", "累计.78高炉合计.原料矿总耗/累计.78高炉合计.综合焦炭", 2, null, false);
            AddNumericTag("累计.4_8高炉合计.焦炭负荷", "累计.4_8高炉合计.原料矿总耗/累计.4_8高炉合计.综合焦炭", 2, null, false);
            AddNumericTag("累计.焦炭负荷", "累计.原料矿总耗/累计.综合焦炭", 2, null, false);


            #endregion

            #region 大中修、高炉有效炉容、冶炼强度

            AddNumericTag("本日.1高炉.大中修", "", null, null, true);
            //((1440*累计天数-累计.1高炉.大中修)/(1440*累计天数))*580
            AddNumericTag("本日.1高炉.高炉有效炉容", "((1440-本日.1高炉.大中修)/1440)*580", null, null, true);
            AddNumericTag("本日.1高炉.冶炼强度", "本日.1高炉.入炉焦炭/本日.1高炉.高炉有效炉容", 3, null, false);
            AddNumericTag("本日.4高炉.大中修", "", null, null, true);
            AddNumericTag("本日.4高炉.高炉有效炉容", "((1440-本日.4高炉.大中修)/1440)*380", null, null, true);
            AddNumericTag("本日.4高炉.冶炼强度", "本日.4高炉.入炉焦炭/本日.4高炉.高炉有效炉容", 3, null, false);
            AddNumericTag("本日.5高炉.大中修", "", null, null, true);
            AddNumericTag("本日.5高炉.高炉有效炉容", "((1440-本日.5高炉.大中修)/1440)*380", null, null, true);
            AddNumericTag("本日.5高炉.冶炼强度", "本日.5高炉.入炉焦炭/本日.5高炉.高炉有效炉容", 3, null, false);
            AddNumericTag("本日.6高炉.大中修", "", null, null, true);
            AddNumericTag("本日.6高炉.高炉有效炉容", "((1440-本日.6高炉.大中修)/1440)*380", null, null, true);
            AddNumericTag("本日.6高炉.冶炼强度", "本日.6高炉.入炉焦炭/本日.6高炉.高炉有效炉容", 3, null, false);
            AddNumericTag("本日.7高炉.大中修", "", null, null, true);
            AddNumericTag("本日.7高炉.高炉有效炉容", "((1440-本日.7高炉.大中修)/1440)*580", null, null, true);
            AddNumericTag("本日.7高炉.冶炼强度", "本日.7高炉.入炉焦炭/本日.7高炉.高炉有效炉容", 3, null, false);
            AddNumericTag("本日.8高炉.大中修", "", null, null, true);
            AddNumericTag("本日.8高炉.高炉有效炉容", "((1440-本日.8高炉.大中修)/1440)*580", null, null, true);
            AddNumericTag("本日.8高炉.冶炼强度", "本日.8高炉.入炉焦炭/本日.8高炉.高炉有效炉容", 3, null, false);
            //AddTag("本日.456高炉合计.大中修", "本日.4高炉.大中修+本日.5高炉.大中修+本日.6高炉.大中修", null, null, true);
            AddNumericTag("本日.456高炉合计.高炉有效炉容", "本日.4高炉.高炉有效炉容+本日.5高炉.高炉有效炉容+本日.6高炉.高炉有效炉容", null, null, false);
            AddNumericTag("本日.456高炉合计.冶炼强度", "本日.456高炉合计.入炉焦炭/本日.456高炉合计.高炉有效炉容", 3, null, false);
            //AddTag("本日.78高炉合计.大中修", "本日.7高炉.大中修+本日.8高炉.大中修", null, null, true);
            AddNumericTag("本日.78高炉合计.高炉有效炉容", "本日.7高炉.高炉有效炉容+本日.8高炉.高炉有效炉容", null, null, false);
            AddNumericTag("本日.78高炉合计.冶炼强度", "本日.78高炉合计.入炉焦炭/本日.78高炉合计.高炉有效炉容", 3, null, false);
            //AddTag("本日.4_8高炉合计.大中修", "本日.4高炉.大中修+本日.5高炉.大中修+本日.6高炉.大中修+本日.7高炉.大中修+本日.8高炉.大中修", null, null, true);
            AddNumericTag("本日.4_8高炉合计.高炉有效炉容", "本日.4高炉.高炉有效炉容+本日.5高炉.高炉有效炉容+本日.6高炉.高炉有效炉容+本日.7高炉.高炉有效炉容+本日.8高炉.高炉有效炉容", null, null, false);
            AddNumericTag("本日.4_8高炉合计.冶炼强度", "本日.4_8高炉合计.入炉焦炭/本日.4_8高炉合计.高炉有效炉容", 3, null, false);
            //AddTag("本日.大中修", "本日.4高炉.大中修+本日.5高炉.大中修+本日.6高炉.大中修", null, null, true);
            AddNumericTag("本日.高炉有效炉容", "本日.4高炉.高炉有效炉容+本日.5高炉.高炉有效炉容+本日.6高炉.高炉有效炉容+本日.7高炉.高炉有效炉容+本日.8高炉.高炉有效炉容+本日.1高炉.高炉有效炉容", null, null, false);
            AddNumericTag("本日.冶炼强度", "本日.入炉焦炭/本日.高炉有效炉容", 3, null, false);

            AddNumericTag("累计.1高炉.大中修", "昨日累计.1高炉.大中修+本日.1高炉.大中修", null, null, true);
            AddNumericTag("累计.1高炉.高炉有效炉容", "昨日累计.1高炉.高炉有效炉容+本日.1高炉.高炉有效炉容", null, null, true);
            AddNumericTag("累计.1高炉.冶炼强度", "累计.1高炉.入炉焦炭/累计.1高炉.高炉有效炉容", 3, null, false);
            AddNumericTag("累计.4高炉.大中修", "昨日累计.4高炉.大中修+本日.4高炉.大中修", null, null, true);
            AddNumericTag("累计.4高炉.高炉有效炉容", "昨日累计.4高炉.高炉有效炉容+本日.4高炉.高炉有效炉容", null, null, true);
            AddNumericTag("累计.4高炉.冶炼强度", "累计.4高炉.入炉焦炭/累计.4高炉.高炉有效炉容", 3, null, false);
            AddNumericTag("累计.5高炉.大中修", "昨日累计.5高炉.大中修+本日.5高炉.大中修", null, null, true);
            AddNumericTag("累计.5高炉.高炉有效炉容", "昨日累计.5高炉.高炉有效炉容+本日.5高炉.高炉有效炉容", null, null, true);
            AddNumericTag("累计.5高炉.冶炼强度", "累计.5高炉.入炉焦炭/累计.5高炉.高炉有效炉容", 3, null, false);
            AddNumericTag("累计.6高炉.大中修", "昨日累计.6高炉.大中修+本日.6高炉.大中修", null, null, true);
            AddNumericTag("累计.6高炉.高炉有效炉容", "昨日累计.6高炉.高炉有效炉容+本日.6高炉.高炉有效炉容", null, null, true);
            AddNumericTag("累计.6高炉.冶炼强度", "累计.6高炉.入炉焦炭/累计.6高炉.高炉有效炉容", 3, null, false);
            AddNumericTag("累计.7高炉.大中修", "昨日累计.7高炉.大中修+本日.7高炉.大中修", null, null, true);
            AddNumericTag("累计.7高炉.高炉有效炉容", "昨日累计.7高炉.高炉有效炉容+本日.7高炉.高炉有效炉容", null, null, true);
            AddNumericTag("累计.7高炉.冶炼强度", "累计.7高炉.入炉焦炭/累计.7高炉.高炉有效炉容", 3, null, false);
            AddNumericTag("累计.8高炉.大中修", "昨日累计.8高炉.大中修+本日.8高炉.大中修", null, null, true);
            AddNumericTag("累计.8高炉.高炉有效炉容", "昨日累计.8高炉.高炉有效炉容+本日.8高炉.高炉有效炉容", null, null, true);
            AddNumericTag("累计.8高炉.冶炼强度", "累计.8高炉.入炉焦炭/累计.8高炉.高炉有效炉容", 3, null, false);
            //AddTag("累计.456高炉合计.大中修", "累计.4高炉.大中修+累计.5高炉.大中修+累计.6高炉.大中修", null, null, true);
            AddNumericTag("累计.456高炉合计.高炉有效炉容", "累计.4高炉.高炉有效炉容+累计.5高炉.高炉有效炉容+累计.6高炉.高炉有效炉容", null, null, false);
            AddNumericTag("累计.456高炉合计.冶炼强度", "累计.456高炉合计.入炉焦炭/累计.456高炉合计.高炉有效炉容", 3, null, false);
            //AddTag("累计.78高炉合计.大中修", "累计.7高炉.大中修+累计.8高炉.大中修", null, null, true);
            AddNumericTag("累计.78高炉合计.高炉有效炉容", "累计.7高炉.高炉有效炉容+累计.8高炉.高炉有效炉容", null, null, false);
            AddNumericTag("累计.78高炉合计.冶炼强度", "累计.78高炉合计.入炉焦炭/累计.78高炉合计.高炉有效炉容", 3, null, false);
            //AddTag("累计.4_8高炉合计.大中修", "累计.4高炉.大中修+累计.5高炉.大中修+累计.6高炉.大中修+累计.7高炉.大中修+累计.8高炉.大中修", null, null, true);
            AddNumericTag("累计.4_8高炉合计.高炉有效炉容", "累计.4高炉.高炉有效炉容+累计.5高炉.高炉有效炉容+累计.6高炉.高炉有效炉容+累计.7高炉.高炉有效炉容+累计.8高炉.高炉有效炉容", null, null, false);
            AddNumericTag("累计.4_8高炉合计.冶炼强度", "累计.4_8高炉合计.入炉焦炭/累计.4_8高炉合计.高炉有效炉容", 3, null, false);
            //AddTag("累计.大中修", "累计.4高炉.大中修+累计.5高炉.大中修+累计.6高炉.大中修", null, null, true);
            AddNumericTag("累计.高炉有效炉容", "累计.4高炉.高炉有效炉容+累计.5高炉.高炉有效炉容+累计.6高炉.高炉有效炉容+累计.7高炉.高炉有效炉容+累计.8高炉.高炉有效炉容+累计.1高炉.高炉有效炉容", null, null, false);
            AddNumericTag("累计.冶炼强度", "累计.入炉焦炭/累计.高炉有效炉容", 3, null, false);

            AddNumericTag("昨日累计.1高炉.大中修", "", null, null, false);
            AddNumericTag("昨日累计.1高炉.高炉有效炉容", "", null, null, false);
            AddNumericTag("昨日累计.4高炉.大中修", "", null, null, false);
            AddNumericTag("昨日累计.4高炉.高炉有效炉容", "", null, null, false);
            AddNumericTag("昨日累计.5高炉.大中修", "", null, null, false);
            AddNumericTag("昨日累计.5高炉.高炉有效炉容", "", null, null, false);
            AddNumericTag("昨日累计.6高炉.大中修", "", null, null, false);
            AddNumericTag("昨日累计.6高炉.高炉有效炉容", "", null, null, false);
            AddNumericTag("昨日累计.7高炉.大中修", "", null, null, false);
            AddNumericTag("昨日累计.7高炉.高炉有效炉容", "", null, null, false);
            AddNumericTag("昨日累计.8高炉.大中修", "", null, null, false);
            AddNumericTag("昨日累计.8高炉.高炉有效炉容", "", null, null, false);

            #endregion

            #region 利用系数

            AddNumericTag("本日.1高炉.利用系数", "本日.1高炉.全铁产量/本日.1高炉.高炉有效炉容", 3, null, false);
            AddNumericTag("本日.4高炉.利用系数", "本日.4高炉.全铁产量/本日.4高炉.高炉有效炉容", 3, null, false);
            AddNumericTag("本日.5高炉.利用系数", "本日.5高炉.全铁产量/本日.5高炉.高炉有效炉容", 3, null, false);
            AddNumericTag("本日.6高炉.利用系数", "本日.6高炉.全铁产量/本日.6高炉.高炉有效炉容", 3, null, false);
            AddNumericTag("本日.7高炉.利用系数", "本日.7高炉.全铁产量/本日.7高炉.高炉有效炉容", 3, null, false);
            AddNumericTag("本日.8高炉.利用系数", "本日.8高炉.全铁产量/本日.8高炉.高炉有效炉容", 3, null, false);

            AddNumericTag("本日.456高炉合计.利用系数", "本日.456高炉合计.全铁产量/本日.456高炉合计.高炉有效炉容", 3, null, false);
            AddNumericTag("本日.78高炉合计.利用系数", "本日.78高炉合计.全铁产量/本日.78高炉合计.高炉有效炉容", 3, null, false);
            AddNumericTag("本日.4_8高炉合计.利用系数", "本日.4_8高炉合计.全铁产量/本日.4_8高炉合计.高炉有效炉容", 3, null, false);
            AddNumericTag("本日.利用系数", "本日.全铁产量/本日.高炉有效炉容", 3, null, false);


            AddNumericTag("累计.1高炉.利用系数", "累计.1高炉.全铁产量/累计.1高炉.高炉有效炉容", 3, null, false);
            AddNumericTag("累计.4高炉.利用系数", "累计.4高炉.全铁产量/累计.4高炉.高炉有效炉容", 3, null, false);
            AddNumericTag("累计.5高炉.利用系数", "累计.5高炉.全铁产量/累计.5高炉.高炉有效炉容", 3, null, false);
            AddNumericTag("累计.6高炉.利用系数", "累计.6高炉.全铁产量/累计.6高炉.高炉有效炉容", 3, null, false);
            AddNumericTag("累计.7高炉.利用系数", "累计.7高炉.全铁产量/累计.7高炉.高炉有效炉容", 3, null, false);
            AddNumericTag("累计.8高炉.利用系数", "累计.8高炉.全铁产量/累计.8高炉.高炉有效炉容", 3, null, false);

            AddNumericTag("累计.456高炉合计.利用系数", "累计.456高炉合计.全铁产量/累计.456高炉合计.高炉有效炉容", 3, null, false);
            AddNumericTag("累计.78高炉合计.利用系数", "累计.78高炉合计.全铁产量/累计.78高炉合计.高炉有效炉容", 3, null, false);
            AddNumericTag("累计.4_8高炉合计.利用系数", "累计.4_8高炉合计.全铁产量/累计.4_8高炉合计.高炉有效炉容", 3, null, false);
            AddNumericTag("累计.利用系数", "累计.全铁产量/累计.高炉有效炉容", 3, null, false);


            #endregion

            #region 氮气
            AddNumericTag("本日.1高炉.氮气", "", 2, null, true);
            AddNumericTag("本日.4高炉.氮气", "", 2, null, true);
            AddNumericTag("本日.5高炉.氮气", "", 2, null, true);
            AddNumericTag("本日.6高炉.氮气", "", 2, null, true);
            AddNumericTag("本日.7高炉.氮气", "", 2, null, true);
            AddNumericTag("本日.8高炉.氮气", "", 2, null, true);

            AddNumericTag("本日.456高炉合计.氮气", "本日.4高炉.氮气+本日.5高炉.氮气+本日.6高炉.氮气", 2, null, false);
            AddNumericTag("本日.78高炉合计.氮气", "本日.7高炉.氮气+本日.8高炉.氮气", 2, null, false);
            AddNumericTag("本日.4_8高炉合计.氮气", "本日.4高炉.氮气+本日.5高炉.氮气+本日.6高炉.氮气+本日.7高炉.氮气+本日.8高炉.氮气", 2, null, false);
            AddNumericTag("本日.氮气", "本日.4高炉.氮气+本日.5高炉.氮气+本日.6高炉.氮气+本日.7高炉.氮气+本日.8高炉.氮气+本日.1高炉.氮气", 2, null, false);

            AddNumericTag("累计.1高炉.氮气", "昨日累计.1高炉.氮气+本日.1高炉.氮气", 2, null, true);
            AddNumericTag("累计.4高炉.氮气", "昨日累计.4高炉.氮气+本日.4高炉.氮气", 2, null, true);
            AddNumericTag("累计.5高炉.氮气", "昨日累计.5高炉.氮气+本日.5高炉.氮气", 2, null, true);
            AddNumericTag("累计.6高炉.氮气", "昨日累计.6高炉.氮气+本日.6高炉.氮气", 2, null, true);
            AddNumericTag("累计.7高炉.氮气", "昨日累计.7高炉.氮气+本日.7高炉.氮气", 2, null, true);
            AddNumericTag("累计.8高炉.氮气", "昨日累计.8高炉.氮气+本日.8高炉.氮气", 2, null, true);

            AddNumericTag("累计.456高炉合计.氮气", "累计.4高炉.氮气+累计.5高炉.氮气+累计.6高炉.氮气", 2, null, false);
            AddNumericTag("累计.78高炉合计.氮气", "累计.7高炉.氮气+累计.8高炉.氮气", 2, null, false);
            AddNumericTag("累计.4_8高炉合计.氮气", "累计.4高炉.氮气+累计.5高炉.氮气+累计.6高炉.氮气+累计.7高炉.氮气+累计.8高炉.氮气", 2, null, false);
            AddNumericTag("累计.氮气", "累计.4高炉.氮气+累计.5高炉.氮气+累计.6高炉.氮气+累计.7高炉.氮气+累计.8高炉.氮气+累计.1高炉.氮气", 2, null, false);

            AddNumericTag("昨日累计.1高炉.氮气", "", 2, null, false);
            AddNumericTag("昨日累计.4高炉.氮气", "", 2, null, false);
            AddNumericTag("昨日累计.5高炉.氮气", "", 2, null, false);
            AddNumericTag("昨日累计.6高炉.氮气", "", 2, null, false);
            AddNumericTag("昨日累计.7高炉.氮气", "", 2, null, false);
            AddNumericTag("昨日累计.8高炉.氮气", "", 2, null, false);
            #endregion

            #region 氮比

            AddNumericTag("本日.1高炉.氮比", "本日.1高炉.氮气/本日.1高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.4高炉.氮比", "本日.4高炉.氮气/本日.4高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.5高炉.氮比", "本日.5高炉.氮气/本日.5高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.6高炉.氮比", "本日.6高炉.氮气/本日.6高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.7高炉.氮比", "本日.7高炉.氮气/本日.7高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.8高炉.氮比", "本日.8高炉.氮气/本日.8高炉.全铁产量", 0, null, false);

            AddNumericTag("本日.456高炉合计.氮比", "本日.456高炉合计.氮气/本日.456高炉合计.全铁产量", 0, null, false);
            AddNumericTag("本日.78高炉合计.氮比", "本日.78高炉合计.氮气/本日.78高炉合计.全铁产量", 0, null, false);
            AddNumericTag("本日.4_8高炉合计.氮比", "本日.4_8高炉合计.氮气/本日.4_8高炉合计.全铁产量", 0, null, false);
            AddNumericTag("本日.氮比", "本日.氮气/本日.全铁产量", 0, null, false);


            AddNumericTag("累计.1高炉.氮比", "累计.1高炉.氮气/累计.1高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.4高炉.氮比", "累计.4高炉.氮气/累计.4高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.5高炉.氮比", "累计.5高炉.氮气/累计.5高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.6高炉.氮比", "累计.6高炉.氮气/累计.6高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.7高炉.氮比", "累计.7高炉.氮气/累计.7高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.8高炉.氮比", "累计.8高炉.氮气/累计.8高炉.全铁产量", 0, null, false);

            AddNumericTag("累计.456高炉合计.氮比", "累计.456高炉合计.氮气/累计.456高炉合计.全铁产量", 0, null, false);
            AddNumericTag("累计.78高炉合计.氮比", "累计.78高炉合计.氮气/累计.78高炉合计.全铁产量", 0, null, false);
            AddNumericTag("累计.4_8高炉合计.氮比", "累计.4_8高炉合计.氮气/累计.4_8高炉合计.全铁产量", 0, null, false);
            AddNumericTag("累计.氮比", "累计.氮气/累计.全铁产量", 0, null, false);

            #endregion

            #region 氧气
            AddNumericTag("本日.1高炉.氧气", "", 2, null, true);
            AddNumericTag("本日.4高炉.氧气", "", 2, null, true);
            AddNumericTag("本日.5高炉.氧气", "", 2, null, true);
            AddNumericTag("本日.6高炉.氧气", "", 2, null, true);
            AddNumericTag("本日.7高炉.氧气", "", 2, null, true);
            AddNumericTag("本日.8高炉.氧气", "", 2, null, true);

            AddNumericTag("本日.456高炉合计.氧气", "本日.4高炉.氧气+本日.5高炉.氧气+本日.6高炉.氧气", 2, null, false);
            AddNumericTag("本日.78高炉合计.氧气", "本日.7高炉.氧气+本日.8高炉.氧气", 2, null, false);
            AddNumericTag("本日.4_8高炉合计.氧气", "本日.4高炉.氧气+本日.5高炉.氧气+本日.6高炉.氧气+本日.7高炉.氧气+本日.8高炉.氧气", 2, null, false);
            AddNumericTag("本日.氧气", "本日.4高炉.氧气+本日.5高炉.氧气+本日.6高炉.氧气+本日.7高炉.氧气+本日.8高炉.氧气+本日.1高炉.氧气", 2, null, false);

            AddNumericTag("累计.1高炉.氧气", "昨日累计.1高炉.氧气+本日.1高炉.氧气", 2, null, true);
            AddNumericTag("累计.4高炉.氧气", "昨日累计.4高炉.氧气+本日.4高炉.氧气", 2, null, true);
            AddNumericTag("累计.5高炉.氧气", "昨日累计.5高炉.氧气+本日.5高炉.氧气", 2, null, true);
            AddNumericTag("累计.6高炉.氧气", "昨日累计.6高炉.氧气+本日.6高炉.氧气", 2, null, true);
            AddNumericTag("累计.7高炉.氧气", "昨日累计.7高炉.氧气+本日.7高炉.氧气", 2, null, true);
            AddNumericTag("累计.8高炉.氧气", "昨日累计.8高炉.氧气+本日.8高炉.氧气", 2, null, true);

            AddNumericTag("累计.456高炉合计.氧气", "累计.4高炉.氧气+累计.5高炉.氧气+累计.6高炉.氧气", 2, null, false);
            AddNumericTag("累计.78高炉合计.氧气", "累计.7高炉.氧气+累计.8高炉.氧气", 2, null, false);
            AddNumericTag("累计.4_8高炉合计.氧气", "累计.4高炉.氧气+累计.5高炉.氧气+累计.6高炉.氧气+累计.7高炉.氧气+累计.8高炉.氧气", 2, null, false);
            AddNumericTag("累计.氧气", "累计.4高炉.氧气+累计.5高炉.氧气+累计.6高炉.氧气+累计.7高炉.氧气+累计.8高炉.氧气+累计.1高炉.氧气", 2, null, false);

            AddNumericTag("昨日累计.1高炉.氧气", "", 2, null, false);
            AddNumericTag("昨日累计.4高炉.氧气", "", 2, null, false);
            AddNumericTag("昨日累计.5高炉.氧气", "", 2, null, false);
            AddNumericTag("昨日累计.6高炉.氧气", "", 2, null, false);
            AddNumericTag("昨日累计.7高炉.氧气", "", 2, null, false);
            AddNumericTag("昨日累计.8高炉.氧气", "", 2, null, false);
            #endregion

            #region 氧比

            AddNumericTag("本日.1高炉.氧比", "本日.1高炉.氧气/本日.1高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.4高炉.氧比", "本日.4高炉.氧气/本日.4高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.5高炉.氧比", "本日.5高炉.氧气/本日.5高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.6高炉.氧比", "本日.6高炉.氧气/本日.6高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.7高炉.氧比", "本日.7高炉.氧气/本日.7高炉.全铁产量", 0, null, false);
            AddNumericTag("本日.8高炉.氧比", "本日.8高炉.氧气/本日.8高炉.全铁产量", 0, null, false);

            AddNumericTag("本日.456高炉合计.氧比", "本日.456高炉合计.氧气/本日.456高炉合计.全铁产量", 0, null, false);
            AddNumericTag("本日.78高炉合计.氧比", "本日.78高炉合计.氧气/本日.78高炉合计.全铁产量", 0, null, false);
            AddNumericTag("本日.4_8高炉合计.氧比", "本日.4_8高炉合计.氧气/本日.4_8高炉合计.全铁产量", 0, null, false);
            AddNumericTag("本日.氧比", "本日.氧气/本日.全铁产量", 0, null, false);


            AddNumericTag("累计.1高炉.氧比", "累计.1高炉.氧气/累计.1高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.4高炉.氧比", "累计.4高炉.氧气/累计.4高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.5高炉.氧比", "累计.5高炉.氧气/累计.5高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.6高炉.氧比", "累计.6高炉.氧气/累计.6高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.7高炉.氧比", "累计.7高炉.氧气/累计.7高炉.全铁产量", 0, null, false);
            AddNumericTag("累计.8高炉.氧比", "累计.8高炉.氧气/累计.8高炉.全铁产量", 0, null, false);

            AddNumericTag("累计.456高炉合计.氧比", "累计.456高炉合计.氧气/累计.456高炉合计.全铁产量", 0, null, false);
            AddNumericTag("累计.78高炉合计.氧比", "累计.78高炉合计.氧气/累计.78高炉合计.全铁产量", 0, null, false);
            AddNumericTag("累计.4_8高炉合计.氧比", "累计.4_8高炉合计.氧气/累计.4_8高炉合计.全铁产量", 0, null, false);
            AddNumericTag("累计.氧比", "累计.氧气/累计.全铁产量", 0, null, false);

            #endregion


            #region 耗电
            int? dec_耗电 = 0;
            AddNumericTag("本日.1高炉.耗电", "", dec_耗电, null, true);
            AddNumericTag("本日.4高炉.耗电", "", dec_耗电, null, true);
            AddNumericTag("本日.5高炉.耗电", "", dec_耗电, null, true);
            AddNumericTag("本日.6高炉.耗电", "", dec_耗电, null, true);
            AddNumericTag("本日.7高炉.耗电", "", dec_耗电, null, true);
            AddNumericTag("本日.8高炉.耗电", "", dec_耗电, null, true);

            AddNumericTag("本日.456高炉合计.耗电", "本日.4高炉.耗电+本日.5高炉.耗电+本日.6高炉.耗电", dec_耗电, null, false);
            AddNumericTag("本日.78高炉合计.耗电", "本日.7高炉.耗电+本日.8高炉.耗电", dec_耗电, null, false);
            AddNumericTag("本日.4_8高炉合计.耗电", "本日.4高炉.耗电+本日.5高炉.耗电+本日.6高炉.耗电+本日.7高炉.耗电+本日.8高炉.耗电", dec_耗电, null, false);
            AddNumericTag("本日.耗电", "本日.4高炉.耗电+本日.5高炉.耗电+本日.6高炉.耗电+本日.7高炉.耗电+本日.8高炉.耗电+本日.1高炉.耗电", dec_耗电, null, false);

            AddNumericTag("累计.1高炉.耗电", "昨日累计.1高炉.耗电+本日.1高炉.耗电", dec_耗电, null, true);
            AddNumericTag("累计.4高炉.耗电", "昨日累计.4高炉.耗电+本日.4高炉.耗电", dec_耗电, null, true);
            AddNumericTag("累计.5高炉.耗电", "昨日累计.5高炉.耗电+本日.5高炉.耗电", dec_耗电, null, true);
            AddNumericTag("累计.6高炉.耗电", "昨日累计.6高炉.耗电+本日.6高炉.耗电", dec_耗电, null, true);
            AddNumericTag("累计.7高炉.耗电", "昨日累计.7高炉.耗电+本日.7高炉.耗电", dec_耗电, null, true);
            AddNumericTag("累计.8高炉.耗电", "昨日累计.8高炉.耗电+本日.8高炉.耗电", dec_耗电, null, true);

            AddNumericTag("累计.456高炉合计.耗电", "累计.4高炉.耗电+累计.5高炉.耗电+累计.6高炉.耗电", dec_耗电, null, false);
            AddNumericTag("累计.78高炉合计.耗电", "累计.7高炉.耗电+累计.8高炉.耗电", dec_耗电, null, false);
            AddNumericTag("累计.4_8高炉合计.耗电", "累计.4高炉.耗电+累计.5高炉.耗电+累计.6高炉.耗电+累计.7高炉.耗电+累计.8高炉.耗电", dec_耗电, null, false);
            AddNumericTag("累计.耗电", "累计.4高炉.耗电+累计.5高炉.耗电+累计.6高炉.耗电+累计.7高炉.耗电+累计.8高炉.耗电+累计.1高炉.耗电", dec_耗电, null, false);

            AddNumericTag("昨日累计.1高炉.耗电", "", dec_耗电, null, false);
            AddNumericTag("昨日累计.4高炉.耗电", "", dec_耗电, null, false);
            AddNumericTag("昨日累计.5高炉.耗电", "", dec_耗电, null, false);
            AddNumericTag("昨日累计.6高炉.耗电", "", dec_耗电, null, false);
            AddNumericTag("昨日累计.7高炉.耗电", "", dec_耗电, null, false);
            AddNumericTag("昨日累计.8高炉.耗电", "", dec_耗电, null, false);
            #endregion

            #region 电单耗
            int? dec_电单耗 = 2;
            AddNumericTag("本日.1高炉.电单耗", "本日.1高炉.耗电/本日.1高炉.全铁产量", dec_电单耗, null, false);
            AddNumericTag("本日.4高炉.电单耗", "本日.4高炉.耗电/本日.4高炉.全铁产量", dec_电单耗, null, false);
            AddNumericTag("本日.5高炉.电单耗", "本日.5高炉.耗电/本日.5高炉.全铁产量", dec_电单耗, null, false);
            AddNumericTag("本日.6高炉.电单耗", "本日.6高炉.耗电/本日.6高炉.全铁产量", dec_电单耗, null, false);
            AddNumericTag("本日.7高炉.电单耗", "本日.7高炉.耗电/本日.7高炉.全铁产量", dec_电单耗, null, false);
            AddNumericTag("本日.8高炉.电单耗", "本日.8高炉.耗电/本日.8高炉.全铁产量", dec_电单耗, null, false);

            AddNumericTag("本日.456高炉合计.电单耗", "本日.456高炉合计.耗电/本日.456高炉合计.全铁产量", dec_电单耗, null, false);
            AddNumericTag("本日.78高炉合计.电单耗", "本日.78高炉合计.耗电/本日.78高炉合计.全铁产量", dec_电单耗, null, false);
            AddNumericTag("本日.4_8高炉合计.电单耗", "本日.4_8高炉合计.耗电/本日.4_8高炉合计.全铁产量", dec_电单耗, null, false);
            AddNumericTag("本日.电单耗", "本日.耗电/本日.全铁产量", dec_电单耗, null, false);


            AddNumericTag("累计.1高炉.电单耗", "累计.1高炉.耗电/累计.1高炉.全铁产量", dec_电单耗, null, false);
            AddNumericTag("累计.4高炉.电单耗", "累计.4高炉.耗电/累计.4高炉.全铁产量", dec_电单耗, null, false);
            AddNumericTag("累计.5高炉.电单耗", "累计.5高炉.耗电/累计.5高炉.全铁产量", dec_电单耗, null, false);
            AddNumericTag("累计.6高炉.电单耗", "累计.6高炉.耗电/累计.6高炉.全铁产量", dec_电单耗, null, false);
            AddNumericTag("累计.7高炉.电单耗", "累计.7高炉.耗电/累计.7高炉.全铁产量", dec_电单耗, null, false);
            AddNumericTag("累计.8高炉.电单耗", "累计.8高炉.耗电/累计.8高炉.全铁产量", dec_电单耗, null, false);

            AddNumericTag("累计.456高炉合计.电单耗", "累计.456高炉合计.耗电/累计.456高炉合计.全铁产量", dec_电单耗, null, false);
            AddNumericTag("累计.78高炉合计.电单耗", "累计.78高炉合计.耗电/累计.78高炉合计.全铁产量", dec_电单耗, null, false);
            AddNumericTag("累计.4_8高炉合计.电单耗", "累计.4_8高炉合计.耗电/累计.4_8高炉合计.全铁产量", dec_电单耗, null, false);
            AddNumericTag("累计.电单耗", "累计.耗电/累计.全铁产量", dec_电单耗, null, false);

            #endregion

            #region 冷风
            int? dec_冷风 = 0;
            AddNumericTag("本日.1高炉.冷风", "", dec_冷风, null, true);
            AddNumericTag("本日.4高炉.冷风", "", dec_冷风, null, true);
            AddNumericTag("本日.5高炉.冷风", "", dec_冷风, null, true);
            AddNumericTag("本日.6高炉.冷风", "", dec_冷风, null, true);
            AddNumericTag("本日.7高炉.冷风", "", dec_冷风, null, true);
            AddNumericTag("本日.8高炉.冷风", "", dec_冷风, null, true);

            AddNumericTag("本日.456高炉合计.冷风", "本日.4高炉.冷风+本日.5高炉.冷风+本日.6高炉.冷风", dec_冷风, null, false);
            AddNumericTag("本日.78高炉合计.冷风", "本日.7高炉.冷风+本日.8高炉.冷风", dec_冷风, null, false);
            AddNumericTag("本日.4_8高炉合计.冷风", "本日.4高炉.冷风+本日.5高炉.冷风+本日.6高炉.冷风+本日.7高炉.冷风+本日.8高炉.冷风", dec_冷风, null, false);
            AddNumericTag("本日.冷风", "本日.4高炉.冷风+本日.5高炉.冷风+本日.6高炉.冷风+本日.7高炉.冷风+本日.8高炉.冷风+本日.1高炉.冷风", dec_冷风, null, false);

            AddNumericTag("累计.1高炉.冷风", "昨日累计.1高炉.冷风+本日.1高炉.冷风", dec_冷风, null, true);
            AddNumericTag("累计.4高炉.冷风", "昨日累计.4高炉.冷风+本日.4高炉.冷风", dec_冷风, null, true);
            AddNumericTag("累计.5高炉.冷风", "昨日累计.5高炉.冷风+本日.5高炉.冷风", dec_冷风, null, true);
            AddNumericTag("累计.6高炉.冷风", "昨日累计.6高炉.冷风+本日.6高炉.冷风", dec_冷风, null, true);
            AddNumericTag("累计.7高炉.冷风", "昨日累计.7高炉.冷风+本日.7高炉.冷风", dec_冷风, null, true);
            AddNumericTag("累计.8高炉.冷风", "昨日累计.8高炉.冷风+本日.8高炉.冷风", dec_冷风, null, true);

            AddNumericTag("累计.456高炉合计.冷风", "累计.4高炉.冷风+累计.5高炉.冷风+累计.6高炉.冷风", dec_冷风, null, false);
            AddNumericTag("累计.78高炉合计.冷风", "累计.7高炉.冷风+累计.8高炉.冷风", dec_冷风, null, false);
            AddNumericTag("累计.4_8高炉合计.冷风", "累计.4高炉.冷风+累计.5高炉.冷风+累计.6高炉.冷风+累计.7高炉.冷风+累计.8高炉.冷风", dec_冷风, null, false);
            AddNumericTag("累计.冷风", "累计.4高炉.冷风+累计.5高炉.冷风+累计.6高炉.冷风+累计.7高炉.冷风+累计.8高炉.冷风+累计.1高炉.冷风", dec_冷风, null, false);

            AddNumericTag("昨日累计.1高炉.冷风", "", dec_冷风, null, false);
            AddNumericTag("昨日累计.4高炉.冷风", "", dec_冷风, null, false);
            AddNumericTag("昨日累计.5高炉.冷风", "", dec_冷风, null, false);
            AddNumericTag("昨日累计.6高炉.冷风", "", dec_冷风, null, false);
            AddNumericTag("昨日累计.7高炉.冷风", "", dec_冷风, null, false);
            AddNumericTag("昨日累计.8高炉.冷风", "", dec_冷风, null, false);
            #endregion

            #region 富氧率
            int? dec_富氧率 = 2;
            AddNumericTag("本日.1高炉.富氧率", "0.785*本日.1高炉.氧气*100/60/本日.1高炉.冷风", dec_富氧率, null, false);
            AddNumericTag("本日.4高炉.富氧率", "0.785*本日.4高炉.氧气*100/60/本日.4高炉.冷风", dec_富氧率, null, false);
            AddNumericTag("本日.5高炉.富氧率", "0.785*本日.5高炉.氧气*100/60/本日.5高炉.冷风", dec_富氧率, null, false);
            AddNumericTag("本日.6高炉.富氧率", "0.785*本日.6高炉.氧气*100/60/本日.6高炉.冷风", dec_富氧率, null, false);
            AddNumericTag("本日.7高炉.富氧率", "0.785*本日.7高炉.氧气*100/60/本日.7高炉.冷风", dec_富氧率, null, false);
            AddNumericTag("本日.8高炉.富氧率", "0.785*本日.8高炉.氧气*100/60/本日.8高炉.冷风", dec_富氧率, null, false);
            AddNumericTag("本日.456高炉合计.富氧率", "0.785*本日.456高炉合计.氧气*100/60/本日.456高炉合计.冷风", dec_富氧率, null, false);
            AddNumericTag("本日.78高炉合计.富氧率", "0.785*本日.78高炉合计.氧气*100/60/本日.78高炉合计.冷风", dec_富氧率, null, false);
            AddNumericTag("本日.4_8高炉合计.富氧率", "0.785*本日.4_8高炉合计.氧气*100/60/本日.4_8高炉合计.冷风", dec_富氧率, null, false);
            AddNumericTag("本日.富氧率", "0.785*本日.氧气*100/60/本日.冷风", dec_富氧率, null, false);

            AddNumericTag("累计.1高炉.富氧率", "0.785*累计.1高炉.氧气*100/60/累计.1高炉.冷风", dec_富氧率, null, false);
            AddNumericTag("累计.4高炉.富氧率", "0.785*累计.4高炉.氧气*100/60/累计.4高炉.冷风", dec_富氧率, null, false);
            AddNumericTag("累计.5高炉.富氧率", "0.785*累计.5高炉.氧气*100/60/累计.5高炉.冷风", dec_富氧率, null, false);
            AddNumericTag("累计.6高炉.富氧率", "0.785*累计.6高炉.氧气*100/60/累计.6高炉.冷风", dec_富氧率, null, false);
            AddNumericTag("累计.7高炉.富氧率", "0.785*累计.7高炉.氧气*100/60/累计.7高炉.冷风", dec_富氧率, null, false);
            AddNumericTag("累计.8高炉.富氧率", "0.785*累计.8高炉.氧气*100/60/累计.8高炉.冷风", dec_富氧率, null, false);
            AddNumericTag("累计.456高炉合计.富氧率", "0.785*累计.456高炉合计.氧气*100/60/累计.456高炉合计.冷风", dec_富氧率, null, false);
            AddNumericTag("累计.78高炉合计.富氧率", "0.785*累计.78高炉合计.氧气*100/60/累计.78高炉合计.冷风", dec_富氧率, null, false);
            AddNumericTag("累计.4_8高炉合计.富氧率", "0.785*累计.4_8高炉合计.氧气*100/60/累计.4_8高炉合计.冷风", dec_富氧率, null, false);
            AddNumericTag("累计.富氧率", "0.785*累计.氧气*100/60/累计.冷风", dec_富氧率, null, false);

            #endregion

            #region 休风
            int? dec_休风 = 0;

            AddNumericTag("本日.1高炉.休风", "本日.1高炉.休风工艺+本日.1高炉.休风设备+本日.1高炉.休风电+本日.1高炉.休风外围+本日.1高炉.休风计划", dec_休风, null, false);
            AddNumericTag("本日.1高炉.休风电", "", dec_休风, null, true);
            AddNumericTag("本日.1高炉.休风工艺", "", dec_休风, null, true);
            AddNumericTag("本日.1高炉.休风计划", "", dec_休风, null, true);
            AddNumericTag("本日.1高炉.休风设备", "", dec_休风, null, true);
            AddNumericTag("本日.1高炉.休风外围", "", dec_休风, null, true);

            AddNumericTag("本日.4高炉.休风", "本日.4高炉.休风工艺+本日.4高炉.休风设备+本日.4高炉.休风电+本日.4高炉.休风外围+本日.4高炉.休风计划", dec_休风, null, false);
            AddNumericTag("本日.4高炉.休风电", "", dec_休风, null, true);
            AddNumericTag("本日.4高炉.休风工艺", "", dec_休风, null, true);
            AddNumericTag("本日.4高炉.休风计划", "", dec_休风, null, true);
            AddNumericTag("本日.4高炉.休风设备", "", dec_休风, null, true);
            AddNumericTag("本日.4高炉.休风外围", "", dec_休风, null, true);

            AddNumericTag("本日.5高炉.休风", "本日.5高炉.休风工艺+本日.5高炉.休风设备+本日.5高炉.休风电+本日.5高炉.休风外围+本日.5高炉.休风计划", dec_休风, null, false);
            AddNumericTag("本日.5高炉.休风电", "", dec_休风, null, true);
            AddNumericTag("本日.5高炉.休风工艺", "", dec_休风, null, true);
            AddNumericTag("本日.5高炉.休风计划", "", dec_休风, null, true);
            AddNumericTag("本日.5高炉.休风设备", "", dec_休风, null, true);
            AddNumericTag("本日.5高炉.休风外围", "", dec_休风, null, true);

            AddNumericTag("本日.6高炉.休风", "本日.6高炉.休风工艺+本日.6高炉.休风设备+本日.6高炉.休风电+本日.6高炉.休风外围+本日.6高炉.休风计划", dec_休风, null, false);
            AddNumericTag("本日.6高炉.休风电", "", dec_休风, null, true);
            AddNumericTag("本日.6高炉.休风工艺", "", dec_休风, null, true);
            AddNumericTag("本日.6高炉.休风计划", "", dec_休风, null, true);
            AddNumericTag("本日.6高炉.休风设备", "", dec_休风, null, true);
            AddNumericTag("本日.6高炉.休风外围", "", dec_休风, null, true);

            AddNumericTag("本日.7高炉.休风", "本日.7高炉.休风计划+本日.7高炉.休风外围+本日.7高炉.休风电+本日.7高炉.休风设备+本日.7高炉.休风工艺", dec_休风, null, false);
            AddNumericTag("本日.7高炉.休风电", "", dec_休风, null, true);
            AddNumericTag("本日.7高炉.休风工艺", "", dec_休风, null, true);
            AddNumericTag("本日.7高炉.休风计划", "", dec_休风, null, true);
            AddNumericTag("本日.7高炉.休风设备", "", dec_休风, null, true);
            AddNumericTag("本日.7高炉.休风外围", "", dec_休风, null, true);

            AddNumericTag("本日.8高炉.休风", "本日.8高炉.休风工艺+本日.8高炉.休风设备+本日.8高炉.休风电+本日.8高炉.休风外围+本日.8高炉.休风计划", dec_休风, null, false);
            AddNumericTag("本日.8高炉.休风电", "", dec_休风, null, true);
            AddNumericTag("本日.8高炉.休风工艺", "", dec_休风, null, true);
            AddNumericTag("本日.8高炉.休风计划", "", dec_休风, null, true);
            AddNumericTag("本日.8高炉.休风设备", "", dec_休风, null, true);
            AddNumericTag("本日.8高炉.休风外围", "", dec_休风, null, true);


            AddNumericTag("本日.456高炉合计.休风", "本日.456高炉合计.休风计划+本日.456高炉合计.休风外围+本日.456高炉合计.休风电+本日.456高炉合计.休风设备+本日.456高炉合计.休风工艺", dec_休风, null, false);
            AddNumericTag("本日.456高炉合计.休风电", "本日.4高炉.休风电+本日.5高炉.休风电+本日.6高炉.休风电", dec_休风, null, false);
            AddNumericTag("本日.456高炉合计.休风工艺", "本日.4高炉.休风工艺+本日.5高炉.休风工艺+本日.6高炉.休风工艺", dec_休风, null, false);
            AddNumericTag("本日.456高炉合计.休风计划", "本日.4高炉.休风计划+本日.5高炉.休风计划+本日.6高炉.休风计划", dec_休风, null, false);
            AddNumericTag("本日.456高炉合计.休风设备", "本日.4高炉.休风设备+本日.5高炉.休风设备+本日.6高炉.休风设备", dec_休风, null, false);
            AddNumericTag("本日.456高炉合计.休风外围", "本日.4高炉.休风外围+本日.5高炉.休风外围+本日.6高炉.休风外围", dec_休风, null, false);

            AddNumericTag("本日.78高炉合计.休风", "本日.78高炉合计.休风工艺+本日.78高炉合计.休风设备+本日.78高炉合计.休风电+本日.78高炉合计.休风外围+本日.78高炉合计.休风计划", dec_休风, null, false);
            AddNumericTag("本日.78高炉合计.休风电", "本日.7高炉.休风电+本日.8高炉.休风电", dec_休风, null, false);
            AddNumericTag("本日.78高炉合计.休风工艺", "本日.7高炉.休风工艺+本日.8高炉.休风工艺", dec_休风, null, false);
            AddNumericTag("本日.78高炉合计.休风计划", "本日.7高炉.休风计划+本日.8高炉.休风计划", dec_休风, null, false);
            AddNumericTag("本日.78高炉合计.休风设备", "本日.7高炉.休风设备+本日.8高炉.休风设备", dec_休风, null, false);
            AddNumericTag("本日.78高炉合计.休风外围", "本日.7高炉.休风外围+本日.8高炉.休风外围", dec_休风, null, false);

            AddNumericTag("本日.4_8高炉合计.休风", "本日.4_8高炉合计.休风工艺+本日.4_8高炉合计.休风设备+本日.4_8高炉合计.休风电+本日.4_8高炉合计.休风外围+本日.4_8高炉合计.休风计划", dec_休风, null, false);
            AddNumericTag("本日.4_8高炉合计.休风电", "本日.4高炉.休风电+本日.5高炉.休风电+本日.6高炉.休风电+本日.7高炉.休风电+本日.8高炉.休风电", dec_休风, null, false);
            AddNumericTag("本日.4_8高炉合计.休风工艺", "本日.4高炉.休风工艺+本日.5高炉.休风工艺+本日.6高炉.休风工艺+本日.7高炉.休风工艺+本日.8高炉.休风工艺", dec_休风, null, false);
            AddNumericTag("本日.4_8高炉合计.休风计划", "本日.4高炉.休风计划+本日.5高炉.休风计划+本日.6高炉.休风计划+本日.7高炉.休风计划+本日.8高炉.休风计划", dec_休风, null, false);
            AddNumericTag("本日.4_8高炉合计.休风设备", "本日.4高炉.休风设备+本日.5高炉.休风设备+本日.6高炉.休风设备+本日.7高炉.休风设备+本日.8高炉.休风设备", dec_休风, null, false);
            AddNumericTag("本日.4_8高炉合计.休风外围", "本日.4高炉.休风外围+本日.5高炉.休风外围+本日.6高炉.休风外围+本日.7高炉.休风外围+本日.8高炉.休风外围", dec_休风, null, false);

            AddNumericTag("本日.休风", "本日.休风工艺+本日.休风设备+本日.休风电+本日.休风外围+本日.休风计划", dec_休风, null, false);
            AddNumericTag("本日.休风电", "本日.4高炉.休风电+本日.5高炉.休风电+本日.6高炉.休风电+本日.7高炉.休风电+本日.8高炉.休风电+本日.1高炉.休风电", dec_休风, null, false);
            AddNumericTag("本日.休风工艺", "本日.4高炉.休风工艺+本日.5高炉.休风工艺+本日.6高炉.休风工艺+本日.7高炉.休风工艺+本日.8高炉.休风工艺+本日.1高炉.休风工艺", dec_休风, null, false);
            AddNumericTag("本日.休风计划", "本日.4高炉.休风计划+本日.5高炉.休风计划+本日.6高炉.休风计划+本日.7高炉.休风计划+本日.8高炉.休风计划+本日.1高炉.休风计划", dec_休风, null, false);
            AddNumericTag("本日.休风设备", "本日.4高炉.休风设备+本日.5高炉.休风设备+本日.6高炉.休风设备+本日.7高炉.休风设备+本日.8高炉.休风设备+本日.1高炉.休风设备", dec_休风, null, false);
            AddNumericTag("本日.休风外围", "本日.4高炉.休风外围+本日.5高炉.休风外围+本日.6高炉.休风外围+本日.7高炉.休风外围+本日.8高炉.休风外围+本日.1高炉.休风外围", dec_休风, null, false);

            AddNumericTag("累计.1高炉.休风", "累计.1高炉.休风工艺+累计.1高炉.休风设备+累计.1高炉.休风电+累计.1高炉.休风外围+累计.1高炉.休风计划", dec_休风, null, false);
            AddNumericTag("累计.1高炉.休风电", "昨日累计.1高炉.休风电+本日.1高炉.休风电", dec_休风, null, true);
            AddNumericTag("累计.1高炉.休风工艺", "昨日累计.1高炉.休风工艺+本日.1高炉.休风工艺", dec_休风, null, true);
            AddNumericTag("累计.1高炉.休风计划", "昨日累计.1高炉.休风计划+本日.1高炉.休风计划", dec_休风, null, true);
            AddNumericTag("累计.1高炉.休风设备", "昨日累计.1高炉.休风设备+本日.1高炉.休风设备", dec_休风, null, true);
            AddNumericTag("累计.1高炉.休风外围", "昨日累计.1高炉.休风外围+本日.1高炉.休风外围", dec_休风, null, true);

            AddNumericTag("累计.4高炉.休风", "累计.4高炉.休风工艺+累计.4高炉.休风设备+累计.4高炉.休风电+累计.4高炉.休风外围+累计.4高炉.休风计划", dec_休风, null, false);
            AddNumericTag("累计.4高炉.休风电", "昨日累计.4高炉.休风电+本日.4高炉.休风电", dec_休风, null, true);
            AddNumericTag("累计.4高炉.休风工艺", "昨日累计.4高炉.休风工艺+本日.4高炉.休风工艺", dec_休风, null, true);
            AddNumericTag("累计.4高炉.休风计划", "昨日累计.4高炉.休风计划+本日.4高炉.休风计划", dec_休风, null, true);
            AddNumericTag("累计.4高炉.休风设备", "昨日累计.4高炉.休风设备+本日.4高炉.休风设备", dec_休风, null, true);
            AddNumericTag("累计.4高炉.休风外围", "昨日累计.4高炉.休风外围+本日.4高炉.休风外围", dec_休风, null, true);

            AddNumericTag("累计.5高炉.休风", "累计.5高炉.休风工艺+累计.5高炉.休风设备+累计.5高炉.休风电+累计.5高炉.休风外围+累计.5高炉.休风计划", dec_休风, null, false);
            AddNumericTag("累计.5高炉.休风电", "昨日累计.5高炉.休风电+本日.5高炉.休风电", dec_休风, null, true);
            AddNumericTag("累计.5高炉.休风工艺", "昨日累计.5高炉.休风工艺+本日.5高炉.休风工艺", dec_休风, null, true);
            AddNumericTag("累计.5高炉.休风计划", "昨日累计.5高炉.休风计划+本日.5高炉.休风计划", dec_休风, null, true);
            AddNumericTag("累计.5高炉.休风设备", "昨日累计.5高炉.休风设备+本日.5高炉.休风设备", dec_休风, null, true);
            AddNumericTag("累计.5高炉.休风外围", "昨日累计.5高炉.休风外围+本日.5高炉.休风外围", dec_休风, null, true);

            AddNumericTag("累计.6高炉.休风", "累计.6高炉.休风工艺+累计.6高炉.休风设备+累计.6高炉.休风电+累计.6高炉.休风外围+累计.6高炉.休风计划", dec_休风, null, false);
            AddNumericTag("累计.6高炉.休风电", "昨日累计.6高炉.休风电+本日.6高炉.休风电", dec_休风, null, true);
            AddNumericTag("累计.6高炉.休风工艺", "昨日累计.6高炉.休风工艺+本日.6高炉.休风工艺", dec_休风, null, true);
            AddNumericTag("累计.6高炉.休风计划", "昨日累计.6高炉.休风计划+本日.6高炉.休风计划", dec_休风, null, true);
            AddNumericTag("累计.6高炉.休风设备", "昨日累计.6高炉.休风设备+本日.6高炉.休风设备", dec_休风, null, true);
            AddNumericTag("累计.6高炉.休风外围", "昨日累计.6高炉.休风外围+本日.6高炉.休风外围", dec_休风, null, true);

            AddNumericTag("累计.7高炉.休风", "累计.7高炉.休风计划+累计.7高炉.休风外围+累计.7高炉.休风电+累计.7高炉.休风设备+累计.7高炉.休风工艺", dec_休风, null, false);
            AddNumericTag("累计.7高炉.休风电", "昨日累计.7高炉.休风电+本日.7高炉.休风电", dec_休风, null, true);
            AddNumericTag("累计.7高炉.休风工艺", "昨日累计.7高炉.休风工艺+本日.7高炉.休风工艺", dec_休风, null, true);
            AddNumericTag("累计.7高炉.休风计划", "昨日累计.7高炉.休风计划+本日.7高炉.休风计划", dec_休风, null, true);
            AddNumericTag("累计.7高炉.休风设备", "昨日累计.7高炉.休风设备+本日.7高炉.休风设备", dec_休风, null, true);
            AddNumericTag("累计.7高炉.休风外围", "昨日累计.7高炉.休风外围+本日.7高炉.休风外围", dec_休风, null, true);

            AddNumericTag("累计.8高炉.休风", "累计.8高炉.休风工艺+累计.8高炉.休风设备+累计.8高炉.休风电+累计.8高炉.休风外围+累计.8高炉.休风计划", dec_休风, null, false);
            AddNumericTag("累计.8高炉.休风电", "昨日累计.8高炉.休风电+本日.8高炉.休风电", dec_休风, null, true);
            AddNumericTag("累计.8高炉.休风工艺", "昨日累计.8高炉.休风工艺+本日.8高炉.休风工艺", dec_休风, null, true);
            AddNumericTag("累计.8高炉.休风计划", "昨日累计.8高炉.休风计划+本日.8高炉.休风计划", dec_休风, null, true);
            AddNumericTag("累计.8高炉.休风设备", "昨日累计.8高炉.休风设备+本日.8高炉.休风设备", dec_休风, null, true);
            AddNumericTag("累计.8高炉.休风外围", "昨日累计.8高炉.休风外围+本日.8高炉.休风外围", dec_休风, null, true);


            AddNumericTag("累计.456高炉合计.休风", "累计.456高炉合计.休风计划+累计.456高炉合计.休风外围+累计.456高炉合计.休风电+累计.456高炉合计.休风设备+累计.456高炉合计.休风工艺", dec_休风, null, false);
            AddNumericTag("累计.456高炉合计.休风电", "累计.4高炉.休风电+累计.5高炉.休风电+累计.6高炉.休风电", dec_休风, null, false);
            AddNumericTag("累计.456高炉合计.休风工艺", "累计.4高炉.休风工艺+累计.5高炉.休风工艺+累计.6高炉.休风工艺", dec_休风, null, false);
            AddNumericTag("累计.456高炉合计.休风计划", "累计.4高炉.休风计划+累计.5高炉.休风计划+累计.6高炉.休风计划", dec_休风, null, false);
            AddNumericTag("累计.456高炉合计.休风设备", "累计.4高炉.休风设备+累计.5高炉.休风设备+累计.6高炉.休风设备", dec_休风, null, false);
            AddNumericTag("累计.456高炉合计.休风外围", "累计.4高炉.休风外围+累计.5高炉.休风外围+累计.6高炉.休风外围", dec_休风, null, false);

            AddNumericTag("累计.78高炉合计.休风", "累计.78高炉合计.休风工艺+累计.78高炉合计.休风设备+累计.78高炉合计.休风电+累计.78高炉合计.休风外围+累计.78高炉合计.休风计划", dec_休风, null, false);
            AddNumericTag("累计.78高炉合计.休风电", "累计.7高炉.休风电+累计.8高炉.休风电", dec_休风, null, false);
            AddNumericTag("累计.78高炉合计.休风工艺", "累计.7高炉.休风工艺+累计.8高炉.休风工艺", dec_休风, null, false);
            AddNumericTag("累计.78高炉合计.休风计划", "累计.7高炉.休风计划+累计.8高炉.休风计划", dec_休风, null, false);
            AddNumericTag("累计.78高炉合计.休风设备", "累计.7高炉.休风设备+累计.8高炉.休风设备", dec_休风, null, false);
            AddNumericTag("累计.78高炉合计.休风外围", "累计.7高炉.休风外围+累计.8高炉.休风外围", dec_休风, null, false);

            AddNumericTag("累计.4_8高炉合计.休风", "累计.4_8高炉合计.休风工艺+累计.4_8高炉合计.休风设备+累计.4_8高炉合计.休风电+累计.4_8高炉合计.休风外围+累计.4_8高炉合计.休风计划", dec_休风, null, false);
            AddNumericTag("累计.4_8高炉合计.休风电", "累计.4高炉.休风电+累计.5高炉.休风电+累计.6高炉.休风电+累计.7高炉.休风电+累计.8高炉.休风电", dec_休风, null, false);
            AddNumericTag("累计.4_8高炉合计.休风工艺", "累计.4高炉.休风工艺+累计.5高炉.休风工艺+累计.6高炉.休风工艺+累计.7高炉.休风工艺+累计.8高炉.休风工艺", dec_休风, null, false);
            AddNumericTag("累计.4_8高炉合计.休风计划", "累计.4高炉.休风计划+累计.5高炉.休风计划+累计.6高炉.休风计划+累计.7高炉.休风计划+累计.8高炉.休风计划", dec_休风, null, false);
            AddNumericTag("累计.4_8高炉合计.休风设备", "累计.4高炉.休风设备+累计.5高炉.休风设备+累计.6高炉.休风设备+累计.7高炉.休风设备+累计.8高炉.休风设备", dec_休风, null, false);
            AddNumericTag("累计.4_8高炉合计.休风外围", "累计.4高炉.休风外围+累计.5高炉.休风外围+累计.6高炉.休风外围+累计.7高炉.休风外围+累计.8高炉.休风外围", dec_休风, null, false);

            AddNumericTag("累计.休风", "累计.休风工艺+累计.休风设备+累计.休风电+累计.休风外围+累计.休风计划", dec_休风, null, false);
            AddNumericTag("累计.休风电", "累计.4高炉.休风电+累计.5高炉.休风电+累计.6高炉.休风电+累计.7高炉.休风电+累计.8高炉.休风电+累计.1高炉.休风电", dec_休风, null, false);
            AddNumericTag("累计.休风工艺", "累计.4高炉.休风工艺+累计.5高炉.休风工艺+累计.6高炉.休风工艺+累计.7高炉.休风工艺+累计.8高炉.休风工艺+累计.1高炉.休风工艺", dec_休风, null, false);
            AddNumericTag("累计.休风计划", "累计.4高炉.休风计划+累计.5高炉.休风计划+累计.6高炉.休风计划+累计.7高炉.休风计划+累计.8高炉.休风计划+累计.1高炉.休风计划", dec_休风, null, false);
            AddNumericTag("累计.休风设备", "累计.4高炉.休风设备+累计.5高炉.休风设备+累计.6高炉.休风设备+累计.7高炉.休风设备+累计.8高炉.休风设备+累计.1高炉.休风设备", dec_休风, null, false);
            AddNumericTag("累计.休风外围", "累计.4高炉.休风外围+累计.5高炉.休风外围+累计.6高炉.休风外围+累计.7高炉.休风外围+累计.8高炉.休风外围+累计.1高炉.休风外围", dec_休风, null, false);


            AddNumericTag("昨日累计.1高炉.休风电", "", dec_休风, null, false);
            AddNumericTag("昨日累计.1高炉.休风工艺", "", dec_休风, null, false);
            AddNumericTag("昨日累计.1高炉.休风计划", "", dec_休风, null, false);
            AddNumericTag("昨日累计.1高炉.休风设备", "", dec_休风, null, false);
            AddNumericTag("昨日累计.1高炉.休风外围", "", dec_休风, null, false);

            AddNumericTag("昨日累计.4高炉.休风电", "", dec_休风, null, false);
            AddNumericTag("昨日累计.4高炉.休风工艺", "", dec_休风, null, false);
            AddNumericTag("昨日累计.4高炉.休风计划", "", dec_休风, null, false);
            AddNumericTag("昨日累计.4高炉.休风设备", "", dec_休风, null, false);
            AddNumericTag("昨日累计.4高炉.休风外围", "", dec_休风, null, false);

            AddNumericTag("昨日累计.5高炉.休风电", "", dec_休风, null, false);
            AddNumericTag("昨日累计.5高炉.休风工艺", "", dec_休风, null, false);
            AddNumericTag("昨日累计.5高炉.休风计划", "", dec_休风, null, false);
            AddNumericTag("昨日累计.5高炉.休风设备", "", dec_休风, null, false);
            AddNumericTag("昨日累计.5高炉.休风外围", "", dec_休风, null, false);

            AddNumericTag("昨日累计.6高炉.休风电", "", dec_休风, null, false);
            AddNumericTag("昨日累计.6高炉.休风工艺", "", dec_休风, null, false);
            AddNumericTag("昨日累计.6高炉.休风计划", "", dec_休风, null, false);
            AddNumericTag("昨日累计.6高炉.休风设备", "", dec_休风, null, false);
            AddNumericTag("昨日累计.6高炉.休风外围", "", dec_休风, null, false);

            AddNumericTag("昨日累计.7高炉.休风电", "", dec_休风, null, false);
            AddNumericTag("昨日累计.7高炉.休风工艺", "", dec_休风, null, false);
            AddNumericTag("昨日累计.7高炉.休风计划", "", dec_休风, null, false);
            AddNumericTag("昨日累计.7高炉.休风设备", "", dec_休风, null, false);
            AddNumericTag("昨日累计.7高炉.休风外围", "", dec_休风, null, false);

            AddNumericTag("昨日累计.8高炉.休风电", "", dec_休风, null, false);
            AddNumericTag("昨日累计.8高炉.休风工艺", "", dec_休风, null, false);
            AddNumericTag("昨日累计.8高炉.休风计划", "", dec_休风, null, false);
            AddNumericTag("昨日累计.8高炉.休风设备", "", dec_休风, null, false);
            AddNumericTag("昨日累计.8高炉.休风外围", "", dec_休风, null, false);

            #endregion

            // AddTag("本日.1高炉.休风时间", "=本日.1高炉.休风", null, null, true);
            #region 休风时间

            int? dec_休风时间 = 0;
            AddNumericTag("本日.1高炉.休风时间", "", dec_休风时间, null, true);
            AddNumericTag("本日.4高炉.休风时间", "", dec_休风时间, null, true);
            AddNumericTag("本日.5高炉.休风时间", "", dec_休风时间, null, true);
            AddNumericTag("本日.6高炉.休风时间", "", dec_休风时间, null, true);
            AddNumericTag("本日.7高炉.休风时间", "", dec_休风时间, null, true);
            AddNumericTag("本日.8高炉.休风时间", "", dec_休风时间, null, true);

            AddNumericTag("本日.456高炉合计.休风时间", "本日.4高炉.休风时间+本日.5高炉.休风时间+本日.6高炉.休风时间", dec_休风时间, null, false);
            AddNumericTag("本日.78高炉合计.休风时间", "本日.7高炉.休风时间+本日.8高炉.休风时间", dec_休风时间, null, false);
            AddNumericTag("本日.4_8高炉合计.休风时间", "本日.4高炉.休风时间+本日.5高炉.休风时间+本日.6高炉.休风时间+本日.7高炉.休风时间+本日.8高炉.休风时间", dec_休风时间, null, false);
            AddNumericTag("本日.休风时间", "本日.4高炉.休风时间+本日.5高炉.休风时间+本日.6高炉.休风时间+本日.7高炉.休风时间+本日.8高炉.休风时间+本日.1高炉.休风时间", dec_休风时间, null, false);

            AddNumericTag("累计.1高炉.休风时间", "昨日累计.1高炉.休风时间+本日.1高炉.休风时间", dec_休风时间, null, true);
            AddNumericTag("累计.4高炉.休风时间", "昨日累计.4高炉.休风时间+本日.4高炉.休风时间", dec_休风时间, null, true);
            AddNumericTag("累计.5高炉.休风时间", "昨日累计.5高炉.休风时间+本日.5高炉.休风时间", dec_休风时间, null, true);
            AddNumericTag("累计.6高炉.休风时间", "昨日累计.6高炉.休风时间+本日.6高炉.休风时间", dec_休风时间, null, true);
            AddNumericTag("累计.7高炉.休风时间", "昨日累计.7高炉.休风时间+本日.7高炉.休风时间", dec_休风时间, null, true);
            AddNumericTag("累计.8高炉.休风时间", "昨日累计.8高炉.休风时间+本日.8高炉.休风时间", dec_休风时间, null, true);

            AddNumericTag("累计.456高炉合计.休风时间", "累计.4高炉.休风时间+累计.5高炉.休风时间+累计.6高炉.休风时间", dec_休风时间, null, false);
            AddNumericTag("累计.78高炉合计.休风时间", "累计.7高炉.休风时间+累计.8高炉.休风时间", dec_休风时间, null, false);
            AddNumericTag("累计.4_8高炉合计.休风时间", "累计.4高炉.休风时间+累计.5高炉.休风时间+累计.6高炉.休风时间+累计.7高炉.休风时间+累计.8高炉.休风时间", dec_休风时间, null, false);
            AddNumericTag("累计.休风时间", "累计.4高炉.休风时间+累计.5高炉.休风时间+累计.6高炉.休风时间+累计.7高炉.休风时间+累计.8高炉.休风时间+累计.1高炉.休风时间", dec_休风时间, null, false);

            AddNumericTag("昨日累计.1高炉.休风时间", "", dec_休风时间, null, false);
            AddNumericTag("昨日累计.4高炉.休风时间", "", dec_休风时间, null, false);
            AddNumericTag("昨日累计.5高炉.休风时间", "", dec_休风时间, null, false);
            AddNumericTag("昨日累计.6高炉.休风时间", "", dec_休风时间, null, false);
            AddNumericTag("昨日累计.7高炉.休风时间", "", dec_休风时间, null, false);
            AddNumericTag("昨日累计.8高炉.休风时间", "", dec_休风时间, null, false);
            #endregion


            #region 休风率
            int? dec_休风率 = 2;
            AddNumericTag("本日.1高炉.休风率", "本日.1高炉.休风时间*100/1440", dec_休风率, null, false);
            AddNumericTag("本日.4高炉.休风率", "本日.4高炉.休风时间*100/1440", dec_休风率, null, false);
            AddNumericTag("本日.5高炉.休风率", "本日.5高炉.休风时间*100/1440", dec_休风率, null, false);
            AddNumericTag("本日.6高炉.休风率", "本日.6高炉.休风时间*100/1440", dec_休风率, null, false);
            AddNumericTag("本日.7高炉.休风率", "本日.7高炉.休风时间*100/1440", dec_休风率, null, false);
            AddNumericTag("本日.8高炉.休风率", "本日.8高炉.休风时间*100/1440", dec_休风率, null, false);
            AddNumericTag("本日.456高炉合计.休风率", "本日.456高炉合计.休风时间*100/(1440*3)", dec_休风率, null, false);
            AddNumericTag("本日.78高炉合计.休风率", "本日.78高炉合计.休风时间*100/(1440*2)", dec_休风率, null, false);
            AddNumericTag("本日.4_8高炉合计.休风率", "本日.4_8高炉合计.休风时间*100/(1440*5)", dec_休风率, null, false);
            AddNumericTag("本日.休风率", "本日.休风时间*100/(1440*6)", dec_休风率, null, false);

            AddNumericTag("累计.1高炉.休风率", "累计.1高炉.休风时间*100/(1440*累计.累计天数)", dec_休风率, null, false);
            AddNumericTag("累计.4高炉.休风率", "累计.4高炉.休风时间*100/(1440*累计.累计天数)", dec_休风率, null, false);
            AddNumericTag("累计.5高炉.休风率", "累计.5高炉.休风时间*100/(1440*累计.累计天数)", dec_休风率, null, false);
            AddNumericTag("累计.6高炉.休风率", "累计.6高炉.休风时间*100/(1440*累计.累计天数)", dec_休风率, null, false);
            AddNumericTag("累计.7高炉.休风率", "累计.7高炉.休风时间*100/(1440*累计.累计天数)", dec_休风率, null, false);
            AddNumericTag("累计.8高炉.休风率", "累计.8高炉.休风时间*100/(1440*累计.累计天数)", dec_休风率, null, false);
            AddNumericTag("累计.456高炉合计.休风率", "累计.456高炉合计.休风时间*100/(1440*3*累计.累计天数)", dec_休风率, null, false);
            AddNumericTag("累计.78高炉合计.休风率", "累计.78高炉合计.休风时间*100/(1440*2*累计.累计天数)", dec_休风率, null, false);
            AddNumericTag("累计.4_8高炉合计.休风率", "累计.4_8高炉合计.休风时间*100/(1440*5*累计.累计天数)", dec_休风率, null, false);
            AddNumericTag("累计.休风率", "累计.休风时间*100/(1440*6*累计.累计天数)", dec_休风率, null, false);

            #endregion

            #region 平均风温
            int? dec_平均风温 = 0;
            AddNumericTag("本日.1高炉.平均风温", "", dec_平均风温, null, true);
            AddNumericTag("本日.4高炉.平均风温", "", dec_平均风温, null, true);
            AddNumericTag("本日.5高炉.平均风温", "", dec_平均风温, null, true);
            AddNumericTag("本日.6高炉.平均风温", "", dec_平均风温, null, true);
            AddNumericTag("本日.7高炉.平均风温", "", dec_平均风温, null, true);
            AddNumericTag("本日.8高炉.平均风温", "", dec_平均风温, null, true);

            AddNumericTag("本日.456高炉合计.平均风温", "(本日.4高炉.平均风温+本日.5高炉.平均风温+本日.6高炉.平均风温)/3", dec_平均风温, null, false);
            AddNumericTag("本日.78高炉合计.平均风温", "(本日.7高炉.平均风温+本日.8高炉.平均风温)/2", dec_平均风温, null, false);
            AddNumericTag("本日.4_8高炉合计.平均风温", "(本日.4高炉.平均风温+本日.5高炉.平均风温+本日.6高炉.平均风温+本日.7高炉.平均风温+本日.8高炉.平均风温)/5", dec_平均风温, null, false);
            AddNumericTag("本日.平均风温", "(本日.4高炉.平均风温+本日.5高炉.平均风温+本日.6高炉.平均风温+本日.7高炉.平均风温+本日.8高炉.平均风温+本日.1高炉.平均风温)/6", dec_平均风温, null, false);

            AddNumericTag("累计.1高炉.平均风温", "", dec_平均风温, null, true);
            AddNumericTag("累计.4高炉.平均风温", "", dec_平均风温, null, true);
            AddNumericTag("累计.5高炉.平均风温", "", dec_平均风温, null, true);
            AddNumericTag("累计.6高炉.平均风温", "", dec_平均风温, null, true);
            AddNumericTag("累计.7高炉.平均风温", "", dec_平均风温, null, true);
            AddNumericTag("累计.8高炉.平均风温", "", dec_平均风温, null, true);

            AddNumericTag("累计.456高炉合计.平均风温", "(累计.4高炉.平均风温+累计.5高炉.平均风温+累计.6高炉.平均风温)/3", dec_平均风温, null, false);
            AddNumericTag("累计.78高炉合计.平均风温", "(累计.7高炉.平均风温+累计.8高炉.平均风温)/2", dec_平均风温, null, false);
            AddNumericTag("累计.4_8高炉合计.平均风温", "(累计.4高炉.平均风温+累计.5高炉.平均风温+累计.6高炉.平均风温+累计.7高炉.平均风温+累计.8高炉.平均风温)/5", dec_平均风温, null, false);
            AddNumericTag("累计.平均风温", "(累计.4高炉.平均风温+累计.5高炉.平均风温+累计.6高炉.平均风温+累计.7高炉.平均风温+累计.8高炉.平均风温+累计.1高炉.平均风温)/6", dec_平均风温, null, false);

            #endregion


            #region 生铁含硅
            int? dec_生铁含硅 = 3;
            AddNumericTag("本日.1高炉.生铁含硅", "", dec_生铁含硅, null, true);
            AddNumericTag("本日.4高炉.生铁含硅", "", dec_生铁含硅, null, true);
            AddNumericTag("本日.5高炉.生铁含硅", "", dec_生铁含硅, null, true);
            AddNumericTag("本日.6高炉.生铁含硅", "", dec_生铁含硅, null, true);
            AddNumericTag("本日.7高炉.生铁含硅", "", dec_生铁含硅, null, true);
            AddNumericTag("本日.8高炉.生铁含硅", "", dec_生铁含硅, null, true);

            AddNumericTag("本日.456高炉合计.生铁含硅", "(本日.4高炉.生铁含硅+本日.5高炉.生铁含硅+本日.6高炉.生铁含硅)/3", dec_生铁含硅, null, false);
            AddNumericTag("本日.78高炉合计.生铁含硅", "(本日.7高炉.生铁含硅+本日.8高炉.生铁含硅)/2", dec_生铁含硅, null, false);
            AddNumericTag("本日.4_8高炉合计.生铁含硅", "(本日.4高炉.生铁含硅+本日.5高炉.生铁含硅+本日.6高炉.生铁含硅+本日.7高炉.生铁含硅+本日.8高炉.生铁含硅)/5", dec_生铁含硅, null, false);
            AddNumericTag("本日.生铁含硅", "(本日.4高炉.生铁含硅+本日.5高炉.生铁含硅+本日.6高炉.生铁含硅+本日.7高炉.生铁含硅+本日.8高炉.生铁含硅+本日.1高炉.生铁含硅)/6", dec_生铁含硅, null, false);

            AddNumericTag("累计.1高炉.生铁含硅", "", dec_生铁含硅, null, true);
            AddNumericTag("累计.4高炉.生铁含硅", "", dec_生铁含硅, null, true);
            AddNumericTag("累计.5高炉.生铁含硅", "", dec_生铁含硅, null, true);
            AddNumericTag("累计.6高炉.生铁含硅", "", dec_生铁含硅, null, true);
            AddNumericTag("累计.7高炉.生铁含硅", "", dec_生铁含硅, null, true);
            AddNumericTag("累计.8高炉.生铁含硅", "", dec_生铁含硅, null, true);

            AddNumericTag("累计.456高炉合计.生铁含硅", "(累计.4高炉.生铁含硅+累计.5高炉.生铁含硅+累计.6高炉.生铁含硅)/3", dec_生铁含硅, null, false);
            AddNumericTag("累计.78高炉合计.生铁含硅", "(累计.7高炉.生铁含硅+累计.8高炉.生铁含硅)/2", dec_生铁含硅, null, false);
            AddNumericTag("累计.4_8高炉合计.生铁含硅", "(累计.4高炉.生铁含硅+累计.5高炉.生铁含硅+累计.6高炉.生铁含硅+累计.7高炉.生铁含硅+累计.8高炉.生铁含硅)/5", dec_生铁含硅, null, false);
            AddNumericTag("累计.生铁含硅", "(累计.4高炉.生铁含硅+累计.5高炉.生铁含硅+累计.6高炉.生铁含硅+累计.7高炉.生铁含硅+累计.8高炉.生铁含硅+累计.1高炉.生铁含硅)/6", dec_生铁含硅, null, false);

            #endregion

            //原料产出消耗
            #region 烧结一车间
            int? dec_烧结一车间 = 2;
            AddNumericTag("本日.供应.烧结一车间", "", dec_烧结一车间, null, true);
            AddNumericTag("本日.块矿面.烧结一车间", "", dec_烧结一车间, null, true);
            AddNumericTag("本日.1高炉.烧结一车间", "", dec_烧结一车间, null, true);
            AddNumericTag("本日.4高炉.烧结一车间", "", dec_烧结一车间, null, true);
            AddNumericTag("本日.5高炉.烧结一车间", "", dec_烧结一车间, null, true);
            AddNumericTag("本日.6高炉.烧结一车间", "", dec_烧结一车间, null, true);
            AddNumericTag("本日.7高炉.烧结一车间", "", dec_烧结一车间, null, true);
            AddNumericTag("本日.8高炉.烧结一车间", "", dec_烧结一车间, null, true);
            AddNumericTag("本日.扣返矿.烧结一车间", "", dec_烧结一车间, null, true);
            AddNumericTag("本日.烧结一车间", "本日.1高炉.烧结一车间+本日.4高炉.烧结一车间+本日.5高炉.烧结一车间+本日.6高炉.烧结一车间+本日.7高炉.烧结一车间+本日.8高炉.烧结一车间-本日.扣返矿.烧结一车间", dec_烧结一车间, null, true);

            AddNumericTag("本日库存.1高炉.烧结一车间", "", dec_烧结一车间, null, true);
            AddNumericTag("本日库存.二车间.烧结一车间", "", dec_烧结一车间, null, true);
            AddNumericTag("本日库存.三车间.烧结一车间", "", dec_烧结一车间, null, true);
            AddNumericTag("本日库存.烧结一车间", "本日库存.1高炉.烧结一车间+本日库存.二车间.烧结一车间+本日库存.三车间.烧结一车间", dec_烧结一车间, null, true);

            AddNumericTag("累计.供应.烧结一车间", "昨日累计.供应.烧结一车间+本日.供应.烧结一车间", dec_烧结一车间, null, true);
            AddNumericTag("累计.块矿面.烧结一车间", "昨日累计.块矿面.烧结一车间+本日.块矿面.烧结一车间", dec_烧结一车间, null, true);
            AddNumericTag("累计.1高炉.烧结一车间", "昨日累计.1高炉.烧结一车间+本日.1高炉.烧结一车间", dec_烧结一车间, null, true);
            AddNumericTag("累计.4高炉.烧结一车间", "昨日累计.4高炉.烧结一车间+本日.4高炉.烧结一车间", dec_烧结一车间, null, true);
            AddNumericTag("累计.5高炉.烧结一车间", "昨日累计.5高炉.烧结一车间+本日.5高炉.烧结一车间", dec_烧结一车间, null, true);
            AddNumericTag("累计.6高炉.烧结一车间", "昨日累计.6高炉.烧结一车间+本日.6高炉.烧结一车间", dec_烧结一车间, null, true);
            AddNumericTag("累计.7高炉.烧结一车间", "昨日累计.7高炉.烧结一车间+本日.7高炉.烧结一车间", dec_烧结一车间, null, true);
            AddNumericTag("累计.8高炉.烧结一车间", "昨日累计.8高炉.烧结一车间+本日.8高炉.烧结一车间", dec_烧结一车间, null, true);
            AddNumericTag("累计.扣返矿.烧结一车间", "昨日累计.扣返矿.烧结一车间-本日.扣返矿.烧结一车间", dec_烧结一车间, null, true);
            AddNumericTag("累计.烧结一车间", "累计.1高炉.烧结一车间+累计.4高炉.烧结一车间+累计.5高炉.烧结一车间+累计.6高炉.烧结一车间+累计.7高炉.烧结一车间+累计.8高炉.烧结一车间-累计.扣返矿.烧结一车间", dec_烧结一车间, null, false);






            AddNumericTag("昨日累计.供应.烧结一车间", "", dec_烧结一车间, null, false);
            AddNumericTag("昨日累计.块矿面.烧结一车间", "", dec_烧结一车间, null, false);
            AddNumericTag("昨日累计.1高炉.烧结一车间", "", dec_烧结一车间, null, false);
            AddNumericTag("昨日累计.4高炉.烧结一车间", "", dec_烧结一车间, null, false);
            AddNumericTag("昨日累计.5高炉.烧结一车间", "", dec_烧结一车间, null, false);
            AddNumericTag("昨日累计.6高炉.烧结一车间", "", dec_烧结一车间, null, false);
            AddNumericTag("昨日累计.7高炉.烧结一车间", "", dec_烧结一车间, null, false);
            AddNumericTag("昨日累计.8高炉.烧结一车间", "", dec_烧结一车间, null, false);
            AddNumericTag("昨日累计.扣返矿.烧结一车间", "", dec_烧结一车间, null, false);




            #endregion

            #region 烧结二车间
            int? dec_烧结二车间 = 2;
            AddNumericTag("本日.供应.烧结二车间", "", dec_烧结二车间, null, true);
            AddNumericTag("本日.块矿面.烧结二车间", "", dec_烧结二车间, null, true);
            AddNumericTag("本日.1高炉.烧结二车间", "", dec_烧结二车间, null, true);
            AddNumericTag("本日.4高炉.烧结二车间", "", dec_烧结二车间, null, true);
            AddNumericTag("本日.5高炉.烧结二车间", "", dec_烧结二车间, null, true);
            AddNumericTag("本日.6高炉.烧结二车间", "", dec_烧结二车间, null, true);
            AddNumericTag("本日.7高炉.烧结二车间", "", dec_烧结二车间, null, true);
            AddNumericTag("本日.8高炉.烧结二车间", "", dec_烧结二车间, null, true);
            AddNumericTag("本日.扣返矿.烧结二车间", "", dec_烧结二车间, null, true);
            AddNumericTag("本日.烧结二车间", "本日.1高炉.烧结二车间+本日.4高炉.烧结二车间+本日.5高炉.烧结二车间+本日.6高炉.烧结二车间+本日.7高炉.烧结二车间+本日.8高炉.烧结二车间-本日.扣返矿.烧结二车间", dec_烧结二车间, null, true);

            AddNumericTag("本日库存.1高炉.烧结二车间", "", dec_烧结二车间, null, true);
            AddNumericTag("本日库存.二车间.烧结二车间", "", dec_烧结二车间, null, true);
            AddNumericTag("本日库存.三车间.烧结二车间", "", dec_烧结二车间, null, true);
            AddNumericTag("本日库存.烧结二车间", "本日库存.1高炉.烧结二车间+本日库存.二车间.烧结二车间+本日库存.三车间.烧结二车间", dec_烧结二车间, null, true);

            AddNumericTag("累计.供应.烧结二车间", "昨日累计.供应.烧结二车间+本日.供应.烧结二车间", dec_烧结二车间, null, true);
            AddNumericTag("累计.块矿面.烧结二车间", "昨日累计.块矿面.烧结二车间+本日.块矿面.烧结二车间", dec_烧结二车间, null, true);
            AddNumericTag("累计.1高炉.烧结二车间", "昨日累计.1高炉.烧结二车间+本日.1高炉.烧结二车间", dec_烧结二车间, null, true);
            AddNumericTag("累计.4高炉.烧结二车间", "昨日累计.4高炉.烧结二车间+本日.4高炉.烧结二车间", dec_烧结二车间, null, true);
            AddNumericTag("累计.5高炉.烧结二车间", "昨日累计.5高炉.烧结二车间+本日.5高炉.烧结二车间", dec_烧结二车间, null, true);
            AddNumericTag("累计.6高炉.烧结二车间", "昨日累计.6高炉.烧结二车间+本日.6高炉.烧结二车间", dec_烧结二车间, null, true);
            AddNumericTag("累计.7高炉.烧结二车间", "昨日累计.7高炉.烧结二车间+本日.7高炉.烧结二车间", dec_烧结二车间, null, true);
            AddNumericTag("累计.8高炉.烧结二车间", "昨日累计.8高炉.烧结二车间+本日.8高炉.烧结二车间", dec_烧结二车间, null, true);
            AddNumericTag("累计.扣返矿.烧结二车间", "昨日累计.扣返矿.烧结二车间-本日.扣返矿.烧结二车间", dec_烧结二车间, null, true);
            AddNumericTag("累计.烧结二车间", "累计.1高炉.烧结二车间+累计.4高炉.烧结二车间+累计.5高炉.烧结二车间+累计.6高炉.烧结二车间+累计.7高炉.烧结二车间+累计.8高炉.烧结二车间-累计.扣返矿.烧结二车间", dec_烧结二车间, null, false);



            AddNumericTag("昨日累计.供应.烧结二车间", "", dec_烧结二车间, null, false);
            AddNumericTag("昨日累计.块矿面.烧结二车间", "", dec_烧结二车间, null, false);
            AddNumericTag("昨日累计.1高炉.烧结二车间", "", dec_烧结二车间, null, false);
            AddNumericTag("昨日累计.4高炉.烧结二车间", "", dec_烧结二车间, null, false);
            AddNumericTag("昨日累计.5高炉.烧结二车间", "", dec_烧结二车间, null, false);
            AddNumericTag("昨日累计.6高炉.烧结二车间", "", dec_烧结二车间, null, false);
            AddNumericTag("昨日累计.7高炉.烧结二车间", "", dec_烧结二车间, null, false);
            AddNumericTag("昨日累计.8高炉.烧结二车间", "", dec_烧结二车间, null, false);
            AddNumericTag("昨日累计.扣返矿.烧结二车间", "", dec_烧结二车间, null, false);




            #endregion

            #region 230一车间
            int? dec_230一车间 = 2;
            AddNumericTag("本日.供应.230一车间", "", dec_230一车间, null, true);
            AddNumericTag("本日.块矿面.230一车间", "", dec_230一车间, null, true);
            AddNumericTag("本日.1高炉.230一车间", "", dec_230一车间, null, true);
            AddNumericTag("本日.4高炉.230一车间", "", dec_230一车间, null, true);
            AddNumericTag("本日.5高炉.230一车间", "", dec_230一车间, null, true);
            AddNumericTag("本日.6高炉.230一车间", "", dec_230一车间, null, true);
            AddNumericTag("本日.7高炉.230一车间", "", dec_230一车间, null, true);
            AddNumericTag("本日.8高炉.230一车间", "", dec_230一车间, null, true);
            AddNumericTag("本日.扣返矿.230一车间", "", dec_230一车间, null, true);
            AddNumericTag("本日.230一车间", "本日.1高炉.230一车间+本日.4高炉.230一车间+本日.5高炉.230一车间+本日.6高炉.230一车间+本日.7高炉.230一车间+本日.8高炉.230一车间-本日.扣返矿.230一车间", dec_230一车间, null, true);

            AddNumericTag("本日库存.1高炉.230一车间", "", dec_230一车间, null, true);
            AddNumericTag("本日库存.二车间.230一车间", "", dec_230一车间, null, true);
            AddNumericTag("本日库存.三车间.230一车间", "", dec_230一车间, null, true);
            AddNumericTag("本日库存.230一车间", "本日库存.1高炉.230一车间+本日库存.二车间.230一车间+本日库存.三车间.230一车间", dec_230一车间, null, true);

            AddNumericTag("累计.供应.230一车间", "昨日累计.供应.230一车间+本日.供应.230一车间", dec_230一车间, null, true);
            AddNumericTag("累计.块矿面.230一车间", "昨日累计.块矿面.230一车间+本日.块矿面.230一车间", dec_230一车间, null, true);
            AddNumericTag("累计.1高炉.230一车间", "昨日累计.1高炉.230一车间+本日.1高炉.230一车间", dec_230一车间, null, true);
            AddNumericTag("累计.4高炉.230一车间", "昨日累计.4高炉.230一车间+本日.4高炉.230一车间", dec_230一车间, null, true);
            AddNumericTag("累计.5高炉.230一车间", "昨日累计.5高炉.230一车间+本日.5高炉.230一车间", dec_230一车间, null, true);
            AddNumericTag("累计.6高炉.230一车间", "昨日累计.6高炉.230一车间+本日.6高炉.230一车间", dec_230一车间, null, true);
            AddNumericTag("累计.7高炉.230一车间", "昨日累计.7高炉.230一车间+本日.7高炉.230一车间", dec_230一车间, null, true);
            AddNumericTag("累计.8高炉.230一车间", "昨日累计.8高炉.230一车间+本日.8高炉.230一车间", dec_230一车间, null, true);
            AddNumericTag("累计.扣返矿.230一车间", "昨日累计.扣返矿.230一车间-本日.扣返矿.230一车间", dec_230一车间, null, true);
            AddNumericTag("累计.230一车间", "累计.1高炉.230一车间+累计.4高炉.230一车间+累计.5高炉.230一车间+累计.6高炉.230一车间+累计.7高炉.230一车间+累计.8高炉.230一车间-累计.扣返矿.230一车间", dec_230一车间, null, false);






            AddNumericTag("昨日累计.供应.230一车间", "", dec_230一车间, null, false);
            AddNumericTag("昨日累计.块矿面.230一车间", "", dec_230一车间, null, false);
            AddNumericTag("昨日累计.1高炉.230一车间", "", dec_230一车间, null, false);
            AddNumericTag("昨日累计.4高炉.230一车间", "", dec_230一车间, null, false);
            AddNumericTag("昨日累计.5高炉.230一车间", "", dec_230一车间, null, false);
            AddNumericTag("昨日累计.6高炉.230一车间", "", dec_230一车间, null, false);
            AddNumericTag("昨日累计.7高炉.230一车间", "", dec_230一车间, null, false);
            AddNumericTag("昨日累计.8高炉.230一车间", "", dec_230一车间, null, false);
            AddNumericTag("昨日累计.扣返矿.230一车间", "", dec_230一车间, null, false);




            #endregion

            #region 230二车间
            int? dec_230二车间 = 2;
            AddNumericTag("本日.供应.230二车间", "", dec_230二车间, null, true);
            AddNumericTag("本日.块矿面.230二车间", "", dec_230二车间, null, true);
            AddNumericTag("本日.1高炉.230二车间", "", dec_230二车间, null, true);
            AddNumericTag("本日.4高炉.230二车间", "", dec_230二车间, null, true);
            AddNumericTag("本日.5高炉.230二车间", "", dec_230二车间, null, true);
            AddNumericTag("本日.6高炉.230二车间", "", dec_230二车间, null, true);
            AddNumericTag("本日.7高炉.230二车间", "", dec_230二车间, null, true);
            AddNumericTag("本日.8高炉.230二车间", "", dec_230二车间, null, true);
            AddNumericTag("本日.扣返矿.230二车间", "", dec_230二车间, null, true);
            AddNumericTag("本日.230二车间", "本日.1高炉.230二车间+本日.4高炉.230二车间+本日.5高炉.230二车间+本日.6高炉.230二车间+本日.7高炉.230二车间+本日.8高炉.230二车间-本日.扣返矿.230二车间", dec_230二车间, null, true);

            AddNumericTag("本日库存.1高炉.230二车间", "", dec_230二车间, null, true);
            AddNumericTag("本日库存.二车间.230二车间", "", dec_230二车间, null, true);
            AddNumericTag("本日库存.三车间.230二车间", "", dec_230二车间, null, true);
            AddNumericTag("本日库存.230二车间", "本日库存.1高炉.230二车间+本日库存.二车间.230二车间+本日库存.三车间.230二车间", dec_230二车间, null, true);

            AddNumericTag("累计.供应.230二车间", "昨日累计.供应.230二车间+本日.供应.230二车间", dec_230二车间, null, true);
            AddNumericTag("累计.块矿面.230二车间", "昨日累计.块矿面.230二车间+本日.块矿面.230二车间", dec_230二车间, null, true);
            AddNumericTag("累计.1高炉.230二车间", "昨日累计.1高炉.230二车间+本日.1高炉.230二车间", dec_230二车间, null, true);
            AddNumericTag("累计.4高炉.230二车间", "昨日累计.4高炉.230二车间+本日.4高炉.230二车间", dec_230二车间, null, true);
            AddNumericTag("累计.5高炉.230二车间", "昨日累计.5高炉.230二车间+本日.5高炉.230二车间", dec_230二车间, null, true);
            AddNumericTag("累计.6高炉.230二车间", "昨日累计.6高炉.230二车间+本日.6高炉.230二车间", dec_230二车间, null, true);
            AddNumericTag("累计.7高炉.230二车间", "昨日累计.7高炉.230二车间+本日.7高炉.230二车间", dec_230二车间, null, true);
            AddNumericTag("累计.8高炉.230二车间", "昨日累计.8高炉.230二车间+本日.8高炉.230二车间", dec_230二车间, null, true);
            AddNumericTag("累计.扣返矿.230二车间", "昨日累计.扣返矿.230二车间-本日.扣返矿.230二车间", dec_230二车间, null, true);
            AddNumericTag("累计.230二车间", "累计.1高炉.230二车间+累计.4高炉.230二车间+累计.5高炉.230二车间+累计.6高炉.230二车间+累计.7高炉.230二车间+累计.8高炉.230二车间-累计.扣返矿.230二车间", dec_230二车间, null, false);






            AddNumericTag("昨日累计.供应.230二车间", "", dec_230二车间, null, false);
            AddNumericTag("昨日累计.块矿面.230二车间", "", dec_230二车间, null, false);
            AddNumericTag("昨日累计.1高炉.230二车间", "", dec_230二车间, null, false);
            AddNumericTag("昨日累计.4高炉.230二车间", "", dec_230二车间, null, false);
            AddNumericTag("昨日累计.5高炉.230二车间", "", dec_230二车间, null, false);
            AddNumericTag("昨日累计.6高炉.230二车间", "", dec_230二车间, null, false);
            AddNumericTag("昨日累计.7高炉.230二车间", "", dec_230二车间, null, false);
            AddNumericTag("昨日累计.8高炉.230二车间", "", dec_230二车间, null, false);
            AddNumericTag("昨日累计.扣返矿.230二车间", "", dec_230二车间, null, false);




            #endregion

            #region 260烧结
            int? dec_260烧结 = 2;
            AddNumericTag("本日.供应.260烧结", "", dec_260烧结, null, true);
            AddNumericTag("本日.块矿面.260烧结", "", dec_260烧结, null, true);
            AddNumericTag("本日.1高炉.260烧结", "", dec_260烧结, null, true);
            AddNumericTag("本日.4高炉.260烧结", "", dec_260烧结, null, true);
            AddNumericTag("本日.5高炉.260烧结", "", dec_260烧结, null, true);
            AddNumericTag("本日.6高炉.260烧结", "", dec_260烧结, null, true);
            AddNumericTag("本日.7高炉.260烧结", "", dec_260烧结, null, true);
            AddNumericTag("本日.8高炉.260烧结", "", dec_260烧结, null, true);
            AddNumericTag("本日.扣返矿.260烧结", "", dec_260烧结, null, true);
            AddNumericTag("本日.260烧结", "本日.1高炉.260烧结+本日.4高炉.260烧结+本日.5高炉.260烧结+本日.6高炉.260烧结+本日.7高炉.260烧结+本日.8高炉.260烧结-本日.扣返矿.260烧结", dec_260烧结, null, true);

            AddNumericTag("本日库存.1高炉.260烧结", "", dec_260烧结, null, true);
            AddNumericTag("本日库存.二车间.260烧结", "", dec_260烧结, null, true);
            AddNumericTag("本日库存.三车间.260烧结", "", dec_260烧结, null, true);
            AddNumericTag("本日库存.260烧结", "本日库存.1高炉.260烧结+本日库存.二车间.260烧结+本日库存.三车间.260烧结", dec_260烧结, null, true);

            AddNumericTag("累计.供应.260烧结", "昨日累计.供应.260烧结+本日.供应.260烧结", dec_260烧结, null, true);
            AddNumericTag("累计.块矿面.260烧结", "昨日累计.块矿面.260烧结+本日.块矿面.260烧结", dec_260烧结, null, true);
            AddNumericTag("累计.1高炉.260烧结", "昨日累计.1高炉.260烧结+本日.1高炉.260烧结", dec_260烧结, null, true);
            AddNumericTag("累计.4高炉.260烧结", "昨日累计.4高炉.260烧结+本日.4高炉.260烧结", dec_260烧结, null, true);
            AddNumericTag("累计.5高炉.260烧结", "昨日累计.5高炉.260烧结+本日.5高炉.260烧结", dec_260烧结, null, true);
            AddNumericTag("累计.6高炉.260烧结", "昨日累计.6高炉.260烧结+本日.6高炉.260烧结", dec_260烧结, null, true);
            AddNumericTag("累计.7高炉.260烧结", "昨日累计.7高炉.260烧结+本日.7高炉.260烧结", dec_260烧结, null, true);
            AddNumericTag("累计.8高炉.260烧结", "昨日累计.8高炉.260烧结+本日.8高炉.260烧结", dec_260烧结, null, true);
            AddNumericTag("累计.扣返矿.260烧结", "昨日累计.扣返矿.260烧结-本日.扣返矿.260烧结", dec_260烧结, null, true);
            AddNumericTag("累计.260烧结", "累计.1高炉.260烧结+累计.4高炉.260烧结+累计.5高炉.260烧结+累计.6高炉.260烧结+累计.7高炉.260烧结+累计.8高炉.260烧结-累计.扣返矿.260烧结", dec_260烧结, null, false);






            AddNumericTag("昨日累计.供应.260烧结", "", dec_260烧结, null, false);
            AddNumericTag("昨日累计.块矿面.260烧结", "", dec_260烧结, null, false);
            AddNumericTag("昨日累计.1高炉.260烧结", "", dec_260烧结, null, false);
            AddNumericTag("昨日累计.4高炉.260烧结", "", dec_260烧结, null, false);
            AddNumericTag("昨日累计.5高炉.260烧结", "", dec_260烧结, null, false);
            AddNumericTag("昨日累计.6高炉.260烧结", "", dec_260烧结, null, false);
            AddNumericTag("昨日累计.7高炉.260烧结", "", dec_260烧结, null, false);
            AddNumericTag("昨日累计.8高炉.260烧结", "", dec_260烧结, null, false);
            AddNumericTag("昨日累计.扣返矿.260烧结", "", dec_260烧结, null, false);




            #endregion

            #region 第二炼铁厂烧结
            int? dec_第二炼铁厂烧结 = 2;
            AddNumericTag("本日.供应.第二炼铁厂烧结", "", dec_第二炼铁厂烧结, null, true);
            AddNumericTag("本日.块矿面.第二炼铁厂烧结", "", dec_第二炼铁厂烧结, null, true);
            AddNumericTag("本日.1高炉.第二炼铁厂烧结", "", dec_第二炼铁厂烧结, null, true);
            AddNumericTag("本日.4高炉.第二炼铁厂烧结", "", dec_第二炼铁厂烧结, null, true);
            AddNumericTag("本日.5高炉.第二炼铁厂烧结", "", dec_第二炼铁厂烧结, null, true);
            AddNumericTag("本日.6高炉.第二炼铁厂烧结", "", dec_第二炼铁厂烧结, null, true);
            AddNumericTag("本日.7高炉.第二炼铁厂烧结", "", dec_第二炼铁厂烧结, null, true);
            AddNumericTag("本日.8高炉.第二炼铁厂烧结", "", dec_第二炼铁厂烧结, null, true);
            AddNumericTag("本日.扣返矿.第二炼铁厂烧结", "", dec_第二炼铁厂烧结, null, true);
            AddNumericTag("本日.第二炼铁厂烧结", "本日.1高炉.第二炼铁厂烧结+本日.4高炉.第二炼铁厂烧结+本日.5高炉.第二炼铁厂烧结+本日.6高炉.第二炼铁厂烧结+本日.7高炉.第二炼铁厂烧结+本日.8高炉.第二炼铁厂烧结-本日.扣返矿.第二炼铁厂烧结", dec_第二炼铁厂烧结, null, true);

            AddNumericTag("本日库存.1高炉.第二炼铁厂烧结", "", dec_第二炼铁厂烧结, null, true);
            AddNumericTag("本日库存.二车间.第二炼铁厂烧结", "", dec_第二炼铁厂烧结, null, true);
            AddNumericTag("本日库存.三车间.第二炼铁厂烧结", "", dec_第二炼铁厂烧结, null, true);
            AddNumericTag("本日库存.第二炼铁厂烧结", "本日库存.1高炉.第二炼铁厂烧结+本日库存.二车间.第二炼铁厂烧结+本日库存.三车间.第二炼铁厂烧结", dec_第二炼铁厂烧结, null, true);

            AddNumericTag("累计.供应.第二炼铁厂烧结", "昨日累计.供应.第二炼铁厂烧结+本日.供应.第二炼铁厂烧结", dec_第二炼铁厂烧结, null, true);
            AddNumericTag("累计.块矿面.第二炼铁厂烧结", "昨日累计.块矿面.第二炼铁厂烧结+本日.块矿面.第二炼铁厂烧结", dec_第二炼铁厂烧结, null, true);
            AddNumericTag("累计.1高炉.第二炼铁厂烧结", "昨日累计.1高炉.第二炼铁厂烧结+本日.1高炉.第二炼铁厂烧结", dec_第二炼铁厂烧结, null, true);
            AddNumericTag("累计.4高炉.第二炼铁厂烧结", "昨日累计.4高炉.第二炼铁厂烧结+本日.4高炉.第二炼铁厂烧结", dec_第二炼铁厂烧结, null, true);
            AddNumericTag("累计.5高炉.第二炼铁厂烧结", "昨日累计.5高炉.第二炼铁厂烧结+本日.5高炉.第二炼铁厂烧结", dec_第二炼铁厂烧结, null, true);
            AddNumericTag("累计.6高炉.第二炼铁厂烧结", "昨日累计.6高炉.第二炼铁厂烧结+本日.6高炉.第二炼铁厂烧结", dec_第二炼铁厂烧结, null, true);
            AddNumericTag("累计.7高炉.第二炼铁厂烧结", "昨日累计.7高炉.第二炼铁厂烧结+本日.7高炉.第二炼铁厂烧结", dec_第二炼铁厂烧结, null, true);
            AddNumericTag("累计.8高炉.第二炼铁厂烧结", "昨日累计.8高炉.第二炼铁厂烧结+本日.8高炉.第二炼铁厂烧结", dec_第二炼铁厂烧结, null, true);
            AddNumericTag("累计.扣返矿.第二炼铁厂烧结", "昨日累计.扣返矿.第二炼铁厂烧结-本日.扣返矿.第二炼铁厂烧结", dec_第二炼铁厂烧结, null, true);
            AddNumericTag("累计.第二炼铁厂烧结", "累计.1高炉.第二炼铁厂烧结+累计.4高炉.第二炼铁厂烧结+累计.5高炉.第二炼铁厂烧结+累计.6高炉.第二炼铁厂烧结+累计.7高炉.第二炼铁厂烧结+累计.8高炉.第二炼铁厂烧结-累计.扣返矿.第二炼铁厂烧结", dec_第二炼铁厂烧结, null, false);






            AddNumericTag("昨日累计.供应.第二炼铁厂烧结", "", dec_第二炼铁厂烧结, null, false);
            AddNumericTag("昨日累计.块矿面.第二炼铁厂烧结", "", dec_第二炼铁厂烧结, null, false);
            AddNumericTag("昨日累计.1高炉.第二炼铁厂烧结", "", dec_第二炼铁厂烧结, null, false);
            AddNumericTag("昨日累计.4高炉.第二炼铁厂烧结", "", dec_第二炼铁厂烧结, null, false);
            AddNumericTag("昨日累计.5高炉.第二炼铁厂烧结", "", dec_第二炼铁厂烧结, null, false);
            AddNumericTag("昨日累计.6高炉.第二炼铁厂烧结", "", dec_第二炼铁厂烧结, null, false);
            AddNumericTag("昨日累计.7高炉.第二炼铁厂烧结", "", dec_第二炼铁厂烧结, null, false);
            AddNumericTag("昨日累计.8高炉.第二炼铁厂烧结", "", dec_第二炼铁厂烧结, null, false);
            AddNumericTag("昨日累计.扣返矿.第二炼铁厂烧结", "", dec_第二炼铁厂烧结, null, false);




            #endregion

            #region 球团一车间
            int? dec_球团一车间 = 2;
            AddNumericTag("本日.供应.球团一车间", "", dec_球团一车间, null, true);
            AddNumericTag("本日.块矿面.球团一车间", "", dec_球团一车间, null, true);
            AddNumericTag("本日.1高炉.球团一车间", "", dec_球团一车间, null, true);
            AddNumericTag("本日.4高炉.球团一车间", "", dec_球团一车间, null, true);
            AddNumericTag("本日.5高炉.球团一车间", "", dec_球团一车间, null, true);
            AddNumericTag("本日.6高炉.球团一车间", "", dec_球团一车间, null, true);
            AddNumericTag("本日.7高炉.球团一车间", "", dec_球团一车间, null, true);
            AddNumericTag("本日.8高炉.球团一车间", "", dec_球团一车间, null, true);
            AddNumericTag("本日.扣返矿.球团一车间", "", dec_球团一车间, null, true);
            AddNumericTag("本日.球团一车间", "本日.1高炉.球团一车间+本日.4高炉.球团一车间+本日.5高炉.球团一车间+本日.6高炉.球团一车间+本日.7高炉.球团一车间+本日.8高炉.球团一车间-本日.扣返矿.球团一车间", dec_球团一车间, null, true);

            AddNumericTag("本日库存.1高炉.球团一车间", "", dec_球团一车间, null, true);
            AddNumericTag("本日库存.二车间.球团一车间", "", dec_球团一车间, null, true);
            AddNumericTag("本日库存.三车间.球团一车间", "", dec_球团一车间, null, true);
            AddNumericTag("本日库存.球团一车间", "本日库存.1高炉.球团一车间+本日库存.二车间.球团一车间+本日库存.三车间.球团一车间", dec_球团一车间, null, true);

            AddNumericTag("累计.供应.球团一车间", "昨日累计.供应.球团一车间+本日.供应.球团一车间", dec_球团一车间, null, true);
            AddNumericTag("累计.块矿面.球团一车间", "昨日累计.块矿面.球团一车间+本日.块矿面.球团一车间", dec_球团一车间, null, true);
            AddNumericTag("累计.1高炉.球团一车间", "昨日累计.1高炉.球团一车间+本日.1高炉.球团一车间", dec_球团一车间, null, true);
            AddNumericTag("累计.4高炉.球团一车间", "昨日累计.4高炉.球团一车间+本日.4高炉.球团一车间", dec_球团一车间, null, true);
            AddNumericTag("累计.5高炉.球团一车间", "昨日累计.5高炉.球团一车间+本日.5高炉.球团一车间", dec_球团一车间, null, true);
            AddNumericTag("累计.6高炉.球团一车间", "昨日累计.6高炉.球团一车间+本日.6高炉.球团一车间", dec_球团一车间, null, true);
            AddNumericTag("累计.7高炉.球团一车间", "昨日累计.7高炉.球团一车间+本日.7高炉.球团一车间", dec_球团一车间, null, true);
            AddNumericTag("累计.8高炉.球团一车间", "昨日累计.8高炉.球团一车间+本日.8高炉.球团一车间", dec_球团一车间, null, true);
            AddNumericTag("累计.扣返矿.球团一车间", "昨日累计.扣返矿.球团一车间-本日.扣返矿.球团一车间", dec_球团一车间, null, true);
            AddNumericTag("累计.球团一车间", "累计.1高炉.球团一车间+累计.4高炉.球团一车间+累计.5高炉.球团一车间+累计.6高炉.球团一车间+累计.7高炉.球团一车间+累计.8高炉.球团一车间-累计.扣返矿.球团一车间", dec_球团一车间, null, false);






            AddNumericTag("昨日累计.供应.球团一车间", "", dec_球团一车间, null, false);
            AddNumericTag("昨日累计.块矿面.球团一车间", "", dec_球团一车间, null, false);
            AddNumericTag("昨日累计.1高炉.球团一车间", "", dec_球团一车间, null, false);
            AddNumericTag("昨日累计.4高炉.球团一车间", "", dec_球团一车间, null, false);
            AddNumericTag("昨日累计.5高炉.球团一车间", "", dec_球团一车间, null, false);
            AddNumericTag("昨日累计.6高炉.球团一车间", "", dec_球团一车间, null, false);
            AddNumericTag("昨日累计.7高炉.球团一车间", "", dec_球团一车间, null, false);
            AddNumericTag("昨日累计.8高炉.球团一车间", "", dec_球团一车间, null, false);
            AddNumericTag("昨日累计.扣返矿.球团一车间", "", dec_球团一车间, null, false);




            #endregion

            #region 球团二车间
            int? dec_球团二车间 = 2;
            AddNumericTag("本日.供应.球团二车间", "", dec_球团二车间, null, true);
            AddNumericTag("本日.块矿面.球团二车间", "", dec_球团二车间, null, true);
            AddNumericTag("本日.1高炉.球团二车间", "", dec_球团二车间, null, true);
            AddNumericTag("本日.4高炉.球团二车间", "", dec_球团二车间, null, true);
            AddNumericTag("本日.5高炉.球团二车间", "", dec_球团二车间, null, true);
            AddNumericTag("本日.6高炉.球团二车间", "", dec_球团二车间, null, true);
            AddNumericTag("本日.7高炉.球团二车间", "", dec_球团二车间, null, true);
            AddNumericTag("本日.8高炉.球团二车间", "", dec_球团二车间, null, true);
            AddNumericTag("本日.扣返矿.球团二车间", "", dec_球团二车间, null, true);
            AddNumericTag("本日.球团二车间", "本日.1高炉.球团二车间+本日.4高炉.球团二车间+本日.5高炉.球团二车间+本日.6高炉.球团二车间+本日.7高炉.球团二车间+本日.8高炉.球团二车间-本日.扣返矿.球团二车间", dec_球团二车间, null, true);

            AddNumericTag("本日库存.1高炉.球团二车间", "", dec_球团二车间, null, true);
            AddNumericTag("本日库存.二车间.球团二车间", "", dec_球团二车间, null, true);
            AddNumericTag("本日库存.三车间.球团二车间", "", dec_球团二车间, null, true);
            AddNumericTag("本日库存.球团二车间", "本日库存.1高炉.球团二车间+本日库存.二车间.球团二车间+本日库存.三车间.球团二车间", dec_球团二车间, null, true);

            AddNumericTag("累计.供应.球团二车间", "昨日累计.供应.球团二车间+本日.供应.球团二车间", dec_球团二车间, null, true);
            AddNumericTag("累计.块矿面.球团二车间", "昨日累计.块矿面.球团二车间+本日.块矿面.球团二车间", dec_球团二车间, null, true);
            AddNumericTag("累计.1高炉.球团二车间", "昨日累计.1高炉.球团二车间+本日.1高炉.球团二车间", dec_球团二车间, null, true);
            AddNumericTag("累计.4高炉.球团二车间", "昨日累计.4高炉.球团二车间+本日.4高炉.球团二车间", dec_球团二车间, null, true);
            AddNumericTag("累计.5高炉.球团二车间", "昨日累计.5高炉.球团二车间+本日.5高炉.球团二车间", dec_球团二车间, null, true);
            AddNumericTag("累计.6高炉.球团二车间", "昨日累计.6高炉.球团二车间+本日.6高炉.球团二车间", dec_球团二车间, null, true);
            AddNumericTag("累计.7高炉.球团二车间", "昨日累计.7高炉.球团二车间+本日.7高炉.球团二车间", dec_球团二车间, null, true);
            AddNumericTag("累计.8高炉.球团二车间", "昨日累计.8高炉.球团二车间+本日.8高炉.球团二车间", dec_球团二车间, null, true);
            AddNumericTag("累计.扣返矿.球团二车间", "昨日累计.扣返矿.球团二车间-本日.扣返矿.球团二车间", dec_球团二车间, null, true);
            AddNumericTag("累计.球团二车间", "累计.1高炉.球团二车间+累计.4高炉.球团二车间+累计.5高炉.球团二车间+累计.6高炉.球团二车间+累计.7高炉.球团二车间+累计.8高炉.球团二车间-累计.扣返矿.球团二车间", dec_球团二车间, null, false);






            AddNumericTag("昨日累计.供应.球团二车间", "", dec_球团二车间, null, false);
            AddNumericTag("昨日累计.块矿面.球团二车间", "", dec_球团二车间, null, false);
            AddNumericTag("昨日累计.1高炉.球团二车间", "", dec_球团二车间, null, false);
            AddNumericTag("昨日累计.4高炉.球团二车间", "", dec_球团二车间, null, false);
            AddNumericTag("昨日累计.5高炉.球团二车间", "", dec_球团二车间, null, false);
            AddNumericTag("昨日累计.6高炉.球团二车间", "", dec_球团二车间, null, false);
            AddNumericTag("昨日累计.7高炉.球团二车间", "", dec_球团二车间, null, false);
            AddNumericTag("昨日累计.8高炉.球团二车间", "", dec_球团二车间, null, false);
            AddNumericTag("昨日累计.扣返矿.球团二车间", "", dec_球团二车间, null, false);




            #endregion

            #region 北区球团
            int? dec_北区球团 = 2;
            AddNumericTag("本日.供应.北区球团", "", dec_北区球团, null, true);
            AddNumericTag("本日.块矿面.北区球团", "", dec_北区球团, null, true);
            AddNumericTag("本日.1高炉.北区球团", "", dec_北区球团, null, true);
            AddNumericTag("本日.4高炉.北区球团", "", dec_北区球团, null, true);
            AddNumericTag("本日.5高炉.北区球团", "", dec_北区球团, null, true);
            AddNumericTag("本日.6高炉.北区球团", "", dec_北区球团, null, true);
            AddNumericTag("本日.7高炉.北区球团", "", dec_北区球团, null, true);
            AddNumericTag("本日.8高炉.北区球团", "", dec_北区球团, null, true);
            AddNumericTag("本日.扣返矿.北区球团", "", dec_北区球团, null, true);
            AddNumericTag("本日.北区球团", "本日.1高炉.北区球团+本日.4高炉.北区球团+本日.5高炉.北区球团+本日.6高炉.北区球团+本日.7高炉.北区球团+本日.8高炉.北区球团-本日.扣返矿.北区球团", dec_北区球团, null, true);

            AddNumericTag("本日库存.1高炉.北区球团", "", dec_北区球团, null, true);
            AddNumericTag("本日库存.二车间.北区球团", "", dec_北区球团, null, true);
            AddNumericTag("本日库存.三车间.北区球团", "", dec_北区球团, null, true);
            AddNumericTag("本日库存.北区球团", "本日库存.1高炉.北区球团+本日库存.二车间.北区球团+本日库存.三车间.北区球团", dec_北区球团, null, true);

            AddNumericTag("累计.供应.北区球团", "昨日累计.供应.北区球团+本日.供应.北区球团", dec_北区球团, null, true);
            AddNumericTag("累计.块矿面.北区球团", "昨日累计.块矿面.北区球团+本日.块矿面.北区球团", dec_北区球团, null, true);
            AddNumericTag("累计.1高炉.北区球团", "昨日累计.1高炉.北区球团+本日.1高炉.北区球团", dec_北区球团, null, true);
            AddNumericTag("累计.4高炉.北区球团", "昨日累计.4高炉.北区球团+本日.4高炉.北区球团", dec_北区球团, null, true);
            AddNumericTag("累计.5高炉.北区球团", "昨日累计.5高炉.北区球团+本日.5高炉.北区球团", dec_北区球团, null, true);
            AddNumericTag("累计.6高炉.北区球团", "昨日累计.6高炉.北区球团+本日.6高炉.北区球团", dec_北区球团, null, true);
            AddNumericTag("累计.7高炉.北区球团", "昨日累计.7高炉.北区球团+本日.7高炉.北区球团", dec_北区球团, null, true);
            AddNumericTag("累计.8高炉.北区球团", "昨日累计.8高炉.北区球团+本日.8高炉.北区球团", dec_北区球团, null, true);
            AddNumericTag("累计.扣返矿.北区球团", "昨日累计.扣返矿.北区球团-本日.扣返矿.北区球团", dec_北区球团, null, true);
            AddNumericTag("累计.北区球团", "累计.1高炉.北区球团+累计.4高炉.北区球团+累计.5高炉.北区球团+累计.6高炉.北区球团+累计.7高炉.北区球团+累计.8高炉.北区球团-累计.扣返矿.北区球团", dec_北区球团, null, false);






            AddNumericTag("昨日累计.供应.北区球团", "", dec_北区球团, null, false);
            AddNumericTag("昨日累计.块矿面.北区球团", "", dec_北区球团, null, false);
            AddNumericTag("昨日累计.1高炉.北区球团", "", dec_北区球团, null, false);
            AddNumericTag("昨日累计.4高炉.北区球团", "", dec_北区球团, null, false);
            AddNumericTag("昨日累计.5高炉.北区球团", "", dec_北区球团, null, false);
            AddNumericTag("昨日累计.6高炉.北区球团", "", dec_北区球团, null, false);
            AddNumericTag("昨日累计.7高炉.北区球团", "", dec_北区球团, null, false);
            AddNumericTag("昨日累计.8高炉.北区球团", "", dec_北区球团, null, false);
            AddNumericTag("昨日累计.扣返矿.北区球团", "", dec_北区球团, null, false);




            #endregion

            #region 纽曼块
            int? dec_纽曼块 = 2;
            AddNumericTag("本日.供应.纽曼块", "", dec_纽曼块, null, true);
            AddNumericTag("本日.块矿面.纽曼块", "", dec_纽曼块, null, true);
            AddNumericTag("本日.1高炉.纽曼块", "", dec_纽曼块, null, true);
            AddNumericTag("本日.4高炉.纽曼块", "", dec_纽曼块, null, true);
            AddNumericTag("本日.5高炉.纽曼块", "", dec_纽曼块, null, true);
            AddNumericTag("本日.6高炉.纽曼块", "", dec_纽曼块, null, true);
            AddNumericTag("本日.7高炉.纽曼块", "", dec_纽曼块, null, true);
            AddNumericTag("本日.8高炉.纽曼块", "", dec_纽曼块, null, true);
            AddNumericTag("本日.扣返矿.纽曼块", "", dec_纽曼块, null, true);
            AddNumericTag("本日.纽曼块", "本日.1高炉.纽曼块+本日.4高炉.纽曼块+本日.5高炉.纽曼块+本日.6高炉.纽曼块+本日.7高炉.纽曼块+本日.8高炉.纽曼块-本日.扣返矿.纽曼块", dec_纽曼块, null, true);

            AddNumericTag("本日库存.1高炉.纽曼块", "", dec_纽曼块, null, true);
            AddNumericTag("本日库存.二车间.纽曼块", "", dec_纽曼块, null, true);
            AddNumericTag("本日库存.三车间.纽曼块", "", dec_纽曼块, null, true);
            AddNumericTag("本日库存.纽曼块", "本日库存.1高炉.纽曼块+本日库存.二车间.纽曼块+本日库存.三车间.纽曼块", dec_纽曼块, null, true);

            AddNumericTag("累计.供应.纽曼块", "昨日累计.供应.纽曼块+本日.供应.纽曼块", dec_纽曼块, null, true);
            AddNumericTag("累计.块矿面.纽曼块", "昨日累计.块矿面.纽曼块+本日.块矿面.纽曼块", dec_纽曼块, null, true);
            AddNumericTag("累计.1高炉.纽曼块", "昨日累计.1高炉.纽曼块+本日.1高炉.纽曼块", dec_纽曼块, null, true);
            AddNumericTag("累计.4高炉.纽曼块", "昨日累计.4高炉.纽曼块+本日.4高炉.纽曼块", dec_纽曼块, null, true);
            AddNumericTag("累计.5高炉.纽曼块", "昨日累计.5高炉.纽曼块+本日.5高炉.纽曼块", dec_纽曼块, null, true);
            AddNumericTag("累计.6高炉.纽曼块", "昨日累计.6高炉.纽曼块+本日.6高炉.纽曼块", dec_纽曼块, null, true);
            AddNumericTag("累计.7高炉.纽曼块", "昨日累计.7高炉.纽曼块+本日.7高炉.纽曼块", dec_纽曼块, null, true);
            AddNumericTag("累计.8高炉.纽曼块", "昨日累计.8高炉.纽曼块+本日.8高炉.纽曼块", dec_纽曼块, null, true);
            AddNumericTag("累计.扣返矿.纽曼块", "昨日累计.扣返矿.纽曼块-本日.扣返矿.纽曼块", dec_纽曼块, null, true);
            AddNumericTag("累计.纽曼块", "累计.1高炉.纽曼块+累计.4高炉.纽曼块+累计.5高炉.纽曼块+累计.6高炉.纽曼块+累计.7高炉.纽曼块+累计.8高炉.纽曼块-累计.扣返矿.纽曼块", dec_纽曼块, null, false);






            AddNumericTag("昨日累计.供应.纽曼块", "", dec_纽曼块, null, false);
            AddNumericTag("昨日累计.块矿面.纽曼块", "", dec_纽曼块, null, false);
            AddNumericTag("昨日累计.1高炉.纽曼块", "", dec_纽曼块, null, false);
            AddNumericTag("昨日累计.4高炉.纽曼块", "", dec_纽曼块, null, false);
            AddNumericTag("昨日累计.5高炉.纽曼块", "", dec_纽曼块, null, false);
            AddNumericTag("昨日累计.6高炉.纽曼块", "", dec_纽曼块, null, false);
            AddNumericTag("昨日累计.7高炉.纽曼块", "", dec_纽曼块, null, false);
            AddNumericTag("昨日累计.8高炉.纽曼块", "", dec_纽曼块, null, false);
            AddNumericTag("昨日累计.扣返矿.纽曼块", "", dec_纽曼块, null, false);




            #endregion

            #region 澳块
            int? dec_澳块 = 2;
            AddNumericTag("本日.供应.澳块", "", dec_澳块, null, true);
            AddNumericTag("本日.块矿面.澳块", "", dec_澳块, null, true);
            AddNumericTag("本日.1高炉.澳块", "", dec_澳块, null, true);
            AddNumericTag("本日.4高炉.澳块", "", dec_澳块, null, true);
            AddNumericTag("本日.5高炉.澳块", "", dec_澳块, null, true);
            AddNumericTag("本日.6高炉.澳块", "", dec_澳块, null, true);
            AddNumericTag("本日.7高炉.澳块", "", dec_澳块, null, true);
            AddNumericTag("本日.8高炉.澳块", "", dec_澳块, null, true);
            AddNumericTag("本日.扣返矿.澳块", "", dec_澳块, null, true);
            AddNumericTag("本日.澳块", "本日.1高炉.澳块+本日.4高炉.澳块+本日.5高炉.澳块+本日.6高炉.澳块+本日.7高炉.澳块+本日.8高炉.澳块-本日.扣返矿.澳块", dec_澳块, null, true);

            AddNumericTag("本日库存.1高炉.澳块", "", dec_澳块, null, true);
            AddNumericTag("本日库存.二车间.澳块", "", dec_澳块, null, true);
            AddNumericTag("本日库存.三车间.澳块", "", dec_澳块, null, true);
            AddNumericTag("本日库存.澳块", "本日库存.1高炉.澳块+本日库存.二车间.澳块+本日库存.三车间.澳块", dec_澳块, null, true);

            AddNumericTag("累计.供应.澳块", "昨日累计.供应.澳块+本日.供应.澳块", dec_澳块, null, true);
            AddNumericTag("累计.块矿面.澳块", "昨日累计.块矿面.澳块+本日.块矿面.澳块", dec_澳块, null, true);
            AddNumericTag("累计.1高炉.澳块", "昨日累计.1高炉.澳块+本日.1高炉.澳块", dec_澳块, null, true);
            AddNumericTag("累计.4高炉.澳块", "昨日累计.4高炉.澳块+本日.4高炉.澳块", dec_澳块, null, true);
            AddNumericTag("累计.5高炉.澳块", "昨日累计.5高炉.澳块+本日.5高炉.澳块", dec_澳块, null, true);
            AddNumericTag("累计.6高炉.澳块", "昨日累计.6高炉.澳块+本日.6高炉.澳块", dec_澳块, null, true);
            AddNumericTag("累计.7高炉.澳块", "昨日累计.7高炉.澳块+本日.7高炉.澳块", dec_澳块, null, true);
            AddNumericTag("累计.8高炉.澳块", "昨日累计.8高炉.澳块+本日.8高炉.澳块", dec_澳块, null, true);
            AddNumericTag("累计.扣返矿.澳块", "昨日累计.扣返矿.澳块-本日.扣返矿.澳块", dec_澳块, null, true);
            AddNumericTag("累计.澳块", "累计.1高炉.澳块+累计.4高炉.澳块+累计.5高炉.澳块+累计.6高炉.澳块+累计.7高炉.澳块+累计.8高炉.澳块-累计.扣返矿.澳块", dec_澳块, null, false);






            AddNumericTag("昨日累计.供应.澳块", "", dec_澳块, null, false);
            AddNumericTag("昨日累计.块矿面.澳块", "", dec_澳块, null, false);
            AddNumericTag("昨日累计.1高炉.澳块", "", dec_澳块, null, false);
            AddNumericTag("昨日累计.4高炉.澳块", "", dec_澳块, null, false);
            AddNumericTag("昨日累计.5高炉.澳块", "", dec_澳块, null, false);
            AddNumericTag("昨日累计.6高炉.澳块", "", dec_澳块, null, false);
            AddNumericTag("昨日累计.7高炉.澳块", "", dec_澳块, null, false);
            AddNumericTag("昨日累计.8高炉.澳块", "", dec_澳块, null, false);
            AddNumericTag("昨日累计.扣返矿.澳块", "", dec_澳块, null, false);




            #endregion

            #region 伊朗块
            int? dec_伊朗块 = 2;
            AddNumericTag("本日.供应.伊朗块", "", dec_伊朗块, null, true);
            AddNumericTag("本日.块矿面.伊朗块", "", dec_伊朗块, null, true);
            AddNumericTag("本日.1高炉.伊朗块", "", dec_伊朗块, null, true);
            AddNumericTag("本日.4高炉.伊朗块", "", dec_伊朗块, null, true);
            AddNumericTag("本日.5高炉.伊朗块", "", dec_伊朗块, null, true);
            AddNumericTag("本日.6高炉.伊朗块", "", dec_伊朗块, null, true);
            AddNumericTag("本日.7高炉.伊朗块", "", dec_伊朗块, null, true);
            AddNumericTag("本日.8高炉.伊朗块", "", dec_伊朗块, null, true);
            AddNumericTag("本日.扣返矿.伊朗块", "", dec_伊朗块, null, true);
            AddNumericTag("本日.伊朗块", "本日.1高炉.伊朗块+本日.4高炉.伊朗块+本日.5高炉.伊朗块+本日.6高炉.伊朗块+本日.7高炉.伊朗块+本日.8高炉.伊朗块-本日.扣返矿.伊朗块", dec_伊朗块, null, true);

            AddNumericTag("本日库存.1高炉.伊朗块", "", dec_伊朗块, null, true);
            AddNumericTag("本日库存.二车间.伊朗块", "", dec_伊朗块, null, true);
            AddNumericTag("本日库存.三车间.伊朗块", "", dec_伊朗块, null, true);
            AddNumericTag("本日库存.伊朗块", "本日库存.1高炉.伊朗块+本日库存.二车间.伊朗块+本日库存.三车间.伊朗块", dec_伊朗块, null, true);

            AddNumericTag("累计.供应.伊朗块", "昨日累计.供应.伊朗块+本日.供应.伊朗块", dec_伊朗块, null, true);
            AddNumericTag("累计.块矿面.伊朗块", "昨日累计.块矿面.伊朗块+本日.块矿面.伊朗块", dec_伊朗块, null, true);
            AddNumericTag("累计.1高炉.伊朗块", "昨日累计.1高炉.伊朗块+本日.1高炉.伊朗块", dec_伊朗块, null, true);
            AddNumericTag("累计.4高炉.伊朗块", "昨日累计.4高炉.伊朗块+本日.4高炉.伊朗块", dec_伊朗块, null, true);
            AddNumericTag("累计.5高炉.伊朗块", "昨日累计.5高炉.伊朗块+本日.5高炉.伊朗块", dec_伊朗块, null, true);
            AddNumericTag("累计.6高炉.伊朗块", "昨日累计.6高炉.伊朗块+本日.6高炉.伊朗块", dec_伊朗块, null, true);
            AddNumericTag("累计.7高炉.伊朗块", "昨日累计.7高炉.伊朗块+本日.7高炉.伊朗块", dec_伊朗块, null, true);
            AddNumericTag("累计.8高炉.伊朗块", "昨日累计.8高炉.伊朗块+本日.8高炉.伊朗块", dec_伊朗块, null, true);
            AddNumericTag("累计.扣返矿.伊朗块", "昨日累计.扣返矿.伊朗块-本日.扣返矿.伊朗块", dec_伊朗块, null, true);
            AddNumericTag("累计.伊朗块", "累计.1高炉.伊朗块+累计.4高炉.伊朗块+累计.5高炉.伊朗块+累计.6高炉.伊朗块+累计.7高炉.伊朗块+累计.8高炉.伊朗块-累计.扣返矿.伊朗块", dec_伊朗块, null, false);






            AddNumericTag("昨日累计.供应.伊朗块", "", dec_伊朗块, null, false);
            AddNumericTag("昨日累计.块矿面.伊朗块", "", dec_伊朗块, null, false);
            AddNumericTag("昨日累计.1高炉.伊朗块", "", dec_伊朗块, null, false);
            AddNumericTag("昨日累计.4高炉.伊朗块", "", dec_伊朗块, null, false);
            AddNumericTag("昨日累计.5高炉.伊朗块", "", dec_伊朗块, null, false);
            AddNumericTag("昨日累计.6高炉.伊朗块", "", dec_伊朗块, null, false);
            AddNumericTag("昨日累计.7高炉.伊朗块", "", dec_伊朗块, null, false);
            AddNumericTag("昨日累计.8高炉.伊朗块", "", dec_伊朗块, null, false);
            AddNumericTag("昨日累计.扣返矿.伊朗块", "", dec_伊朗块, null, false);




            #endregion

            #region 巴西块
            int? dec_巴西块 = 2;
            AddNumericTag("本日.供应.巴西块", "", dec_巴西块, null, true);
            AddNumericTag("本日.块矿面.巴西块", "", dec_巴西块, null, true);
            AddNumericTag("本日.1高炉.巴西块", "", dec_巴西块, null, true);
            AddNumericTag("本日.4高炉.巴西块", "", dec_巴西块, null, true);
            AddNumericTag("本日.5高炉.巴西块", "", dec_巴西块, null, true);
            AddNumericTag("本日.6高炉.巴西块", "", dec_巴西块, null, true);
            AddNumericTag("本日.7高炉.巴西块", "", dec_巴西块, null, true);
            AddNumericTag("本日.8高炉.巴西块", "", dec_巴西块, null, true);
            AddNumericTag("本日.扣返矿.巴西块", "", dec_巴西块, null, true);
            AddNumericTag("本日.巴西块", "本日.1高炉.巴西块+本日.4高炉.巴西块+本日.5高炉.巴西块+本日.6高炉.巴西块+本日.7高炉.巴西块+本日.8高炉.巴西块-本日.扣返矿.巴西块", dec_巴西块, null, true);

            AddNumericTag("本日库存.1高炉.巴西块", "", dec_巴西块, null, true);
            AddNumericTag("本日库存.二车间.巴西块", "", dec_巴西块, null, true);
            AddNumericTag("本日库存.三车间.巴西块", "", dec_巴西块, null, true);
            AddNumericTag("本日库存.巴西块", "本日库存.1高炉.巴西块+本日库存.二车间.巴西块+本日库存.三车间.巴西块", dec_巴西块, null, true);

            AddNumericTag("累计.供应.巴西块", "昨日累计.供应.巴西块+本日.供应.巴西块", dec_巴西块, null, true);
            AddNumericTag("累计.块矿面.巴西块", "昨日累计.块矿面.巴西块+本日.块矿面.巴西块", dec_巴西块, null, true);
            AddNumericTag("累计.1高炉.巴西块", "昨日累计.1高炉.巴西块+本日.1高炉.巴西块", dec_巴西块, null, true);
            AddNumericTag("累计.4高炉.巴西块", "昨日累计.4高炉.巴西块+本日.4高炉.巴西块", dec_巴西块, null, true);
            AddNumericTag("累计.5高炉.巴西块", "昨日累计.5高炉.巴西块+本日.5高炉.巴西块", dec_巴西块, null, true);
            AddNumericTag("累计.6高炉.巴西块", "昨日累计.6高炉.巴西块+本日.6高炉.巴西块", dec_巴西块, null, true);
            AddNumericTag("累计.7高炉.巴西块", "昨日累计.7高炉.巴西块+本日.7高炉.巴西块", dec_巴西块, null, true);
            AddNumericTag("累计.8高炉.巴西块", "昨日累计.8高炉.巴西块+本日.8高炉.巴西块", dec_巴西块, null, true);
            AddNumericTag("累计.扣返矿.巴西块", "昨日累计.扣返矿.巴西块-本日.扣返矿.巴西块", dec_巴西块, null, true);
            AddNumericTag("累计.巴西块", "累计.1高炉.巴西块+累计.4高炉.巴西块+累计.5高炉.巴西块+累计.6高炉.巴西块+累计.7高炉.巴西块+累计.8高炉.巴西块-累计.扣返矿.巴西块", dec_巴西块, null, false);






            AddNumericTag("昨日累计.供应.巴西块", "", dec_巴西块, null, false);
            AddNumericTag("昨日累计.块矿面.巴西块", "", dec_巴西块, null, false);
            AddNumericTag("昨日累计.1高炉.巴西块", "", dec_巴西块, null, false);
            AddNumericTag("昨日累计.4高炉.巴西块", "", dec_巴西块, null, false);
            AddNumericTag("昨日累计.5高炉.巴西块", "", dec_巴西块, null, false);
            AddNumericTag("昨日累计.6高炉.巴西块", "", dec_巴西块, null, false);
            AddNumericTag("昨日累计.7高炉.巴西块", "", dec_巴西块, null, false);
            AddNumericTag("昨日累计.8高炉.巴西块", "", dec_巴西块, null, false);
            AddNumericTag("昨日累计.扣返矿.巴西块", "", dec_巴西块, null, false);




            #endregion

            #region 铁矿石
            int? dec_铁矿石 = 2;
            AddNumericTag("本日.供应.铁矿石", "", dec_铁矿石, null, true);
            AddNumericTag("本日.块矿面.铁矿石", "", dec_铁矿石, null, true);
            AddNumericTag("本日.1高炉.铁矿石", "", dec_铁矿石, null, true);
            AddNumericTag("本日.4高炉.铁矿石", "", dec_铁矿石, null, true);
            AddNumericTag("本日.5高炉.铁矿石", "", dec_铁矿石, null, true);
            AddNumericTag("本日.6高炉.铁矿石", "", dec_铁矿石, null, true);
            AddNumericTag("本日.7高炉.铁矿石", "", dec_铁矿石, null, true);
            AddNumericTag("本日.8高炉.铁矿石", "", dec_铁矿石, null, true);
            AddNumericTag("本日.扣返矿.铁矿石", "", dec_铁矿石, null, true);
            AddNumericTag("本日.铁矿石", "本日.1高炉.铁矿石+本日.4高炉.铁矿石+本日.5高炉.铁矿石+本日.6高炉.铁矿石+本日.7高炉.铁矿石+本日.8高炉.铁矿石-本日.扣返矿.铁矿石", dec_铁矿石, null, true);

            AddNumericTag("本日库存.1高炉.铁矿石", "", dec_铁矿石, null, true);
            AddNumericTag("本日库存.二车间.铁矿石", "", dec_铁矿石, null, true);
            AddNumericTag("本日库存.三车间.铁矿石", "", dec_铁矿石, null, true);
            AddNumericTag("本日库存.铁矿石", "本日库存.1高炉.铁矿石+本日库存.二车间.铁矿石+本日库存.三车间.铁矿石", dec_铁矿石, null, true);

            AddNumericTag("累计.供应.铁矿石", "昨日累计.供应.铁矿石+本日.供应.铁矿石", dec_铁矿石, null, true);
            AddNumericTag("累计.块矿面.铁矿石", "昨日累计.块矿面.铁矿石+本日.块矿面.铁矿石", dec_铁矿石, null, true);
            AddNumericTag("累计.1高炉.铁矿石", "昨日累计.1高炉.铁矿石+本日.1高炉.铁矿石", dec_铁矿石, null, true);
            AddNumericTag("累计.4高炉.铁矿石", "昨日累计.4高炉.铁矿石+本日.4高炉.铁矿石", dec_铁矿石, null, true);
            AddNumericTag("累计.5高炉.铁矿石", "昨日累计.5高炉.铁矿石+本日.5高炉.铁矿石", dec_铁矿石, null, true);
            AddNumericTag("累计.6高炉.铁矿石", "昨日累计.6高炉.铁矿石+本日.6高炉.铁矿石", dec_铁矿石, null, true);
            AddNumericTag("累计.7高炉.铁矿石", "昨日累计.7高炉.铁矿石+本日.7高炉.铁矿石", dec_铁矿石, null, true);
            AddNumericTag("累计.8高炉.铁矿石", "昨日累计.8高炉.铁矿石+本日.8高炉.铁矿石", dec_铁矿石, null, true);
            AddNumericTag("累计.扣返矿.铁矿石", "昨日累计.扣返矿.铁矿石-本日.扣返矿.铁矿石", dec_铁矿石, null, true);
            AddNumericTag("累计.铁矿石", "累计.1高炉.铁矿石+累计.4高炉.铁矿石+累计.5高炉.铁矿石+累计.6高炉.铁矿石+累计.7高炉.铁矿石+累计.8高炉.铁矿石-累计.扣返矿.铁矿石", dec_铁矿石, null, false);






            AddNumericTag("昨日累计.供应.铁矿石", "", dec_铁矿石, null, false);
            AddNumericTag("昨日累计.块矿面.铁矿石", "", dec_铁矿石, null, false);
            AddNumericTag("昨日累计.1高炉.铁矿石", "", dec_铁矿石, null, false);
            AddNumericTag("昨日累计.4高炉.铁矿石", "", dec_铁矿石, null, false);
            AddNumericTag("昨日累计.5高炉.铁矿石", "", dec_铁矿石, null, false);
            AddNumericTag("昨日累计.6高炉.铁矿石", "", dec_铁矿石, null, false);
            AddNumericTag("昨日累计.7高炉.铁矿石", "", dec_铁矿石, null, false);
            AddNumericTag("昨日累计.8高炉.铁矿石", "", dec_铁矿石, null, false);
            AddNumericTag("昨日累计.扣返矿.铁矿石", "", dec_铁矿石, null, false);




            #endregion

            #region 赛拉利昂块
            int? dec_赛拉利昂块 = 2;
            AddNumericTag("本日.供应.赛拉利昂块", "", dec_赛拉利昂块, null, true);
            AddNumericTag("本日.块矿面.赛拉利昂块", "", dec_赛拉利昂块, null, true);
            AddNumericTag("本日.1高炉.赛拉利昂块", "", dec_赛拉利昂块, null, true);
            AddNumericTag("本日.4高炉.赛拉利昂块", "", dec_赛拉利昂块, null, true);
            AddNumericTag("本日.5高炉.赛拉利昂块", "", dec_赛拉利昂块, null, true);
            AddNumericTag("本日.6高炉.赛拉利昂块", "", dec_赛拉利昂块, null, true);
            AddNumericTag("本日.7高炉.赛拉利昂块", "", dec_赛拉利昂块, null, true);
            AddNumericTag("本日.8高炉.赛拉利昂块", "", dec_赛拉利昂块, null, true);
            AddNumericTag("本日.扣返矿.赛拉利昂块", "", dec_赛拉利昂块, null, true);
            AddNumericTag("本日.赛拉利昂块", "本日.1高炉.赛拉利昂块+本日.4高炉.赛拉利昂块+本日.5高炉.赛拉利昂块+本日.6高炉.赛拉利昂块+本日.7高炉.赛拉利昂块+本日.8高炉.赛拉利昂块-本日.扣返矿.赛拉利昂块", dec_赛拉利昂块, null, true);

            AddNumericTag("本日库存.1高炉.赛拉利昂块", "", dec_赛拉利昂块, null, true);
            AddNumericTag("本日库存.二车间.赛拉利昂块", "", dec_赛拉利昂块, null, true);
            AddNumericTag("本日库存.三车间.赛拉利昂块", "", dec_赛拉利昂块, null, true);
            AddNumericTag("本日库存.赛拉利昂块", "本日库存.1高炉.赛拉利昂块+本日库存.二车间.赛拉利昂块+本日库存.三车间.赛拉利昂块", dec_赛拉利昂块, null, true);

            AddNumericTag("累计.供应.赛拉利昂块", "昨日累计.供应.赛拉利昂块+本日.供应.赛拉利昂块", dec_赛拉利昂块, null, true);
            AddNumericTag("累计.块矿面.赛拉利昂块", "昨日累计.块矿面.赛拉利昂块+本日.块矿面.赛拉利昂块", dec_赛拉利昂块, null, true);
            AddNumericTag("累计.1高炉.赛拉利昂块", "昨日累计.1高炉.赛拉利昂块+本日.1高炉.赛拉利昂块", dec_赛拉利昂块, null, true);
            AddNumericTag("累计.4高炉.赛拉利昂块", "昨日累计.4高炉.赛拉利昂块+本日.4高炉.赛拉利昂块", dec_赛拉利昂块, null, true);
            AddNumericTag("累计.5高炉.赛拉利昂块", "昨日累计.5高炉.赛拉利昂块+本日.5高炉.赛拉利昂块", dec_赛拉利昂块, null, true);
            AddNumericTag("累计.6高炉.赛拉利昂块", "昨日累计.6高炉.赛拉利昂块+本日.6高炉.赛拉利昂块", dec_赛拉利昂块, null, true);
            AddNumericTag("累计.7高炉.赛拉利昂块", "昨日累计.7高炉.赛拉利昂块+本日.7高炉.赛拉利昂块", dec_赛拉利昂块, null, true);
            AddNumericTag("累计.8高炉.赛拉利昂块", "昨日累计.8高炉.赛拉利昂块+本日.8高炉.赛拉利昂块", dec_赛拉利昂块, null, true);
            AddNumericTag("累计.扣返矿.赛拉利昂块", "昨日累计.扣返矿.赛拉利昂块-本日.扣返矿.赛拉利昂块", dec_赛拉利昂块, null, true);
            AddNumericTag("累计.赛拉利昂块", "累计.1高炉.赛拉利昂块+累计.4高炉.赛拉利昂块+累计.5高炉.赛拉利昂块+累计.6高炉.赛拉利昂块+累计.7高炉.赛拉利昂块+累计.8高炉.赛拉利昂块-累计.扣返矿.赛拉利昂块", dec_赛拉利昂块, null, false);






            AddNumericTag("昨日累计.供应.赛拉利昂块", "", dec_赛拉利昂块, null, false);
            AddNumericTag("昨日累计.块矿面.赛拉利昂块", "", dec_赛拉利昂块, null, false);
            AddNumericTag("昨日累计.1高炉.赛拉利昂块", "", dec_赛拉利昂块, null, false);
            AddNumericTag("昨日累计.4高炉.赛拉利昂块", "", dec_赛拉利昂块, null, false);
            AddNumericTag("昨日累计.5高炉.赛拉利昂块", "", dec_赛拉利昂块, null, false);
            AddNumericTag("昨日累计.6高炉.赛拉利昂块", "", dec_赛拉利昂块, null, false);
            AddNumericTag("昨日累计.7高炉.赛拉利昂块", "", dec_赛拉利昂块, null, false);
            AddNumericTag("昨日累计.8高炉.赛拉利昂块", "", dec_赛拉利昂块, null, false);
            AddNumericTag("昨日累计.扣返矿.赛拉利昂块", "", dec_赛拉利昂块, null, false);




            #endregion

            #region 钢粒
            int? dec_钢粒 = 2;
            AddNumericTag("本日.供应.钢粒", "", dec_钢粒, null, true);
            AddNumericTag("本日.块矿面.钢粒", "", dec_钢粒, null, true);
            AddNumericTag("本日.1高炉.钢粒", "", dec_钢粒, null, true);
            AddNumericTag("本日.4高炉.钢粒", "", dec_钢粒, null, true);
            AddNumericTag("本日.5高炉.钢粒", "", dec_钢粒, null, true);
            AddNumericTag("本日.6高炉.钢粒", "", dec_钢粒, null, true);
            AddNumericTag("本日.7高炉.钢粒", "", dec_钢粒, null, true);
            AddNumericTag("本日.8高炉.钢粒", "", dec_钢粒, null, true);
            AddNumericTag("本日.扣返矿.钢粒", "", dec_钢粒, null, true);
            AddNumericTag("本日.钢粒", "本日.1高炉.钢粒+本日.4高炉.钢粒+本日.5高炉.钢粒+本日.6高炉.钢粒+本日.7高炉.钢粒+本日.8高炉.钢粒-本日.扣返矿.钢粒", dec_钢粒, null, true);

            AddNumericTag("本日库存.1高炉.钢粒", "", dec_钢粒, null, true);
            AddNumericTag("本日库存.二车间.钢粒", "", dec_钢粒, null, true);
            AddNumericTag("本日库存.三车间.钢粒", "", dec_钢粒, null, true);
            AddNumericTag("本日库存.钢粒", "本日库存.1高炉.钢粒+本日库存.二车间.钢粒+本日库存.三车间.钢粒", dec_钢粒, null, true);

            AddNumericTag("累计.供应.钢粒", "昨日累计.供应.钢粒+本日.供应.钢粒", dec_钢粒, null, true);
            AddNumericTag("累计.块矿面.钢粒", "昨日累计.块矿面.钢粒+本日.块矿面.钢粒", dec_钢粒, null, true);
            AddNumericTag("累计.1高炉.钢粒", "昨日累计.1高炉.钢粒+本日.1高炉.钢粒", dec_钢粒, null, true);
            AddNumericTag("累计.4高炉.钢粒", "昨日累计.4高炉.钢粒+本日.4高炉.钢粒", dec_钢粒, null, true);
            AddNumericTag("累计.5高炉.钢粒", "昨日累计.5高炉.钢粒+本日.5高炉.钢粒", dec_钢粒, null, true);
            AddNumericTag("累计.6高炉.钢粒", "昨日累计.6高炉.钢粒+本日.6高炉.钢粒", dec_钢粒, null, true);
            AddNumericTag("累计.7高炉.钢粒", "昨日累计.7高炉.钢粒+本日.7高炉.钢粒", dec_钢粒, null, true);
            AddNumericTag("累计.8高炉.钢粒", "昨日累计.8高炉.钢粒+本日.8高炉.钢粒", dec_钢粒, null, true);
            AddNumericTag("累计.扣返矿.钢粒", "昨日累计.扣返矿.钢粒-本日.扣返矿.钢粒", dec_钢粒, null, true);
            AddNumericTag("累计.钢粒", "累计.1高炉.钢粒+累计.4高炉.钢粒+累计.5高炉.钢粒+累计.6高炉.钢粒+累计.7高炉.钢粒+累计.8高炉.钢粒-累计.扣返矿.钢粒", dec_钢粒, null, false);






            AddNumericTag("昨日累计.供应.钢粒", "", dec_钢粒, null, false);
            AddNumericTag("昨日累计.块矿面.钢粒", "", dec_钢粒, null, false);
            AddNumericTag("昨日累计.1高炉.钢粒", "", dec_钢粒, null, false);
            AddNumericTag("昨日累计.4高炉.钢粒", "", dec_钢粒, null, false);
            AddNumericTag("昨日累计.5高炉.钢粒", "", dec_钢粒, null, false);
            AddNumericTag("昨日累计.6高炉.钢粒", "", dec_钢粒, null, false);
            AddNumericTag("昨日累计.7高炉.钢粒", "", dec_钢粒, null, false);
            AddNumericTag("昨日累计.8高炉.钢粒", "", dec_钢粒, null, false);
            AddNumericTag("昨日累计.扣返矿.钢粒", "", dec_钢粒, null, false);




            #endregion

            #region 炉渣
            int? dec_炉渣 = 2;
            AddNumericTag("本日.供应.炉渣", "", dec_炉渣, null, true);
            AddNumericTag("本日.块矿面.炉渣", "", dec_炉渣, null, true);
            AddNumericTag("本日.1高炉.炉渣", "", dec_炉渣, null, true);
            AddNumericTag("本日.4高炉.炉渣", "", dec_炉渣, null, true);
            AddNumericTag("本日.5高炉.炉渣", "", dec_炉渣, null, true);
            AddNumericTag("本日.6高炉.炉渣", "", dec_炉渣, null, true);
            AddNumericTag("本日.7高炉.炉渣", "", dec_炉渣, null, true);
            AddNumericTag("本日.8高炉.炉渣", "", dec_炉渣, null, true);
            AddNumericTag("本日.扣返矿.炉渣", "", dec_炉渣, null, true);
            AddNumericTag("本日.炉渣", "本日.1高炉.炉渣+本日.4高炉.炉渣+本日.5高炉.炉渣+本日.6高炉.炉渣+本日.7高炉.炉渣+本日.8高炉.炉渣-本日.扣返矿.炉渣", dec_炉渣, null, true);

            AddNumericTag("本日库存.1高炉.炉渣", "", dec_炉渣, null, true);
            AddNumericTag("本日库存.二车间.炉渣", "", dec_炉渣, null, true);
            AddNumericTag("本日库存.三车间.炉渣", "", dec_炉渣, null, true);
            AddNumericTag("本日库存.炉渣", "本日库存.1高炉.炉渣+本日库存.二车间.炉渣+本日库存.三车间.炉渣", dec_炉渣, null, true);

            AddNumericTag("累计.供应.炉渣", "昨日累计.供应.炉渣+本日.供应.炉渣", dec_炉渣, null, true);
            AddNumericTag("累计.块矿面.炉渣", "昨日累计.块矿面.炉渣+本日.块矿面.炉渣", dec_炉渣, null, true);
            AddNumericTag("累计.1高炉.炉渣", "昨日累计.1高炉.炉渣+本日.1高炉.炉渣", dec_炉渣, null, true);
            AddNumericTag("累计.4高炉.炉渣", "昨日累计.4高炉.炉渣+本日.4高炉.炉渣", dec_炉渣, null, true);
            AddNumericTag("累计.5高炉.炉渣", "昨日累计.5高炉.炉渣+本日.5高炉.炉渣", dec_炉渣, null, true);
            AddNumericTag("累计.6高炉.炉渣", "昨日累计.6高炉.炉渣+本日.6高炉.炉渣", dec_炉渣, null, true);
            AddNumericTag("累计.7高炉.炉渣", "昨日累计.7高炉.炉渣+本日.7高炉.炉渣", dec_炉渣, null, true);
            AddNumericTag("累计.8高炉.炉渣", "昨日累计.8高炉.炉渣+本日.8高炉.炉渣", dec_炉渣, null, true);
            AddNumericTag("累计.扣返矿.炉渣", "昨日累计.扣返矿.炉渣-本日.扣返矿.炉渣", dec_炉渣, null, true);
            AddNumericTag("累计.炉渣", "累计.1高炉.炉渣+累计.4高炉.炉渣+累计.5高炉.炉渣+累计.6高炉.炉渣+累计.7高炉.炉渣+累计.8高炉.炉渣-累计.扣返矿.炉渣", dec_炉渣, null, false);






            AddNumericTag("昨日累计.供应.炉渣", "", dec_炉渣, null, false);
            AddNumericTag("昨日累计.块矿面.炉渣", "", dec_炉渣, null, false);
            AddNumericTag("昨日累计.1高炉.炉渣", "", dec_炉渣, null, false);
            AddNumericTag("昨日累计.4高炉.炉渣", "", dec_炉渣, null, false);
            AddNumericTag("昨日累计.5高炉.炉渣", "", dec_炉渣, null, false);
            AddNumericTag("昨日累计.6高炉.炉渣", "", dec_炉渣, null, false);
            AddNumericTag("昨日累计.7高炉.炉渣", "", dec_炉渣, null, false);
            AddNumericTag("昨日累计.8高炉.炉渣", "", dec_炉渣, null, false);
            AddNumericTag("昨日累计.扣返矿.炉渣", "", dec_炉渣, null, false);




            #endregion

            #region 水渣铁
            int? dec_水渣铁 = 2;
            AddNumericTag("本日.供应.水渣铁", "", dec_水渣铁, null, true);
            AddNumericTag("本日.块矿面.水渣铁", "", dec_水渣铁, null, true);
            AddNumericTag("本日.1高炉.水渣铁", "", dec_水渣铁, null, true);
            AddNumericTag("本日.4高炉.水渣铁", "", dec_水渣铁, null, true);
            AddNumericTag("本日.5高炉.水渣铁", "", dec_水渣铁, null, true);
            AddNumericTag("本日.6高炉.水渣铁", "", dec_水渣铁, null, true);
            AddNumericTag("本日.7高炉.水渣铁", "", dec_水渣铁, null, true);
            AddNumericTag("本日.8高炉.水渣铁", "", dec_水渣铁, null, true);
            AddNumericTag("本日.扣返矿.水渣铁", "", dec_水渣铁, null, true);
            AddNumericTag("本日.水渣铁", "本日.1高炉.水渣铁+本日.4高炉.水渣铁+本日.5高炉.水渣铁+本日.6高炉.水渣铁+本日.7高炉.水渣铁+本日.8高炉.水渣铁-本日.扣返矿.水渣铁", dec_水渣铁, null, true);

            AddNumericTag("本日库存.1高炉.水渣铁", "", dec_水渣铁, null, true);
            AddNumericTag("本日库存.二车间.水渣铁", "", dec_水渣铁, null, true);
            AddNumericTag("本日库存.三车间.水渣铁", "", dec_水渣铁, null, true);
            AddNumericTag("本日库存.水渣铁", "本日库存.1高炉.水渣铁+本日库存.二车间.水渣铁+本日库存.三车间.水渣铁", dec_水渣铁, null, true);

            AddNumericTag("累计.供应.水渣铁", "昨日累计.供应.水渣铁+本日.供应.水渣铁", dec_水渣铁, null, true);
            AddNumericTag("累计.块矿面.水渣铁", "昨日累计.块矿面.水渣铁+本日.块矿面.水渣铁", dec_水渣铁, null, true);
            AddNumericTag("累计.1高炉.水渣铁", "昨日累计.1高炉.水渣铁+本日.1高炉.水渣铁", dec_水渣铁, null, true);
            AddNumericTag("累计.4高炉.水渣铁", "昨日累计.4高炉.水渣铁+本日.4高炉.水渣铁", dec_水渣铁, null, true);
            AddNumericTag("累计.5高炉.水渣铁", "昨日累计.5高炉.水渣铁+本日.5高炉.水渣铁", dec_水渣铁, null, true);
            AddNumericTag("累计.6高炉.水渣铁", "昨日累计.6高炉.水渣铁+本日.6高炉.水渣铁", dec_水渣铁, null, true);
            AddNumericTag("累计.7高炉.水渣铁", "昨日累计.7高炉.水渣铁+本日.7高炉.水渣铁", dec_水渣铁, null, true);
            AddNumericTag("累计.8高炉.水渣铁", "昨日累计.8高炉.水渣铁+本日.8高炉.水渣铁", dec_水渣铁, null, true);
            AddNumericTag("累计.扣返矿.水渣铁", "昨日累计.扣返矿.水渣铁-本日.扣返矿.水渣铁", dec_水渣铁, null, true);
            AddNumericTag("累计.水渣铁", "累计.1高炉.水渣铁+累计.4高炉.水渣铁+累计.5高炉.水渣铁+累计.6高炉.水渣铁+累计.7高炉.水渣铁+累计.8高炉.水渣铁-累计.扣返矿.水渣铁", dec_水渣铁, null, false);






            AddNumericTag("昨日累计.供应.水渣铁", "", dec_水渣铁, null, false);
            AddNumericTag("昨日累计.块矿面.水渣铁", "", dec_水渣铁, null, false);
            AddNumericTag("昨日累计.1高炉.水渣铁", "", dec_水渣铁, null, false);
            AddNumericTag("昨日累计.4高炉.水渣铁", "", dec_水渣铁, null, false);
            AddNumericTag("昨日累计.5高炉.水渣铁", "", dec_水渣铁, null, false);
            AddNumericTag("昨日累计.6高炉.水渣铁", "", dec_水渣铁, null, false);
            AddNumericTag("昨日累计.7高炉.水渣铁", "", dec_水渣铁, null, false);
            AddNumericTag("昨日累计.8高炉.水渣铁", "", dec_水渣铁, null, false);
            AddNumericTag("昨日累计.扣返矿.水渣铁", "", dec_水渣铁, null, false);




            #endregion

            #region 破铁面
            int? dec_破铁面 = 2;
            AddNumericTag("本日.供应.破铁面", "", dec_破铁面, null, true);
            AddNumericTag("本日.块矿面.破铁面", "", dec_破铁面, null, true);
            AddNumericTag("本日.1高炉.破铁面", "", dec_破铁面, null, true);
            AddNumericTag("本日.4高炉.破铁面", "", dec_破铁面, null, true);
            AddNumericTag("本日.5高炉.破铁面", "", dec_破铁面, null, true);
            AddNumericTag("本日.6高炉.破铁面", "", dec_破铁面, null, true);
            AddNumericTag("本日.7高炉.破铁面", "", dec_破铁面, null, true);
            AddNumericTag("本日.8高炉.破铁面", "", dec_破铁面, null, true);
            AddNumericTag("本日.扣返矿.破铁面", "", dec_破铁面, null, true);
            AddNumericTag("本日.破铁面", "本日.1高炉.破铁面+本日.4高炉.破铁面+本日.5高炉.破铁面+本日.6高炉.破铁面+本日.7高炉.破铁面+本日.8高炉.破铁面-本日.扣返矿.破铁面", dec_破铁面, null, true);

            AddNumericTag("本日库存.1高炉.破铁面", "", dec_破铁面, null, true);
            AddNumericTag("本日库存.二车间.破铁面", "", dec_破铁面, null, true);
            AddNumericTag("本日库存.三车间.破铁面", "", dec_破铁面, null, true);
            AddNumericTag("本日库存.破铁面", "本日库存.1高炉.破铁面+本日库存.二车间.破铁面+本日库存.三车间.破铁面", dec_破铁面, null, true);

            AddNumericTag("累计.供应.破铁面", "昨日累计.供应.破铁面+本日.供应.破铁面", dec_破铁面, null, true);
            AddNumericTag("累计.块矿面.破铁面", "昨日累计.块矿面.破铁面+本日.块矿面.破铁面", dec_破铁面, null, true);
            AddNumericTag("累计.1高炉.破铁面", "昨日累计.1高炉.破铁面+本日.1高炉.破铁面", dec_破铁面, null, true);
            AddNumericTag("累计.4高炉.破铁面", "昨日累计.4高炉.破铁面+本日.4高炉.破铁面", dec_破铁面, null, true);
            AddNumericTag("累计.5高炉.破铁面", "昨日累计.5高炉.破铁面+本日.5高炉.破铁面", dec_破铁面, null, true);
            AddNumericTag("累计.6高炉.破铁面", "昨日累计.6高炉.破铁面+本日.6高炉.破铁面", dec_破铁面, null, true);
            AddNumericTag("累计.7高炉.破铁面", "昨日累计.7高炉.破铁面+本日.7高炉.破铁面", dec_破铁面, null, true);
            AddNumericTag("累计.8高炉.破铁面", "昨日累计.8高炉.破铁面+本日.8高炉.破铁面", dec_破铁面, null, true);
            AddNumericTag("累计.扣返矿.破铁面", "昨日累计.扣返矿.破铁面-本日.扣返矿.破铁面", dec_破铁面, null, true);
            AddNumericTag("累计.破铁面", "累计.1高炉.破铁面+累计.4高炉.破铁面+累计.5高炉.破铁面+累计.6高炉.破铁面+累计.7高炉.破铁面+累计.8高炉.破铁面-累计.扣返矿.破铁面", dec_破铁面, null, false);






            AddNumericTag("昨日累计.供应.破铁面", "", dec_破铁面, null, false);
            AddNumericTag("昨日累计.块矿面.破铁面", "", dec_破铁面, null, false);
            AddNumericTag("昨日累计.1高炉.破铁面", "", dec_破铁面, null, false);
            AddNumericTag("昨日累计.4高炉.破铁面", "", dec_破铁面, null, false);
            AddNumericTag("昨日累计.5高炉.破铁面", "", dec_破铁面, null, false);
            AddNumericTag("昨日累计.6高炉.破铁面", "", dec_破铁面, null, false);
            AddNumericTag("昨日累计.7高炉.破铁面", "", dec_破铁面, null, false);
            AddNumericTag("昨日累计.8高炉.破铁面", "", dec_破铁面, null, false);
            AddNumericTag("昨日累计.扣返矿.破铁面", "", dec_破铁面, null, false);




            #endregion

            #region 破铁块
            int? dec_破铁块 = 2;
            AddNumericTag("本日.供应.破铁块", "", dec_破铁块, null, true);
            AddNumericTag("本日.块矿面.破铁块", "", dec_破铁块, null, true);
            AddNumericTag("本日.1高炉.破铁块", "", dec_破铁块, null, true);
            AddNumericTag("本日.4高炉.破铁块", "", dec_破铁块, null, true);
            AddNumericTag("本日.5高炉.破铁块", "", dec_破铁块, null, true);
            AddNumericTag("本日.6高炉.破铁块", "", dec_破铁块, null, true);
            AddNumericTag("本日.7高炉.破铁块", "", dec_破铁块, null, true);
            AddNumericTag("本日.8高炉.破铁块", "", dec_破铁块, null, true);
            AddNumericTag("本日.扣返矿.破铁块", "", dec_破铁块, null, true);
            AddNumericTag("本日.破铁块", "本日.1高炉.破铁块+本日.4高炉.破铁块+本日.5高炉.破铁块+本日.6高炉.破铁块+本日.7高炉.破铁块+本日.8高炉.破铁块-本日.扣返矿.破铁块", dec_破铁块, null, true);

            AddNumericTag("本日库存.1高炉.破铁块", "", dec_破铁块, null, true);
            AddNumericTag("本日库存.二车间.破铁块", "", dec_破铁块, null, true);
            AddNumericTag("本日库存.三车间.破铁块", "", dec_破铁块, null, true);
            AddNumericTag("本日库存.破铁块", "本日库存.1高炉.破铁块+本日库存.二车间.破铁块+本日库存.三车间.破铁块", dec_破铁块, null, true);

            AddNumericTag("累计.供应.破铁块", "昨日累计.供应.破铁块+本日.供应.破铁块", dec_破铁块, null, true);
            AddNumericTag("累计.块矿面.破铁块", "昨日累计.块矿面.破铁块+本日.块矿面.破铁块", dec_破铁块, null, true);
            AddNumericTag("累计.1高炉.破铁块", "昨日累计.1高炉.破铁块+本日.1高炉.破铁块", dec_破铁块, null, true);
            AddNumericTag("累计.4高炉.破铁块", "昨日累计.4高炉.破铁块+本日.4高炉.破铁块", dec_破铁块, null, true);
            AddNumericTag("累计.5高炉.破铁块", "昨日累计.5高炉.破铁块+本日.5高炉.破铁块", dec_破铁块, null, true);
            AddNumericTag("累计.6高炉.破铁块", "昨日累计.6高炉.破铁块+本日.6高炉.破铁块", dec_破铁块, null, true);
            AddNumericTag("累计.7高炉.破铁块", "昨日累计.7高炉.破铁块+本日.7高炉.破铁块", dec_破铁块, null, true);
            AddNumericTag("累计.8高炉.破铁块", "昨日累计.8高炉.破铁块+本日.8高炉.破铁块", dec_破铁块, null, true);
            AddNumericTag("累计.扣返矿.破铁块", "昨日累计.扣返矿.破铁块-本日.扣返矿.破铁块", dec_破铁块, null, true);
            AddNumericTag("累计.破铁块", "累计.1高炉.破铁块+累计.4高炉.破铁块+累计.5高炉.破铁块+累计.6高炉.破铁块+累计.7高炉.破铁块+累计.8高炉.破铁块-累计.扣返矿.破铁块", dec_破铁块, null, false);






            AddNumericTag("昨日累计.供应.破铁块", "", dec_破铁块, null, false);
            AddNumericTag("昨日累计.块矿面.破铁块", "", dec_破铁块, null, false);
            AddNumericTag("昨日累计.1高炉.破铁块", "", dec_破铁块, null, false);
            AddNumericTag("昨日累计.4高炉.破铁块", "", dec_破铁块, null, false);
            AddNumericTag("昨日累计.5高炉.破铁块", "", dec_破铁块, null, false);
            AddNumericTag("昨日累计.6高炉.破铁块", "", dec_破铁块, null, false);
            AddNumericTag("昨日累计.7高炉.破铁块", "", dec_破铁块, null, false);
            AddNumericTag("昨日累计.8高炉.破铁块", "", dec_破铁块, null, false);
            AddNumericTag("昨日累计.扣返矿.破铁块", "", dec_破铁块, null, false);




            #endregion

            #region 微粉水渣铁
            int? dec_微粉水渣铁 = 2;
            AddNumericTag("本日.供应.微粉水渣铁", "", dec_微粉水渣铁, null, true);
            AddNumericTag("本日.块矿面.微粉水渣铁", "", dec_微粉水渣铁, null, true);
            AddNumericTag("本日.1高炉.微粉水渣铁", "", dec_微粉水渣铁, null, true);
            AddNumericTag("本日.4高炉.微粉水渣铁", "", dec_微粉水渣铁, null, true);
            AddNumericTag("本日.5高炉.微粉水渣铁", "", dec_微粉水渣铁, null, true);
            AddNumericTag("本日.6高炉.微粉水渣铁", "", dec_微粉水渣铁, null, true);
            AddNumericTag("本日.7高炉.微粉水渣铁", "", dec_微粉水渣铁, null, true);
            AddNumericTag("本日.8高炉.微粉水渣铁", "", dec_微粉水渣铁, null, true);
            AddNumericTag("本日.扣返矿.微粉水渣铁", "", dec_微粉水渣铁, null, true);
            AddNumericTag("本日.微粉水渣铁", "本日.1高炉.微粉水渣铁+本日.4高炉.微粉水渣铁+本日.5高炉.微粉水渣铁+本日.6高炉.微粉水渣铁+本日.7高炉.微粉水渣铁+本日.8高炉.微粉水渣铁-本日.扣返矿.微粉水渣铁", dec_微粉水渣铁, null, true);

            AddNumericTag("本日库存.1高炉.微粉水渣铁", "", dec_微粉水渣铁, null, true);
            AddNumericTag("本日库存.二车间.微粉水渣铁", "", dec_微粉水渣铁, null, true);
            AddNumericTag("本日库存.三车间.微粉水渣铁", "", dec_微粉水渣铁, null, true);
            AddNumericTag("本日库存.微粉水渣铁", "本日库存.1高炉.微粉水渣铁+本日库存.二车间.微粉水渣铁+本日库存.三车间.微粉水渣铁", dec_微粉水渣铁, null, true);

            AddNumericTag("累计.供应.微粉水渣铁", "昨日累计.供应.微粉水渣铁+本日.供应.微粉水渣铁", dec_微粉水渣铁, null, true);
            AddNumericTag("累计.块矿面.微粉水渣铁", "昨日累计.块矿面.微粉水渣铁+本日.块矿面.微粉水渣铁", dec_微粉水渣铁, null, true);
            AddNumericTag("累计.1高炉.微粉水渣铁", "昨日累计.1高炉.微粉水渣铁+本日.1高炉.微粉水渣铁", dec_微粉水渣铁, null, true);
            AddNumericTag("累计.4高炉.微粉水渣铁", "昨日累计.4高炉.微粉水渣铁+本日.4高炉.微粉水渣铁", dec_微粉水渣铁, null, true);
            AddNumericTag("累计.5高炉.微粉水渣铁", "昨日累计.5高炉.微粉水渣铁+本日.5高炉.微粉水渣铁", dec_微粉水渣铁, null, true);
            AddNumericTag("累计.6高炉.微粉水渣铁", "昨日累计.6高炉.微粉水渣铁+本日.6高炉.微粉水渣铁", dec_微粉水渣铁, null, true);
            AddNumericTag("累计.7高炉.微粉水渣铁", "昨日累计.7高炉.微粉水渣铁+本日.7高炉.微粉水渣铁", dec_微粉水渣铁, null, true);
            AddNumericTag("累计.8高炉.微粉水渣铁", "昨日累计.8高炉.微粉水渣铁+本日.8高炉.微粉水渣铁", dec_微粉水渣铁, null, true);
            AddNumericTag("累计.扣返矿.微粉水渣铁", "昨日累计.扣返矿.微粉水渣铁-本日.扣返矿.微粉水渣铁", dec_微粉水渣铁, null, true);
            AddNumericTag("累计.微粉水渣铁", "累计.1高炉.微粉水渣铁+累计.4高炉.微粉水渣铁+累计.5高炉.微粉水渣铁+累计.6高炉.微粉水渣铁+累计.7高炉.微粉水渣铁+累计.8高炉.微粉水渣铁-累计.扣返矿.微粉水渣铁", dec_微粉水渣铁, null, false);






            AddNumericTag("昨日累计.供应.微粉水渣铁", "", dec_微粉水渣铁, null, false);
            AddNumericTag("昨日累计.块矿面.微粉水渣铁", "", dec_微粉水渣铁, null, false);
            AddNumericTag("昨日累计.1高炉.微粉水渣铁", "", dec_微粉水渣铁, null, false);
            AddNumericTag("昨日累计.4高炉.微粉水渣铁", "", dec_微粉水渣铁, null, false);
            AddNumericTag("昨日累计.5高炉.微粉水渣铁", "", dec_微粉水渣铁, null, false);
            AddNumericTag("昨日累计.6高炉.微粉水渣铁", "", dec_微粉水渣铁, null, false);
            AddNumericTag("昨日累计.7高炉.微粉水渣铁", "", dec_微粉水渣铁, null, false);
            AddNumericTag("昨日累计.8高炉.微粉水渣铁", "", dec_微粉水渣铁, null, false);
            AddNumericTag("昨日累计.扣返矿.微粉水渣铁", "", dec_微粉水渣铁, null, false);




            #endregion

            #region 渣铁
            int? dec_渣铁 = 2;
            AddNumericTag("本日.供应.渣铁", "", dec_渣铁, null, true);
            AddNumericTag("本日.块矿面.渣铁", "", dec_渣铁, null, true);
            AddNumericTag("本日.1高炉.渣铁", "", dec_渣铁, null, true);
            AddNumericTag("本日.4高炉.渣铁", "", dec_渣铁, null, true);
            AddNumericTag("本日.5高炉.渣铁", "", dec_渣铁, null, true);
            AddNumericTag("本日.6高炉.渣铁", "", dec_渣铁, null, true);
            AddNumericTag("本日.7高炉.渣铁", "", dec_渣铁, null, true);
            AddNumericTag("本日.8高炉.渣铁", "", dec_渣铁, null, true);
            AddNumericTag("本日.扣返矿.渣铁", "", dec_渣铁, null, true);
            AddNumericTag("本日.渣铁", "本日.1高炉.渣铁+本日.4高炉.渣铁+本日.5高炉.渣铁+本日.6高炉.渣铁+本日.7高炉.渣铁+本日.8高炉.渣铁-本日.扣返矿.渣铁", dec_渣铁, null, true);

            AddNumericTag("本日库存.1高炉.渣铁", "", dec_渣铁, null, true);
            AddNumericTag("本日库存.二车间.渣铁", "", dec_渣铁, null, true);
            AddNumericTag("本日库存.三车间.渣铁", "", dec_渣铁, null, true);
            AddNumericTag("本日库存.渣铁", "本日库存.1高炉.渣铁+本日库存.二车间.渣铁+本日库存.三车间.渣铁", dec_渣铁, null, true);

            AddNumericTag("累计.供应.渣铁", "昨日累计.供应.渣铁+本日.供应.渣铁", dec_渣铁, null, true);
            AddNumericTag("累计.块矿面.渣铁", "昨日累计.块矿面.渣铁+本日.块矿面.渣铁", dec_渣铁, null, true);
            AddNumericTag("累计.1高炉.渣铁", "昨日累计.1高炉.渣铁+本日.1高炉.渣铁", dec_渣铁, null, true);
            AddNumericTag("累计.4高炉.渣铁", "昨日累计.4高炉.渣铁+本日.4高炉.渣铁", dec_渣铁, null, true);
            AddNumericTag("累计.5高炉.渣铁", "昨日累计.5高炉.渣铁+本日.5高炉.渣铁", dec_渣铁, null, true);
            AddNumericTag("累计.6高炉.渣铁", "昨日累计.6高炉.渣铁+本日.6高炉.渣铁", dec_渣铁, null, true);
            AddNumericTag("累计.7高炉.渣铁", "昨日累计.7高炉.渣铁+本日.7高炉.渣铁", dec_渣铁, null, true);
            AddNumericTag("累计.8高炉.渣铁", "昨日累计.8高炉.渣铁+本日.8高炉.渣铁", dec_渣铁, null, true);
            AddNumericTag("累计.扣返矿.渣铁", "昨日累计.扣返矿.渣铁-本日.扣返矿.渣铁", dec_渣铁, null, true);
            AddNumericTag("累计.渣铁", "累计.1高炉.渣铁+累计.4高炉.渣铁+累计.5高炉.渣铁+累计.6高炉.渣铁+累计.7高炉.渣铁+累计.8高炉.渣铁-累计.扣返矿.渣铁", dec_渣铁, null, false);






            AddNumericTag("昨日累计.供应.渣铁", "", dec_渣铁, null, false);
            AddNumericTag("昨日累计.块矿面.渣铁", "", dec_渣铁, null, false);
            AddNumericTag("昨日累计.1高炉.渣铁", "", dec_渣铁, null, false);
            AddNumericTag("昨日累计.4高炉.渣铁", "", dec_渣铁, null, false);
            AddNumericTag("昨日累计.5高炉.渣铁", "", dec_渣铁, null, false);
            AddNumericTag("昨日累计.6高炉.渣铁", "", dec_渣铁, null, false);
            AddNumericTag("昨日累计.7高炉.渣铁", "", dec_渣铁, null, false);
            AddNumericTag("昨日累计.8高炉.渣铁", "", dec_渣铁, null, false);
            AddNumericTag("昨日累计.扣返矿.渣铁", "", dec_渣铁, null, false);




            #endregion

            #region 外购焦粉
            int? dec_外购焦粉 = 2;
            AddNumericTag("本日.供应.外购焦粉", "", dec_外购焦粉, null, true);
            AddNumericTag("本日.块矿面.外购焦粉", "", dec_外购焦粉, null, true);
            AddNumericTag("本日.1高炉.外购焦粉", "", dec_外购焦粉, null, true);
            AddNumericTag("本日.4高炉.外购焦粉", "", dec_外购焦粉, null, true);
            AddNumericTag("本日.5高炉.外购焦粉", "", dec_外购焦粉, null, true);
            AddNumericTag("本日.6高炉.外购焦粉", "", dec_外购焦粉, null, true);
            AddNumericTag("本日.7高炉.外购焦粉", "", dec_外购焦粉, null, true);
            AddNumericTag("本日.8高炉.外购焦粉", "", dec_外购焦粉, null, true);
            AddNumericTag("本日.扣返矿.外购焦粉", "", dec_外购焦粉, null, true);
            AddNumericTag("本日.外购焦粉", "本日.1高炉.外购焦粉+本日.4高炉.外购焦粉+本日.5高炉.外购焦粉+本日.6高炉.外购焦粉+本日.7高炉.外购焦粉+本日.8高炉.外购焦粉-本日.扣返矿.外购焦粉", dec_外购焦粉, null, true);

            AddNumericTag("本日库存.1高炉.外购焦粉", "", dec_外购焦粉, null, true);
            AddNumericTag("本日库存.二车间.外购焦粉", "", dec_外购焦粉, null, true);
            AddNumericTag("本日库存.三车间.外购焦粉", "", dec_外购焦粉, null, true);
            AddNumericTag("本日库存.外购焦粉", "本日库存.1高炉.外购焦粉+本日库存.二车间.外购焦粉+本日库存.三车间.外购焦粉", dec_外购焦粉, null, true);

            AddNumericTag("累计.供应.外购焦粉", "昨日累计.供应.外购焦粉+本日.供应.外购焦粉", dec_外购焦粉, null, true);
            AddNumericTag("累计.块矿面.外购焦粉", "昨日累计.块矿面.外购焦粉+本日.块矿面.外购焦粉", dec_外购焦粉, null, true);
            AddNumericTag("累计.1高炉.外购焦粉", "昨日累计.1高炉.外购焦粉+本日.1高炉.外购焦粉", dec_外购焦粉, null, true);
            AddNumericTag("累计.4高炉.外购焦粉", "昨日累计.4高炉.外购焦粉+本日.4高炉.外购焦粉", dec_外购焦粉, null, true);
            AddNumericTag("累计.5高炉.外购焦粉", "昨日累计.5高炉.外购焦粉+本日.5高炉.外购焦粉", dec_外购焦粉, null, true);
            AddNumericTag("累计.6高炉.外购焦粉", "昨日累计.6高炉.外购焦粉+本日.6高炉.外购焦粉", dec_外购焦粉, null, true);
            AddNumericTag("累计.7高炉.外购焦粉", "昨日累计.7高炉.外购焦粉+本日.7高炉.外购焦粉", dec_外购焦粉, null, true);
            AddNumericTag("累计.8高炉.外购焦粉", "昨日累计.8高炉.外购焦粉+本日.8高炉.外购焦粉", dec_外购焦粉, null, true);
            AddNumericTag("累计.扣返矿.外购焦粉", "昨日累计.扣返矿.外购焦粉-本日.扣返矿.外购焦粉", dec_外购焦粉, null, true);
            AddNumericTag("累计.外购焦粉", "累计.1高炉.外购焦粉+累计.4高炉.外购焦粉+累计.5高炉.外购焦粉+累计.6高炉.外购焦粉+累计.7高炉.外购焦粉+累计.8高炉.外购焦粉-累计.扣返矿.外购焦粉", dec_外购焦粉, null, false);






            AddNumericTag("昨日累计.供应.外购焦粉", "", dec_外购焦粉, null, false);
            AddNumericTag("昨日累计.块矿面.外购焦粉", "", dec_外购焦粉, null, false);
            AddNumericTag("昨日累计.1高炉.外购焦粉", "", dec_外购焦粉, null, false);
            AddNumericTag("昨日累计.4高炉.外购焦粉", "", dec_外购焦粉, null, false);
            AddNumericTag("昨日累计.5高炉.外购焦粉", "", dec_外购焦粉, null, false);
            AddNumericTag("昨日累计.6高炉.外购焦粉", "", dec_外购焦粉, null, false);
            AddNumericTag("昨日累计.7高炉.外购焦粉", "", dec_外购焦粉, null, false);
            AddNumericTag("昨日累计.8高炉.外购焦粉", "", dec_外购焦粉, null, false);
            AddNumericTag("昨日累计.扣返矿.外购焦粉", "", dec_外购焦粉, null, false);




            #endregion

            #region 新焦化小焦
            int? dec_新焦化小焦 = 2;
            AddNumericTag("本日.供应.新焦化小焦", "", dec_新焦化小焦, null, true);
            AddNumericTag("本日.块矿面.新焦化小焦", "", dec_新焦化小焦, null, true);
            AddNumericTag("本日.1高炉.新焦化小焦", "", dec_新焦化小焦, null, true);
            AddNumericTag("本日.4高炉.新焦化小焦", "", dec_新焦化小焦, null, true);
            AddNumericTag("本日.5高炉.新焦化小焦", "", dec_新焦化小焦, null, true);
            AddNumericTag("本日.6高炉.新焦化小焦", "", dec_新焦化小焦, null, true);
            AddNumericTag("本日.7高炉.新焦化小焦", "", dec_新焦化小焦, null, true);
            AddNumericTag("本日.8高炉.新焦化小焦", "", dec_新焦化小焦, null, true);
            AddNumericTag("本日.扣返矿.新焦化小焦", "", dec_新焦化小焦, null, true);
            AddNumericTag("本日.新焦化小焦", "本日.1高炉.新焦化小焦+本日.4高炉.新焦化小焦+本日.5高炉.新焦化小焦+本日.6高炉.新焦化小焦+本日.7高炉.新焦化小焦+本日.8高炉.新焦化小焦-本日.扣返矿.新焦化小焦", dec_新焦化小焦, null, true);

            AddNumericTag("本日库存.1高炉.新焦化小焦", "", dec_新焦化小焦, null, true);
            AddNumericTag("本日库存.二车间.新焦化小焦", "", dec_新焦化小焦, null, true);
            AddNumericTag("本日库存.三车间.新焦化小焦", "", dec_新焦化小焦, null, true);
            AddNumericTag("本日库存.新焦化小焦", "本日库存.1高炉.新焦化小焦+本日库存.二车间.新焦化小焦+本日库存.三车间.新焦化小焦", dec_新焦化小焦, null, true);

            AddNumericTag("累计.供应.新焦化小焦", "昨日累计.供应.新焦化小焦+本日.供应.新焦化小焦", dec_新焦化小焦, null, true);
            AddNumericTag("累计.块矿面.新焦化小焦", "昨日累计.块矿面.新焦化小焦+本日.块矿面.新焦化小焦", dec_新焦化小焦, null, true);
            AddNumericTag("累计.1高炉.新焦化小焦", "昨日累计.1高炉.新焦化小焦+本日.1高炉.新焦化小焦", dec_新焦化小焦, null, true);
            AddNumericTag("累计.4高炉.新焦化小焦", "昨日累计.4高炉.新焦化小焦+本日.4高炉.新焦化小焦", dec_新焦化小焦, null, true);
            AddNumericTag("累计.5高炉.新焦化小焦", "昨日累计.5高炉.新焦化小焦+本日.5高炉.新焦化小焦", dec_新焦化小焦, null, true);
            AddNumericTag("累计.6高炉.新焦化小焦", "昨日累计.6高炉.新焦化小焦+本日.6高炉.新焦化小焦", dec_新焦化小焦, null, true);
            AddNumericTag("累计.7高炉.新焦化小焦", "昨日累计.7高炉.新焦化小焦+本日.7高炉.新焦化小焦", dec_新焦化小焦, null, true);
            AddNumericTag("累计.8高炉.新焦化小焦", "昨日累计.8高炉.新焦化小焦+本日.8高炉.新焦化小焦", dec_新焦化小焦, null, true);
            AddNumericTag("累计.扣返矿.新焦化小焦", "昨日累计.扣返矿.新焦化小焦-本日.扣返矿.新焦化小焦", dec_新焦化小焦, null, true);
            AddNumericTag("累计.新焦化小焦", "累计.1高炉.新焦化小焦+累计.4高炉.新焦化小焦+累计.5高炉.新焦化小焦+累计.6高炉.新焦化小焦+累计.7高炉.新焦化小焦+累计.8高炉.新焦化小焦-累计.扣返矿.新焦化小焦", dec_新焦化小焦, null, false);






            AddNumericTag("昨日累计.供应.新焦化小焦", "", dec_新焦化小焦, null, false);
            AddNumericTag("昨日累计.块矿面.新焦化小焦", "", dec_新焦化小焦, null, false);
            AddNumericTag("昨日累计.1高炉.新焦化小焦", "", dec_新焦化小焦, null, false);
            AddNumericTag("昨日累计.4高炉.新焦化小焦", "", dec_新焦化小焦, null, false);
            AddNumericTag("昨日累计.5高炉.新焦化小焦", "", dec_新焦化小焦, null, false);
            AddNumericTag("昨日累计.6高炉.新焦化小焦", "", dec_新焦化小焦, null, false);
            AddNumericTag("昨日累计.7高炉.新焦化小焦", "", dec_新焦化小焦, null, false);
            AddNumericTag("昨日累计.8高炉.新焦化小焦", "", dec_新焦化小焦, null, false);
            AddNumericTag("昨日累计.扣返矿.新焦化小焦", "", dec_新焦化小焦, null, false);




            #endregion

            #region 外购焦炭
            int? dec_外购焦炭 = 2;
            AddNumericTag("本日.供应.外购焦炭", "", dec_外购焦炭, null, true);
            AddNumericTag("本日.块矿面.外购焦炭", "", dec_外购焦炭, null, true);
            AddNumericTag("本日.1高炉.外购焦炭", "", dec_外购焦炭, null, true);
            AddNumericTag("本日.4高炉.外购焦炭", "", dec_外购焦炭, null, true);
            AddNumericTag("本日.5高炉.外购焦炭", "", dec_外购焦炭, null, true);
            AddNumericTag("本日.6高炉.外购焦炭", "", dec_外购焦炭, null, true);
            AddNumericTag("本日.7高炉.外购焦炭", "", dec_外购焦炭, null, true);
            AddNumericTag("本日.8高炉.外购焦炭", "", dec_外购焦炭, null, true);
            AddNumericTag("本日.扣返矿.外购焦炭", "", dec_外购焦炭, null, true);
            AddNumericTag("本日.外购焦炭", "本日.1高炉.外购焦炭+本日.4高炉.外购焦炭+本日.5高炉.外购焦炭+本日.6高炉.外购焦炭+本日.7高炉.外购焦炭+本日.8高炉.外购焦炭-本日.扣返矿.外购焦炭", dec_外购焦炭, null, true);

            AddNumericTag("本日库存.1高炉.外购焦炭", "", dec_外购焦炭, null, true);
            AddNumericTag("本日库存.二车间.外购焦炭", "", dec_外购焦炭, null, true);
            AddNumericTag("本日库存.三车间.外购焦炭", "", dec_外购焦炭, null, true);
            AddNumericTag("本日库存.外购焦炭", "本日库存.1高炉.外购焦炭+本日库存.二车间.外购焦炭+本日库存.三车间.外购焦炭", dec_外购焦炭, null, true);

            AddNumericTag("累计.供应.外购焦炭", "昨日累计.供应.外购焦炭+本日.供应.外购焦炭", dec_外购焦炭, null, true);
            AddNumericTag("累计.块矿面.外购焦炭", "昨日累计.块矿面.外购焦炭+本日.块矿面.外购焦炭", dec_外购焦炭, null, true);
            AddNumericTag("累计.1高炉.外购焦炭", "昨日累计.1高炉.外购焦炭+本日.1高炉.外购焦炭", dec_外购焦炭, null, true);
            AddNumericTag("累计.4高炉.外购焦炭", "昨日累计.4高炉.外购焦炭+本日.4高炉.外购焦炭", dec_外购焦炭, null, true);
            AddNumericTag("累计.5高炉.外购焦炭", "昨日累计.5高炉.外购焦炭+本日.5高炉.外购焦炭", dec_外购焦炭, null, true);
            AddNumericTag("累计.6高炉.外购焦炭", "昨日累计.6高炉.外购焦炭+本日.6高炉.外购焦炭", dec_外购焦炭, null, true);
            AddNumericTag("累计.7高炉.外购焦炭", "昨日累计.7高炉.外购焦炭+本日.7高炉.外购焦炭", dec_外购焦炭, null, true);
            AddNumericTag("累计.8高炉.外购焦炭", "昨日累计.8高炉.外购焦炭+本日.8高炉.外购焦炭", dec_外购焦炭, null, true);
            AddNumericTag("累计.扣返矿.外购焦炭", "昨日累计.扣返矿.外购焦炭-本日.扣返矿.外购焦炭", dec_外购焦炭, null, true);
            AddNumericTag("累计.外购焦炭", "累计.1高炉.外购焦炭+累计.4高炉.外购焦炭+累计.5高炉.外购焦炭+累计.6高炉.外购焦炭+累计.7高炉.外购焦炭+累计.8高炉.外购焦炭-累计.扣返矿.外购焦炭", dec_外购焦炭, null, false);






            AddNumericTag("昨日累计.供应.外购焦炭", "", dec_外购焦炭, null, false);
            AddNumericTag("昨日累计.块矿面.外购焦炭", "", dec_外购焦炭, null, false);
            AddNumericTag("昨日累计.1高炉.外购焦炭", "", dec_外购焦炭, null, false);
            AddNumericTag("昨日累计.4高炉.外购焦炭", "", dec_外购焦炭, null, false);
            AddNumericTag("昨日累计.5高炉.外购焦炭", "", dec_外购焦炭, null, false);
            AddNumericTag("昨日累计.6高炉.外购焦炭", "", dec_外购焦炭, null, false);
            AddNumericTag("昨日累计.7高炉.外购焦炭", "", dec_外购焦炭, null, false);
            AddNumericTag("昨日累计.8高炉.外购焦炭", "", dec_外购焦炭, null, false);
            AddNumericTag("昨日累计.扣返矿.外购焦炭", "", dec_外购焦炭, null, false);




            #endregion

            #region 自产焦炭
            int? dec_自产焦炭 = 2;
            AddNumericTag("本日.供应.自产焦炭", "", dec_自产焦炭, null, true);
            AddNumericTag("本日.块矿面.自产焦炭", "", dec_自产焦炭, null, true);
            AddNumericTag("本日.1高炉.自产焦炭", "", dec_自产焦炭, null, true);
            AddNumericTag("本日.4高炉.自产焦炭", "", dec_自产焦炭, null, true);
            AddNumericTag("本日.5高炉.自产焦炭", "", dec_自产焦炭, null, true);
            AddNumericTag("本日.6高炉.自产焦炭", "", dec_自产焦炭, null, true);
            AddNumericTag("本日.7高炉.自产焦炭", "", dec_自产焦炭, null, true);
            AddNumericTag("本日.8高炉.自产焦炭", "", dec_自产焦炭, null, true);
            AddNumericTag("本日.扣返矿.自产焦炭", "", dec_自产焦炭, null, true);
            AddNumericTag("本日.自产焦炭", "本日.1高炉.自产焦炭+本日.4高炉.自产焦炭+本日.5高炉.自产焦炭+本日.6高炉.自产焦炭+本日.7高炉.自产焦炭+本日.8高炉.自产焦炭-本日.扣返矿.自产焦炭", dec_自产焦炭, null, true);

            AddNumericTag("本日库存.1高炉.自产焦炭", "", dec_自产焦炭, null, true);
            AddNumericTag("本日库存.二车间.自产焦炭", "", dec_自产焦炭, null, true);
            AddNumericTag("本日库存.三车间.自产焦炭", "", dec_自产焦炭, null, true);
            AddNumericTag("本日库存.自产焦炭", "本日库存.1高炉.自产焦炭+本日库存.二车间.自产焦炭+本日库存.三车间.自产焦炭", dec_自产焦炭, null, true);

            AddNumericTag("累计.供应.自产焦炭", "昨日累计.供应.自产焦炭+本日.供应.自产焦炭", dec_自产焦炭, null, true);
            AddNumericTag("累计.块矿面.自产焦炭", "昨日累计.块矿面.自产焦炭+本日.块矿面.自产焦炭", dec_自产焦炭, null, true);
            AddNumericTag("累计.1高炉.自产焦炭", "昨日累计.1高炉.自产焦炭+本日.1高炉.自产焦炭", dec_自产焦炭, null, true);
            AddNumericTag("累计.4高炉.自产焦炭", "昨日累计.4高炉.自产焦炭+本日.4高炉.自产焦炭", dec_自产焦炭, null, true);
            AddNumericTag("累计.5高炉.自产焦炭", "昨日累计.5高炉.自产焦炭+本日.5高炉.自产焦炭", dec_自产焦炭, null, true);
            AddNumericTag("累计.6高炉.自产焦炭", "昨日累计.6高炉.自产焦炭+本日.6高炉.自产焦炭", dec_自产焦炭, null, true);
            AddNumericTag("累计.7高炉.自产焦炭", "昨日累计.7高炉.自产焦炭+本日.7高炉.自产焦炭", dec_自产焦炭, null, true);
            AddNumericTag("累计.8高炉.自产焦炭", "昨日累计.8高炉.自产焦炭+本日.8高炉.自产焦炭", dec_自产焦炭, null, true);
            AddNumericTag("累计.扣返矿.自产焦炭", "昨日累计.扣返矿.自产焦炭-本日.扣返矿.自产焦炭", dec_自产焦炭, null, true);
            AddNumericTag("累计.自产焦炭", "累计.1高炉.自产焦炭+累计.4高炉.自产焦炭+累计.5高炉.自产焦炭+累计.6高炉.自产焦炭+累计.7高炉.自产焦炭+累计.8高炉.自产焦炭-累计.扣返矿.自产焦炭", dec_自产焦炭, null, false);






            AddNumericTag("昨日累计.供应.自产焦炭", "", dec_自产焦炭, null, false);
            AddNumericTag("昨日累计.块矿面.自产焦炭", "", dec_自产焦炭, null, false);
            AddNumericTag("昨日累计.1高炉.自产焦炭", "", dec_自产焦炭, null, false);
            AddNumericTag("昨日累计.4高炉.自产焦炭", "", dec_自产焦炭, null, false);
            AddNumericTag("昨日累计.5高炉.自产焦炭", "", dec_自产焦炭, null, false);
            AddNumericTag("昨日累计.6高炉.自产焦炭", "", dec_自产焦炭, null, false);
            AddNumericTag("昨日累计.7高炉.自产焦炭", "", dec_自产焦炭, null, false);
            AddNumericTag("昨日累计.8高炉.自产焦炭", "", dec_自产焦炭, null, false);
            AddNumericTag("昨日累计.扣返矿.自产焦炭", "", dec_自产焦炭, null, false);




            #endregion

            #region 喷煤粉
            int? dec_喷煤粉 = 2;
            AddNumericTag("本日.供应.喷煤粉", "", dec_喷煤粉, null, true);
            AddNumericTag("本日.块矿面.喷煤粉", "", dec_喷煤粉, null, true);
            AddNumericTag("本日.1高炉.喷煤粉", "", dec_喷煤粉, null, true);
            AddNumericTag("本日.4高炉.喷煤粉", "", dec_喷煤粉, null, true);
            AddNumericTag("本日.5高炉.喷煤粉", "", dec_喷煤粉, null, true);
            AddNumericTag("本日.6高炉.喷煤粉", "", dec_喷煤粉, null, true);
            AddNumericTag("本日.7高炉.喷煤粉", "", dec_喷煤粉, null, true);
            AddNumericTag("本日.8高炉.喷煤粉", "", dec_喷煤粉, null, true);
            AddNumericTag("本日.扣返矿.喷煤粉", "", dec_喷煤粉, null, true);
            AddNumericTag("本日.喷煤粉", "本日.1高炉.喷煤粉+本日.4高炉.喷煤粉+本日.5高炉.喷煤粉+本日.6高炉.喷煤粉+本日.7高炉.喷煤粉+本日.8高炉.喷煤粉-本日.扣返矿.喷煤粉", dec_喷煤粉, null, true);

            AddNumericTag("本日库存.1高炉.喷煤粉", "", dec_喷煤粉, null, true);
            AddNumericTag("本日库存.二车间.喷煤粉", "", dec_喷煤粉, null, true);
            AddNumericTag("本日库存.三车间.喷煤粉", "", dec_喷煤粉, null, true);
            AddNumericTag("本日库存.喷煤粉", "本日库存.1高炉.喷煤粉+本日库存.二车间.喷煤粉+本日库存.三车间.喷煤粉", dec_喷煤粉, null, true);

            AddNumericTag("累计.供应.喷煤粉", "昨日累计.供应.喷煤粉+本日.供应.喷煤粉", dec_喷煤粉, null, true);
            AddNumericTag("累计.块矿面.喷煤粉", "昨日累计.块矿面.喷煤粉+本日.块矿面.喷煤粉", dec_喷煤粉, null, true);
            AddNumericTag("累计.1高炉.喷煤粉", "昨日累计.1高炉.喷煤粉+本日.1高炉.喷煤粉", dec_喷煤粉, null, true);
            AddNumericTag("累计.4高炉.喷煤粉", "昨日累计.4高炉.喷煤粉+本日.4高炉.喷煤粉", dec_喷煤粉, null, true);
            AddNumericTag("累计.5高炉.喷煤粉", "昨日累计.5高炉.喷煤粉+本日.5高炉.喷煤粉", dec_喷煤粉, null, true);
            AddNumericTag("累计.6高炉.喷煤粉", "昨日累计.6高炉.喷煤粉+本日.6高炉.喷煤粉", dec_喷煤粉, null, true);
            AddNumericTag("累计.7高炉.喷煤粉", "昨日累计.7高炉.喷煤粉+本日.7高炉.喷煤粉", dec_喷煤粉, null, true);
            AddNumericTag("累计.8高炉.喷煤粉", "昨日累计.8高炉.喷煤粉+本日.8高炉.喷煤粉", dec_喷煤粉, null, true);
            AddNumericTag("累计.扣返矿.喷煤粉", "昨日累计.扣返矿.喷煤粉-本日.扣返矿.喷煤粉", dec_喷煤粉, null, true);
            AddNumericTag("累计.喷煤粉", "累计.1高炉.喷煤粉+累计.4高炉.喷煤粉+累计.5高炉.喷煤粉+累计.6高炉.喷煤粉+累计.7高炉.喷煤粉+累计.8高炉.喷煤粉-累计.扣返矿.喷煤粉", dec_喷煤粉, null, false);






            AddNumericTag("昨日累计.供应.喷煤粉", "", dec_喷煤粉, null, false);
            AddNumericTag("昨日累计.块矿面.喷煤粉", "", dec_喷煤粉, null, false);
            AddNumericTag("昨日累计.1高炉.喷煤粉", "", dec_喷煤粉, null, false);
            AddNumericTag("昨日累计.4高炉.喷煤粉", "", dec_喷煤粉, null, false);
            AddNumericTag("昨日累计.5高炉.喷煤粉", "", dec_喷煤粉, null, false);
            AddNumericTag("昨日累计.6高炉.喷煤粉", "", dec_喷煤粉, null, false);
            AddNumericTag("昨日累计.7高炉.喷煤粉", "", dec_喷煤粉, null, false);
            AddNumericTag("昨日累计.8高炉.喷煤粉", "", dec_喷煤粉, null, false);
            AddNumericTag("昨日累计.扣返矿.喷煤粉", "", dec_喷煤粉, null, false);




            #endregion

            #region 烟煤粉
            int? dec_烟煤粉 = 2;
            AddNumericTag("本日.供应.烟煤粉", "", dec_烟煤粉, null, true);
            AddNumericTag("本日.块矿面.烟煤粉", "", dec_烟煤粉, null, true);
            AddNumericTag("本日.1高炉.烟煤粉", "", dec_烟煤粉, null, true);
            AddNumericTag("本日.4高炉.烟煤粉", "", dec_烟煤粉, null, true);
            AddNumericTag("本日.5高炉.烟煤粉", "", dec_烟煤粉, null, true);
            AddNumericTag("本日.6高炉.烟煤粉", "", dec_烟煤粉, null, true);
            AddNumericTag("本日.7高炉.烟煤粉", "", dec_烟煤粉, null, true);
            AddNumericTag("本日.8高炉.烟煤粉", "", dec_烟煤粉, null, true);
            AddNumericTag("本日.扣返矿.烟煤粉", "", dec_烟煤粉, null, true);
            AddNumericTag("本日.烟煤粉", "本日.1高炉.烟煤粉+本日.4高炉.烟煤粉+本日.5高炉.烟煤粉+本日.6高炉.烟煤粉+本日.7高炉.烟煤粉+本日.8高炉.烟煤粉-本日.扣返矿.烟煤粉", dec_烟煤粉, null, true);

            AddNumericTag("本日库存.1高炉.烟煤粉", "", dec_烟煤粉, null, true);
            AddNumericTag("本日库存.二车间.烟煤粉", "", dec_烟煤粉, null, true);
            AddNumericTag("本日库存.三车间.烟煤粉", "", dec_烟煤粉, null, true);
            AddNumericTag("本日库存.烟煤粉", "本日库存.1高炉.烟煤粉+本日库存.二车间.烟煤粉+本日库存.三车间.烟煤粉", dec_烟煤粉, null, true);

            AddNumericTag("累计.供应.烟煤粉", "昨日累计.供应.烟煤粉+本日.供应.烟煤粉", dec_烟煤粉, null, true);
            AddNumericTag("累计.块矿面.烟煤粉", "昨日累计.块矿面.烟煤粉+本日.块矿面.烟煤粉", dec_烟煤粉, null, true);
            AddNumericTag("累计.1高炉.烟煤粉", "昨日累计.1高炉.烟煤粉+本日.1高炉.烟煤粉", dec_烟煤粉, null, true);
            AddNumericTag("累计.4高炉.烟煤粉", "昨日累计.4高炉.烟煤粉+本日.4高炉.烟煤粉", dec_烟煤粉, null, true);
            AddNumericTag("累计.5高炉.烟煤粉", "昨日累计.5高炉.烟煤粉+本日.5高炉.烟煤粉", dec_烟煤粉, null, true);
            AddNumericTag("累计.6高炉.烟煤粉", "昨日累计.6高炉.烟煤粉+本日.6高炉.烟煤粉", dec_烟煤粉, null, true);
            AddNumericTag("累计.7高炉.烟煤粉", "昨日累计.7高炉.烟煤粉+本日.7高炉.烟煤粉", dec_烟煤粉, null, true);
            AddNumericTag("累计.8高炉.烟煤粉", "昨日累计.8高炉.烟煤粉+本日.8高炉.烟煤粉", dec_烟煤粉, null, true);
            AddNumericTag("累计.扣返矿.烟煤粉", "昨日累计.扣返矿.烟煤粉-本日.扣返矿.烟煤粉", dec_烟煤粉, null, true);
            AddNumericTag("累计.烟煤粉", "累计.1高炉.烟煤粉+累计.4高炉.烟煤粉+累计.5高炉.烟煤粉+累计.6高炉.烟煤粉+累计.7高炉.烟煤粉+累计.8高炉.烟煤粉-累计.扣返矿.烟煤粉", dec_烟煤粉, null, false);






            AddNumericTag("昨日累计.供应.烟煤粉", "", dec_烟煤粉, null, false);
            AddNumericTag("昨日累计.块矿面.烟煤粉", "", dec_烟煤粉, null, false);
            AddNumericTag("昨日累计.1高炉.烟煤粉", "", dec_烟煤粉, null, false);
            AddNumericTag("昨日累计.4高炉.烟煤粉", "", dec_烟煤粉, null, false);
            AddNumericTag("昨日累计.5高炉.烟煤粉", "", dec_烟煤粉, null, false);
            AddNumericTag("昨日累计.6高炉.烟煤粉", "", dec_烟煤粉, null, false);
            AddNumericTag("昨日累计.7高炉.烟煤粉", "", dec_烟煤粉, null, false);
            AddNumericTag("昨日累计.8高炉.烟煤粉", "", dec_烟煤粉, null, false);
            AddNumericTag("昨日累计.扣返矿.烟煤粉", "", dec_烟煤粉, null, false);




            #endregion

            #region 外购小焦
            int? dec_外购小焦 = 2;
            AddNumericTag("本日.供应.外购小焦", "", dec_外购小焦, null, true);
            AddNumericTag("本日.块矿面.外购小焦", "", dec_外购小焦, null, true);
            AddNumericTag("本日.1高炉.外购小焦", "", dec_外购小焦, null, true);
            AddNumericTag("本日.4高炉.外购小焦", "", dec_外购小焦, null, true);
            AddNumericTag("本日.5高炉.外购小焦", "", dec_外购小焦, null, true);
            AddNumericTag("本日.6高炉.外购小焦", "", dec_外购小焦, null, true);
            AddNumericTag("本日.7高炉.外购小焦", "", dec_外购小焦, null, true);
            AddNumericTag("本日.8高炉.外购小焦", "", dec_外购小焦, null, true);
            AddNumericTag("本日.扣返矿.外购小焦", "", dec_外购小焦, null, true);
            AddNumericTag("本日.外购小焦", "本日.1高炉.外购小焦+本日.4高炉.外购小焦+本日.5高炉.外购小焦+本日.6高炉.外购小焦+本日.7高炉.外购小焦+本日.8高炉.外购小焦-本日.扣返矿.外购小焦", dec_外购小焦, null, true);

            AddNumericTag("本日库存.1高炉.外购小焦", "", dec_外购小焦, null, true);
            AddNumericTag("本日库存.二车间.外购小焦", "", dec_外购小焦, null, true);
            AddNumericTag("本日库存.三车间.外购小焦", "", dec_外购小焦, null, true);
            AddNumericTag("本日库存.外购小焦", "本日库存.1高炉.外购小焦+本日库存.二车间.外购小焦+本日库存.三车间.外购小焦", dec_外购小焦, null, true);

            AddNumericTag("累计.供应.外购小焦", "昨日累计.供应.外购小焦+本日.供应.外购小焦", dec_外购小焦, null, true);
            AddNumericTag("累计.块矿面.外购小焦", "昨日累计.块矿面.外购小焦+本日.块矿面.外购小焦", dec_外购小焦, null, true);
            AddNumericTag("累计.1高炉.外购小焦", "昨日累计.1高炉.外购小焦+本日.1高炉.外购小焦", dec_外购小焦, null, true);
            AddNumericTag("累计.4高炉.外购小焦", "昨日累计.4高炉.外购小焦+本日.4高炉.外购小焦", dec_外购小焦, null, true);
            AddNumericTag("累计.5高炉.外购小焦", "昨日累计.5高炉.外购小焦+本日.5高炉.外购小焦", dec_外购小焦, null, true);
            AddNumericTag("累计.6高炉.外购小焦", "昨日累计.6高炉.外购小焦+本日.6高炉.外购小焦", dec_外购小焦, null, true);
            AddNumericTag("累计.7高炉.外购小焦", "昨日累计.7高炉.外购小焦+本日.7高炉.外购小焦", dec_外购小焦, null, true);
            AddNumericTag("累计.8高炉.外购小焦", "昨日累计.8高炉.外购小焦+本日.8高炉.外购小焦", dec_外购小焦, null, true);
            AddNumericTag("累计.扣返矿.外购小焦", "昨日累计.扣返矿.外购小焦-本日.扣返矿.外购小焦", dec_外购小焦, null, true);
            AddNumericTag("累计.外购小焦", "累计.1高炉.外购小焦+累计.4高炉.外购小焦+累计.5高炉.外购小焦+累计.6高炉.外购小焦+累计.7高炉.外购小焦+累计.8高炉.外购小焦-累计.扣返矿.外购小焦", dec_外购小焦, null, false);






            AddNumericTag("昨日累计.供应.外购小焦", "", dec_外购小焦, null, false);
            AddNumericTag("昨日累计.块矿面.外购小焦", "", dec_外购小焦, null, false);
            AddNumericTag("昨日累计.1高炉.外购小焦", "", dec_外购小焦, null, false);
            AddNumericTag("昨日累计.4高炉.外购小焦", "", dec_外购小焦, null, false);
            AddNumericTag("昨日累计.5高炉.外购小焦", "", dec_外购小焦, null, false);
            AddNumericTag("昨日累计.6高炉.外购小焦", "", dec_外购小焦, null, false);
            AddNumericTag("昨日累计.7高炉.外购小焦", "", dec_外购小焦, null, false);
            AddNumericTag("昨日累计.8高炉.外购小焦", "", dec_外购小焦, null, false);
            AddNumericTag("昨日累计.扣返矿.外购小焦", "", dec_外购小焦, null, false);




            #endregion

            #region 锰矿
            int? dec_锰矿 = 2;
            AddNumericTag("本日.供应.锰矿", "", dec_锰矿, null, true);
            AddNumericTag("本日.块矿面.锰矿", "", dec_锰矿, null, true);
            AddNumericTag("本日.1高炉.锰矿", "", dec_锰矿, null, true);
            AddNumericTag("本日.4高炉.锰矿", "", dec_锰矿, null, true);
            AddNumericTag("本日.5高炉.锰矿", "", dec_锰矿, null, true);
            AddNumericTag("本日.6高炉.锰矿", "", dec_锰矿, null, true);
            AddNumericTag("本日.7高炉.锰矿", "", dec_锰矿, null, true);
            AddNumericTag("本日.8高炉.锰矿", "", dec_锰矿, null, true);
            AddNumericTag("本日.扣返矿.锰矿", "", dec_锰矿, null, true);
            AddNumericTag("本日.锰矿", "本日.1高炉.锰矿+本日.4高炉.锰矿+本日.5高炉.锰矿+本日.6高炉.锰矿+本日.7高炉.锰矿+本日.8高炉.锰矿-本日.扣返矿.锰矿", dec_锰矿, null, true);

            AddNumericTag("本日库存.1高炉.锰矿", "", dec_锰矿, null, true);
            AddNumericTag("本日库存.二车间.锰矿", "", dec_锰矿, null, true);
            AddNumericTag("本日库存.三车间.锰矿", "", dec_锰矿, null, true);
            AddNumericTag("本日库存.锰矿", "本日库存.1高炉.锰矿+本日库存.二车间.锰矿+本日库存.三车间.锰矿", dec_锰矿, null, true);

            AddNumericTag("累计.供应.锰矿", "昨日累计.供应.锰矿+本日.供应.锰矿", dec_锰矿, null, true);
            AddNumericTag("累计.块矿面.锰矿", "昨日累计.块矿面.锰矿+本日.块矿面.锰矿", dec_锰矿, null, true);
            AddNumericTag("累计.1高炉.锰矿", "昨日累计.1高炉.锰矿+本日.1高炉.锰矿", dec_锰矿, null, true);
            AddNumericTag("累计.4高炉.锰矿", "昨日累计.4高炉.锰矿+本日.4高炉.锰矿", dec_锰矿, null, true);
            AddNumericTag("累计.5高炉.锰矿", "昨日累计.5高炉.锰矿+本日.5高炉.锰矿", dec_锰矿, null, true);
            AddNumericTag("累计.6高炉.锰矿", "昨日累计.6高炉.锰矿+本日.6高炉.锰矿", dec_锰矿, null, true);
            AddNumericTag("累计.7高炉.锰矿", "昨日累计.7高炉.锰矿+本日.7高炉.锰矿", dec_锰矿, null, true);
            AddNumericTag("累计.8高炉.锰矿", "昨日累计.8高炉.锰矿+本日.8高炉.锰矿", dec_锰矿, null, true);
            AddNumericTag("累计.扣返矿.锰矿", "昨日累计.扣返矿.锰矿-本日.扣返矿.锰矿", dec_锰矿, null, true);
            AddNumericTag("累计.锰矿", "累计.1高炉.锰矿+累计.4高炉.锰矿+累计.5高炉.锰矿+累计.6高炉.锰矿+累计.7高炉.锰矿+累计.8高炉.锰矿-累计.扣返矿.锰矿", dec_锰矿, null, false);

            AddNumericTag("昨日累计.供应.锰矿", "", dec_锰矿, null, false);
            AddNumericTag("昨日累计.块矿面.锰矿", "", dec_锰矿, null, false);
            AddNumericTag("昨日累计.1高炉.锰矿", "", dec_锰矿, null, false);
            AddNumericTag("昨日累计.4高炉.锰矿", "", dec_锰矿, null, false);
            AddNumericTag("昨日累计.5高炉.锰矿", "", dec_锰矿, null, false);
            AddNumericTag("昨日累计.6高炉.锰矿", "", dec_锰矿, null, false);
            AddNumericTag("昨日累计.7高炉.锰矿", "", dec_锰矿, null, false);
            AddNumericTag("昨日累计.8高炉.锰矿", "", dec_锰矿, null, false);
            AddNumericTag("昨日累计.扣返矿.锰矿", "", dec_锰矿, null, false);

            #endregion

            #region 萤石
            int? dec_萤石 = 2;
            AddNumericTag("本日.供应.萤石", "", dec_萤石, null, true);
            AddNumericTag("本日.块矿面.萤石", "", dec_萤石, null, true);
            AddNumericTag("本日.1高炉.萤石", "", dec_萤石, null, true);
            AddNumericTag("本日.4高炉.萤石", "", dec_萤石, null, true);
            AddNumericTag("本日.5高炉.萤石", "", dec_萤石, null, true);
            AddNumericTag("本日.6高炉.萤石", "", dec_萤石, null, true);
            AddNumericTag("本日.7高炉.萤石", "", dec_萤石, null, true);
            AddNumericTag("本日.8高炉.萤石", "", dec_萤石, null, true);
            AddNumericTag("本日.扣返矿.萤石", "", dec_萤石, null, true);
            AddNumericTag("本日.萤石", "本日.1高炉.萤石+本日.4高炉.萤石+本日.5高炉.萤石+本日.6高炉.萤石+本日.7高炉.萤石+本日.8高炉.萤石-本日.扣返矿.萤石", dec_萤石, null, true);

            AddNumericTag("本日库存.1高炉.萤石", "", dec_萤石, null, true);
            AddNumericTag("本日库存.二车间.萤石", "", dec_萤石, null, true);
            AddNumericTag("本日库存.三车间.萤石", "", dec_萤石, null, true);
            AddNumericTag("本日库存.萤石", "本日库存.1高炉.萤石+本日库存.二车间.萤石+本日库存.三车间.萤石", dec_萤石, null, true);

            AddNumericTag("累计.供应.萤石", "昨日累计.供应.萤石+本日.供应.萤石", dec_萤石, null, true);
            AddNumericTag("累计.块矿面.萤石", "昨日累计.块矿面.萤石+本日.块矿面.萤石", dec_萤石, null, true);
            AddNumericTag("累计.1高炉.萤石", "昨日累计.1高炉.萤石+本日.1高炉.萤石", dec_萤石, null, true);
            AddNumericTag("累计.4高炉.萤石", "昨日累计.4高炉.萤石+本日.4高炉.萤石", dec_萤石, null, true);
            AddNumericTag("累计.5高炉.萤石", "昨日累计.5高炉.萤石+本日.5高炉.萤石", dec_萤石, null, true);
            AddNumericTag("累计.6高炉.萤石", "昨日累计.6高炉.萤石+本日.6高炉.萤石", dec_萤石, null, true);
            AddNumericTag("累计.7高炉.萤石", "昨日累计.7高炉.萤石+本日.7高炉.萤石", dec_萤石, null, true);
            AddNumericTag("累计.8高炉.萤石", "昨日累计.8高炉.萤石+本日.8高炉.萤石", dec_萤石, null, true);
            AddNumericTag("累计.扣返矿.萤石", "昨日累计.扣返矿.萤石-本日.扣返矿.萤石", dec_萤石, null, true);
            AddNumericTag("累计.萤石", "累计.1高炉.萤石+累计.4高炉.萤石+累计.5高炉.萤石+累计.6高炉.萤石+累计.7高炉.萤石+累计.8高炉.萤石-累计.扣返矿.萤石", dec_萤石, null, false);


            AddNumericTag("昨日累计.供应.萤石", "", dec_萤石, null, false);
            AddNumericTag("昨日累计.块矿面.萤石", "", dec_萤石, null, false);
            AddNumericTag("昨日累计.1高炉.萤石", "", dec_萤石, null, false);
            AddNumericTag("昨日累计.4高炉.萤石", "", dec_萤石, null, false);
            AddNumericTag("昨日累计.5高炉.萤石", "", dec_萤石, null, false);
            AddNumericTag("昨日累计.6高炉.萤石", "", dec_萤石, null, false);
            AddNumericTag("昨日累计.7高炉.萤石", "", dec_萤石, null, false);
            AddNumericTag("昨日累计.8高炉.萤石", "", dec_萤石, null, false);
            AddNumericTag("昨日累计.扣返矿.萤石", "", dec_萤石, null, false);

            #endregion

            #region 硅石石子
            int? dec_硅石石子 = 2;
            AddNumericTag("本日.供应.硅石石子", "", dec_硅石石子, null, true);
            AddNumericTag("本日.块矿面.硅石石子", "", dec_硅石石子, null, true);
            AddNumericTag("本日.1高炉.硅石石子", "", dec_硅石石子, null, true);
            AddNumericTag("本日.4高炉.硅石石子", "", dec_硅石石子, null, true);
            AddNumericTag("本日.5高炉.硅石石子", "", dec_硅石石子, null, true);
            AddNumericTag("本日.6高炉.硅石石子", "", dec_硅石石子, null, true);
            AddNumericTag("本日.7高炉.硅石石子", "", dec_硅石石子, null, true);
            AddNumericTag("本日.8高炉.硅石石子", "", dec_硅石石子, null, true);
            AddNumericTag("本日.扣返矿.硅石石子", "", dec_硅石石子, null, true);
            AddNumericTag("本日.硅石石子", "本日.1高炉.硅石石子+本日.4高炉.硅石石子+本日.5高炉.硅石石子+本日.6高炉.硅石石子+本日.7高炉.硅石石子+本日.8高炉.硅石石子-本日.扣返矿.硅石石子", dec_硅石石子, null, true);

            AddNumericTag("本日库存.1高炉.硅石石子", "", dec_硅石石子, null, true);
            AddNumericTag("本日库存.二车间.硅石石子", "", dec_硅石石子, null, true);
            AddNumericTag("本日库存.三车间.硅石石子", "", dec_硅石石子, null, true);
            AddNumericTag("本日库存.硅石石子", "本日库存.1高炉.硅石石子+本日库存.二车间.硅石石子+本日库存.三车间.硅石石子", dec_硅石石子, null, true);

            AddNumericTag("累计.供应.硅石石子", "昨日累计.供应.硅石石子+本日.供应.硅石石子", dec_硅石石子, null, true);
            AddNumericTag("累计.块矿面.硅石石子", "昨日累计.块矿面.硅石石子+本日.块矿面.硅石石子", dec_硅石石子, null, true);
            AddNumericTag("累计.1高炉.硅石石子", "昨日累计.1高炉.硅石石子+本日.1高炉.硅石石子", dec_硅石石子, null, true);
            AddNumericTag("累计.4高炉.硅石石子", "昨日累计.4高炉.硅石石子+本日.4高炉.硅石石子", dec_硅石石子, null, true);
            AddNumericTag("累计.5高炉.硅石石子", "昨日累计.5高炉.硅石石子+本日.5高炉.硅石石子", dec_硅石石子, null, true);
            AddNumericTag("累计.6高炉.硅石石子", "昨日累计.6高炉.硅石石子+本日.6高炉.硅石石子", dec_硅石石子, null, true);
            AddNumericTag("累计.7高炉.硅石石子", "昨日累计.7高炉.硅石石子+本日.7高炉.硅石石子", dec_硅石石子, null, true);
            AddNumericTag("累计.8高炉.硅石石子", "昨日累计.8高炉.硅石石子+本日.8高炉.硅石石子", dec_硅石石子, null, true);
            AddNumericTag("累计.扣返矿.硅石石子", "昨日累计.扣返矿.硅石石子-本日.扣返矿.硅石石子", dec_硅石石子, null, true);
            AddNumericTag("累计.硅石石子", "累计.1高炉.硅石石子+累计.4高炉.硅石石子+累计.5高炉.硅石石子+累计.6高炉.硅石石子+累计.7高炉.硅石石子+累计.8高炉.硅石石子-累计.扣返矿.硅石石子", dec_硅石石子, null, false);

            AddNumericTag("昨日累计.供应.硅石石子", "", dec_硅石石子, null, false);
            AddNumericTag("昨日累计.块矿面.硅石石子", "", dec_硅石石子, null, false);
            AddNumericTag("昨日累计.1高炉.硅石石子", "", dec_硅石石子, null, false);
            AddNumericTag("昨日累计.4高炉.硅石石子", "", dec_硅石石子, null, false);
            AddNumericTag("昨日累计.5高炉.硅石石子", "", dec_硅石石子, null, false);
            AddNumericTag("昨日累计.6高炉.硅石石子", "", dec_硅石石子, null, false);
            AddNumericTag("昨日累计.7高炉.硅石石子", "", dec_硅石石子, null, false);
            AddNumericTag("昨日累计.8高炉.硅石石子", "", dec_硅石石子, null, false);
            AddNumericTag("昨日累计.扣返矿.硅石石子", "", dec_硅石石子, null, false);

            #endregion

            #region 焦炭入炉分析
            AddNumericTag("焦炭入炉分析.78号炉1.A", "", null, null, true);
            AddNumericTag("焦炭入炉分析.78号炉1.C固", "", null, null, true);
            AddNumericTag("焦炭入炉分析.78号炉1.H2O", "", null, null, true);
            AddNumericTag("焦炭入炉分析.78号炉1.S", "", null, null, true);
            AddNumericTag("焦炭入炉分析.78号炉1.V", "", null, null, true);
            AddNumericTag("焦炭入炉分析.78号炉2.A", "", null, null, true);
            AddNumericTag("焦炭入炉分析.78号炉2.C固", "", null, null, true);
            AddNumericTag("焦炭入炉分析.78号炉2.H2O", "", null, null, true);
            AddNumericTag("焦炭入炉分析.78号炉2.S", "", null, null, true);
            AddNumericTag("焦炭入炉分析.78号炉2.V", "", null, null, true);
            AddNumericTag("焦炭入炉分析.二车间1.A", "", null, null, true);
            AddNumericTag("焦炭入炉分析.二车间1.C固", "", null, null, true);
            AddNumericTag("焦炭入炉分析.二车间1.H2O", "", null, null, true);
            AddNumericTag("焦炭入炉分析.二车间1.S", "", null, null, true);
            AddNumericTag("焦炭入炉分析.二车间1.V", "", null, null, true);
            AddNumericTag("焦炭入炉分析.二车间2.A", "", null, null, true);
            AddNumericTag("焦炭入炉分析.二车间2.C固", "", null, null, true);
            AddNumericTag("焦炭入炉分析.二车间2.H2O", "", null, null, true);
            AddNumericTag("焦炭入炉分析.二车间2.S", "", null, null, true);
            AddNumericTag("焦炭入炉分析.二车间2.V", "", null, null, true);
            AddNumericTag("焦炭入炉分析.全厂合计1.A", "", null, null, true);
            AddNumericTag("焦炭入炉分析.全厂合计1.C固", "", null, null, true);
            AddNumericTag("焦炭入炉分析.全厂合计1.H2O", "", null, null, true);
            AddNumericTag("焦炭入炉分析.全厂合计1.S", "", null, null, true);
            AddNumericTag("焦炭入炉分析.全厂合计1.V", "", null, null, true);
            AddNumericTag("焦炭入炉分析.全厂合计2.A", "", null, null, true);
            AddNumericTag("焦炭入炉分析.全厂合计2.C固", "", null, null, true);
            AddNumericTag("焦炭入炉分析.全厂合计2.H2O", "", null, null, true);
            AddNumericTag("焦炭入炉分析.全厂合计2.S", "", null, null, true);
            AddNumericTag("焦炭入炉分析.全厂合计2.V", "", null, null, true);
            AddNumericTag("焦炭入炉分析.1高炉1.A", "", null, null, true);
            AddNumericTag("焦炭入炉分析.1高炉1.C固", "", null, null, true);
            AddNumericTag("焦炭入炉分析.1高炉1.H2O", "", null, null, true);
            AddNumericTag("焦炭入炉分析.1高炉1.S", "", null, null, true);
            AddNumericTag("焦炭入炉分析.1高炉1.V", "", null, null, true);
            AddNumericTag("焦炭入炉分析.1高炉2.A", "", null, null, true);
            AddNumericTag("焦炭入炉分析.1高炉2.C固", "", null, null, true);
            AddNumericTag("焦炭入炉分析.1高炉2.H2O", "", null, null, true);
            AddNumericTag("焦炭入炉分析.1高炉2.S", "", null, null, true);
            AddNumericTag("焦炭入炉分析.1高炉2.V", "", null, null, true);
            #endregion

            #region 喷煤化验分析
            AddNumericTag("喷煤化验分析.质量1.A", "", null, null, true);
            AddNumericTag("喷煤化验分析.质量1.C固", "", null, null, true);
            AddNumericTag("喷煤化验分析.质量1.S", "", null, null, true);
            AddNumericTag("喷煤化验分析.质量1.V", "", null, null, true);
            AddNumericTag("喷煤化验分析.质量2.A", "", null, null, true);
            AddNumericTag("喷煤化验分析.质量2.C固", "", null, null, true);
            AddNumericTag("喷煤化验分析.质量2.S", "", null, null, true);
            AddNumericTag("喷煤化验分析.质量2.V", "", null, null, true);
            #endregion

            #region 入炉矿化验分析
            AddNumericTag("入炉矿化验分析.78号炉1.FeO", "", null, null, true);
            AddNumericTag("入炉矿化验分析.78号炉1.R", "", null, null, true);
            AddNumericTag("入炉矿化验分析.78号炉1.S", "", null, null, true);
            AddNumericTag("入炉矿化验分析.78号炉1.TFe", "", null, null, true);
            AddNumericTag("入炉矿化验分析.78号炉1.块矿品位", "", null, null, true);
            AddNumericTag("入炉矿化验分析.78号炉2.FeO", "", null, null, true);
            AddNumericTag("入炉矿化验分析.78号炉2.R", "", null, null, true);
            AddNumericTag("入炉矿化验分析.78号炉2.S", "", null, null, true);
            AddNumericTag("入炉矿化验分析.78号炉2.TFe", "", null, null, true);
            AddNumericTag("入炉矿化验分析.78号炉2.块矿品位", "", null, null, true);
            AddNumericTag("入炉矿化验分析.二车间1.FeO", "", null, null, true);
            AddNumericTag("入炉矿化验分析.二车间1.R", "", null, null, true);
            AddNumericTag("入炉矿化验分析.二车间1.S", "", null, null, true);
            AddNumericTag("入炉矿化验分析.二车间1.TFe", "", null, null, true);
            AddNumericTag("入炉矿化验分析.二车间1.块矿品位", "", null, null, true);
            AddNumericTag("入炉矿化验分析.二车间2.FeO", "", null, null, true);
            AddNumericTag("入炉矿化验分析.二车间2.R", "", null, null, true);
            AddNumericTag("入炉矿化验分析.二车间2.S", "", null, null, true);
            AddNumericTag("入炉矿化验分析.二车间2.TFe", "", null, null, true);
            AddNumericTag("入炉矿化验分析.二车间2.块矿品位", "", null, null, true);
            AddNumericTag("入炉矿化验分析.全厂合计1.FeO", "", null, null, true);
            AddNumericTag("入炉矿化验分析.全厂合计1.R", "", null, null, true);
            AddNumericTag("入炉矿化验分析.全厂合计1.S", "", null, null, true);
            AddNumericTag("入炉矿化验分析.全厂合计1.TFe", "", null, null, true);
            AddNumericTag("入炉矿化验分析.全厂合计1.块矿品位", "", null, null, true);
            AddNumericTag("入炉矿化验分析.全厂合计2.FeO", "", null, null, true);
            AddNumericTag("入炉矿化验分析.全厂合计2.R", "", null, null, true);
            AddNumericTag("入炉矿化验分析.全厂合计2.S", "", null, null, true);
            AddNumericTag("入炉矿化验分析.全厂合计2.TFe", "", null, null, true);
            AddNumericTag("入炉矿化验分析.全厂合计2.块矿品位", "", null, null, true);
            AddNumericTag("入炉矿化验分析.1高炉1.FeO", "", null, null, true);
            AddNumericTag("入炉矿化验分析.1高炉1.R", "", null, null, true);
            AddNumericTag("入炉矿化验分析.1高炉1.S", "", null, null, true);
            AddNumericTag("入炉矿化验分析.1高炉1.TFe", "", null, null, true);
            AddNumericTag("入炉矿化验分析.1高炉1.块矿品位", "", null, null, true);
            AddNumericTag("入炉矿化验分析.1高炉2.FeO", "", null, null, true);
            AddNumericTag("入炉矿化验分析.1高炉2.R", "", null, null, true);
            AddNumericTag("入炉矿化验分析.1高炉2.S", "", null, null, true);
            AddNumericTag("入炉矿化验分析.1高炉2.TFe", "", null, null, true);
            AddNumericTag("入炉矿化验分析.1高炉2.块矿品位", "", null, null, true);
            #endregion

            #region 烧结矿
            AddNumericTag("本日.烧结矿.230烧结.1号炉", "", null, null, true);
            AddNumericTag("本日.烧结矿.230烧结.二车间", "", null, null, true);
            AddNumericTag("本日.烧结矿.230烧结.三车间", "", null, null, true);
            AddNumericTag("本日.烧结矿.230烧结.合计", "本日.烧结矿.230烧结.1号炉+本日.烧结矿.230烧结.二车间+本日.烧结矿.230烧结.三车间", null, null, false);

            AddNumericTag("累计.烧结矿.230烧结.1号炉", "昨日累计.烧结矿.230烧结.1号炉+本日.烧结矿.230烧结.1号炉", null, null, true);
            AddNumericTag("累计.烧结矿.230烧结.二车间", "昨日累计.烧结矿.230烧结.二车间+本日.烧结矿.230烧结.二车间", null, null, true);
            AddNumericTag("累计.烧结矿.230烧结.三车间", "昨日累计.烧结矿.230烧结.三车间+本日.烧结矿.230烧结.三车间", null, null, true);
            AddNumericTag("累计.烧结矿.230烧结.合计", "累计.烧结矿.230烧结.1号炉+累计.烧结矿.230烧结.二车间+累计.烧结矿.230烧结.三车间", null, null, false);



            AddNumericTag("本日.烧结矿.大炼铁烧结.1号炉", "", null, null, true);
            AddNumericTag("本日.烧结矿.大炼铁烧结.二车间", "", null, null, true);
            AddNumericTag("本日.烧结矿.大炼铁烧结.三车间", "", null, null, true);
            AddNumericTag("本日.烧结矿.大炼铁烧结.合计", "本日.烧结矿.大炼铁烧结.1号炉+本日.烧结矿.大炼铁烧结.二车间+本日.烧结矿.大炼铁烧结.三车间", null, null, false);


            AddNumericTag("累计.烧结矿.大炼铁烧结.1号炉", "昨日累计.烧结矿.大炼铁烧结.1号炉+本日.烧结矿.大炼铁烧结.1号炉", null, null, true);
            AddNumericTag("累计.烧结矿.大炼铁烧结.二车间", "昨日累计.烧结矿.大炼铁烧结.二车间+本日.烧结矿.大炼铁烧结.二车间", null, null, true);
            AddNumericTag("累计.烧结矿.大炼铁烧结.三车间", "昨日累计.烧结矿.大炼铁烧结.三车间+本日.烧结矿.大炼铁烧结.三车间", null, null, true);
            AddNumericTag("累计.烧结矿.大炼铁烧结.合计", "累计.烧结矿.大炼铁烧结.1号炉+累计.烧结矿.大炼铁烧结.二车间+累计.烧结矿.大炼铁烧结.三车间", null, null, false);

            AddNumericTag("本日.烧结矿", "本日.烧结矿.230烧结.合计+本日.烧结矿.大炼铁烧结.合计", null, null, false);
            AddNumericTag("累计.烧结矿", "累计.烧结矿.230烧结.合计+累计.烧结矿.大炼铁烧结.合计", null, null, false);

            AddNumericTag("昨日累计.烧结矿.230烧结.1号炉", "", null, null, true);
            AddNumericTag("昨日累计.烧结矿.230烧结.二车间", "", null, null, true);
            AddNumericTag("昨日累计.烧结矿.230烧结.三车间", "", null, null, true);

            AddNumericTag("昨日累计.烧结矿.大炼铁烧结.1号炉", "", null, null, true);
            AddNumericTag("昨日累计.烧结矿.大炼铁烧结.二车间", "", null, null, true);
            AddNumericTag("昨日累计.烧结矿.大炼铁烧结.三车间", "", null, null, true);
            #endregion

            #region 外购焦炭质量
            AddNumericTag("外购焦炭质量.质量1.A", "", null, null, true);
            AddNumericTag("外购焦炭质量.质量1.C固", "", null, null, true);
            AddNumericTag("外购焦炭质量.质量1.H2O", "", null, null, true);
            AddNumericTag("外购焦炭质量.质量1.M10", "", null, null, true);
            AddNumericTag("外购焦炭质量.质量1.M25", "", null, null, true);
            AddNumericTag("外购焦炭质量.质量1.S", "", null, null, true);
            AddNumericTag("外购焦炭质量.质量1.V", "", null, null, true);
            AddNumericTag("外购焦炭质量.质量2.A", "", null, null, true);
            AddNumericTag("外购焦炭质量.质量2.C固", "", null, null, true);
            AddNumericTag("外购焦炭质量.质量2.H2O", "", null, null, true);
            AddNumericTag("外购焦炭质量.质量2.M10", "", null, null, true);
            AddNumericTag("外购焦炭质量.质量2.M25", "", null, null, true);
            AddNumericTag("外购焦炭质量.质量2.S", "", null, null, true);
            AddNumericTag("外购焦炭质量.质量2.V", "", null, null, true);
            #endregion

            #region 新焦化焦炭质量
            AddNumericTag("新焦化焦炭质量.质量1.A", "", null, null, true);
            AddNumericTag("新焦化焦炭质量.质量1.C固", "", null, null, true);
            AddNumericTag("新焦化焦炭质量.质量1.H2O", "", null, null, true);
            AddNumericTag("新焦化焦炭质量.质量1.M10", "", null, null, true);
            AddNumericTag("新焦化焦炭质量.质量1.M25", "", null, null, true);
            AddNumericTag("新焦化焦炭质量.质量1.S", "", null, null, true);
            AddNumericTag("新焦化焦炭质量.质量1.V", "", null, null, true);
            AddNumericTag("新焦化焦炭质量.质量2.A", "", null, null, true);
            AddNumericTag("新焦化焦炭质量.质量2.C固", "", null, null, true);
            AddNumericTag("新焦化焦炭质量.质量2.H2O", "", null, null, true);
            AddNumericTag("新焦化焦炭质量.质量2.M10", "", null, null, true);
            AddNumericTag("新焦化焦炭质量.质量2.M25", "", null, null, true);
            AddNumericTag("新焦化焦炭质量.质量2.S", "", null, null, true);
            AddNumericTag("新焦化焦炭质量.质量2.V", "", null, null, true);
            #endregion

            #region 烟煤化验分析
            AddNumericTag("烟煤化验分析.质量1.A", "", null, null, true);
            AddNumericTag("烟煤化验分析.质量1.C固", "", null, null, true);
            AddNumericTag("烟煤化验分析.质量1.S", "", null, null, true);
            AddNumericTag("烟煤化验分析.质量1.V", "", null, null, true);
            AddNumericTag("烟煤化验分析.质量2.A", "", null, null, true);
            AddNumericTag("烟煤化验分析.质量2.C固", "", null, null, true);
            AddNumericTag("烟煤化验分析.质量2.S", "", null, null, true);
            AddNumericTag("烟煤化验分析.质量2.V", "", null, null, true);
            #endregion

        }

        private void initTagsByDb()
        {
            LTZN.Repository.QueryModel calmodel = new LTZN.Repository.QueryModel("技术日报");
            foreach (var item in calmodel.myModel.AllCalTags)
            {
                AddNumericTag(item.TagFullName, item.Forma, item.Dec, null, false);
            }
        }
        /// <summary>
        /// 初始化标签显示位置
        /// </summary>
        private void initRibaoView()
        {
            AddViewItem("ribao", 3, 15, "报表日期");
            AddViewItem("ribao", 60, 3, "备注");

            AddViewItem("ribao", 6, 3, "本日.4高炉.全铁产量");
            AddViewItem("ribao", 6, 4, "本日.4高炉.合格铁");
            AddViewItem("ribao", 6, 5, "本日.4高炉.号外铁");
            AddViewItem("ribao", 6, 6, "本日.4高炉.铸造铁");
            AddViewItem("ribao", 6, 7, "本日.4高炉.利用系数");
            AddViewItem("ribao", 6, 8, "本日.4高炉.原料矿单耗");
            AddViewItem("ribao", 6, 9, "本日.4高炉.干焦丁");
            AddViewItem("ribao", 6, 10, "本日.4高炉.综合焦炭单耗");
            AddViewItem("ribao", 6, 11, "本日.4高炉.入炉焦炭单耗");
            AddViewItem("ribao", 6, 12, "本日.4高炉.煤粉单耗");
            AddViewItem("ribao", 6, 13, "本日.4高炉.焦丁比");
            AddViewItem("ribao", 6, 14, "本日.4高炉.冶炼强度");
            AddViewItem("ribao", 6, 15, "本日.4高炉.合格率");
            AddViewItem("ribao", 6, 16, "本日.4高炉.休风率");
            AddViewItem("ribao", 6, 17, "本日.4高炉.平均风温");
            AddViewItem("ribao", 6, 18, "本日.4高炉.耗电");
            AddViewItem("ribao", 6, 19, "本日.4高炉.富氧率");
            AddViewItem("ribao", 6, 20, "本日.4高炉.生铁含硅");
            AddViewItem("ribao", 6, 21, "本日.4高炉.一级品率");
            AddViewItem("ribao", 6, 22, "本日.4高炉.氧比");
            AddViewItem("ribao", 6, 23, "本日.4高炉.氮比");
            AddViewItem("ribao", 6, 24, "本日.4高炉.焦炭负荷");
            AddViewItem("ribao", 6, 25, "本日.4高炉.燃料比");
            AddViewItem("ribao", 6, 26, "本日.4高炉.废铁单耗");
            AddViewItem("ribao", 6, 30, "本日.4高炉.TRT发电");
            AddViewItem("ribao", 6, 31, "本日.4高炉.返矿");
            AddViewItem("ribao", 6, 32, "本日.4高炉.矿槽");
            AddViewItem("ribao", 6, 33, "本日.4高炉.休风工艺");
            AddViewItem("ribao", 6, 34, "本日.4高炉.休风设备");
            AddViewItem("ribao", 6, 35, "本日.4高炉.休风电");
            AddViewItem("ribao", 6, 36, "本日.4高炉.休风外围");
            AddViewItem("ribao", 6, 37, "本日.4高炉.休风计划");
            AddViewItem("ribao", 6, 38, "本日.4高炉.休风");

            AddViewItem("ribao", 7, 3, "累计.4高炉.全铁产量");
            AddViewItem("ribao", 7, 4, "累计.4高炉.合格铁");
            AddViewItem("ribao", 7, 5, "累计.4高炉.号外铁");
            AddViewItem("ribao", 7, 6, "累计.4高炉.铸造铁");
            AddViewItem("ribao", 7, 7, "累计.4高炉.利用系数");
            AddViewItem("ribao", 7, 8, "累计.4高炉.原料矿单耗");
            AddViewItem("ribao", 7, 9, "累计.4高炉.干焦丁");
            AddViewItem("ribao", 7, 10, "累计.4高炉.综合焦炭单耗");
            AddViewItem("ribao", 7, 11, "累计.4高炉.入炉焦炭单耗");
            AddViewItem("ribao", 7, 12, "累计.4高炉.煤粉单耗");
            AddViewItem("ribao", 7, 13, "累计.4高炉.焦丁比");
            AddViewItem("ribao", 7, 14, "累计.4高炉.冶炼强度");
            AddViewItem("ribao", 7, 15, "累计.4高炉.合格率");
            AddViewItem("ribao", 7, 16, "累计.4高炉.休风率");
            AddViewItem("ribao", 7, 17, "累计.4高炉.平均风温");
            AddViewItem("ribao", 7, 18, "累计.4高炉.电单耗");
            AddViewItem("ribao", 7, 19, "累计.4高炉.富氧率");
            AddViewItem("ribao", 7, 20, "累计.4高炉.生铁含硅");
            AddViewItem("ribao", 7, 21, "累计.4高炉.一级品率");
            AddViewItem("ribao", 7, 22, "累计.4高炉.氧比");
            AddViewItem("ribao", 7, 23, "累计.4高炉.氮比");
            AddViewItem("ribao", 7, 24, "累计.4高炉.焦炭负荷");
            AddViewItem("ribao", 7, 25, "累计.4高炉.燃料比");
            AddViewItem("ribao", 7, 26, "累计.4高炉.废铁单耗");
            AddViewItem("ribao", 7, 30, "累计.4高炉.TRT发电");
            AddViewItem("ribao", 7, 31, "累计.4高炉.返矿");
            AddViewItem("ribao", 7, 32, "累计.4高炉.矿槽");
            AddViewItem("ribao", 7, 33, "累计.4高炉.休风工艺");
            AddViewItem("ribao", 7, 34, "累计.4高炉.休风设备");
            AddViewItem("ribao", 7, 35, "累计.4高炉.休风电");
            AddViewItem("ribao", 7, 36, "累计.4高炉.休风外围");
            AddViewItem("ribao", 7, 37, "累计.4高炉.休风计划");
            AddViewItem("ribao", 7, 38, "累计.4高炉.休风");

            AddViewItem("ribao", 8, 3, "本日.5高炉.全铁产量");
            AddViewItem("ribao", 8, 4, "本日.5高炉.合格铁");
            AddViewItem("ribao", 8, 5, "本日.5高炉.号外铁");
            AddViewItem("ribao", 8, 6, "本日.5高炉.铸造铁");
            AddViewItem("ribao", 8, 7, "本日.5高炉.利用系数");
            AddViewItem("ribao", 8, 8, "本日.5高炉.原料矿单耗");
            AddViewItem("ribao", 8, 9, "本日.5高炉.干焦丁");
            AddViewItem("ribao", 8, 10, "本日.5高炉.综合焦炭单耗");
            AddViewItem("ribao", 8, 11, "本日.5高炉.入炉焦炭单耗");
            AddViewItem("ribao", 8, 12, "本日.5高炉.煤粉单耗");
            AddViewItem("ribao", 8, 13, "本日.5高炉.焦丁比");
            AddViewItem("ribao", 8, 14, "本日.5高炉.冶炼强度");
            AddViewItem("ribao", 8, 15, "本日.5高炉.合格率");
            AddViewItem("ribao", 8, 16, "本日.5高炉.休风率");
            AddViewItem("ribao", 8, 17, "本日.5高炉.平均风温");
            AddViewItem("ribao", 8, 18, "本日.5高炉.耗电");
            AddViewItem("ribao", 8, 19, "本日.5高炉.富氧率");
            AddViewItem("ribao", 8, 20, "本日.5高炉.生铁含硅");
            AddViewItem("ribao", 8, 21, "本日.5高炉.一级品率");
            AddViewItem("ribao", 8, 22, "本日.5高炉.氧比");
            AddViewItem("ribao", 8, 23, "本日.5高炉.氮比");
            AddViewItem("ribao", 8, 24, "本日.5高炉.焦炭负荷");
            AddViewItem("ribao", 8, 25, "本日.5高炉.燃料比");
            AddViewItem("ribao", 8, 26, "本日.5高炉.废铁单耗");
            AddViewItem("ribao", 8, 30, "本日.5高炉.TRT发电");
            AddViewItem("ribao", 8, 31, "本日.5高炉.返矿");
            AddViewItem("ribao", 8, 32, "本日.5高炉.矿槽");
            AddViewItem("ribao", 8, 33, "本日.5高炉.休风工艺");
            AddViewItem("ribao", 8, 34, "本日.5高炉.休风设备");
            AddViewItem("ribao", 8, 35, "本日.5高炉.休风电");
            AddViewItem("ribao", 8, 36, "本日.5高炉.休风外围");
            AddViewItem("ribao", 8, 37, "本日.5高炉.休风计划");
            AddViewItem("ribao", 8, 38, "本日.5高炉.休风");

            AddViewItem("ribao", 9, 3, "累计.5高炉.全铁产量");
            AddViewItem("ribao", 9, 4, "累计.5高炉.合格铁");
            AddViewItem("ribao", 9, 5, "累计.5高炉.号外铁");
            AddViewItem("ribao", 9, 6, "累计.5高炉.铸造铁");
            AddViewItem("ribao", 9, 7, "累计.5高炉.利用系数");
            AddViewItem("ribao", 9, 8, "累计.5高炉.原料矿单耗");
            AddViewItem("ribao", 9, 9, "累计.5高炉.干焦丁");
            AddViewItem("ribao", 9, 10, "累计.5高炉.综合焦炭单耗");
            AddViewItem("ribao", 9, 11, "累计.5高炉.入炉焦炭单耗");
            AddViewItem("ribao", 9, 12, "累计.5高炉.煤粉单耗");
            AddViewItem("ribao", 9, 13, "累计.5高炉.焦丁比");
            AddViewItem("ribao", 9, 14, "累计.5高炉.冶炼强度");
            AddViewItem("ribao", 9, 15, "累计.5高炉.合格率");
            AddViewItem("ribao", 9, 16, "累计.5高炉.休风率");
            AddViewItem("ribao", 9, 17, "累计.5高炉.平均风温");
            AddViewItem("ribao", 9, 18, "累计.5高炉.电单耗");
            AddViewItem("ribao", 9, 19, "累计.5高炉.富氧率");
            AddViewItem("ribao", 9, 20, "累计.5高炉.生铁含硅");
            AddViewItem("ribao", 9, 21, "累计.5高炉.一级品率");
            AddViewItem("ribao", 9, 22, "累计.5高炉.氧比");
            AddViewItem("ribao", 9, 23, "累计.5高炉.氮比");
            AddViewItem("ribao", 9, 24, "累计.5高炉.焦炭负荷");
            AddViewItem("ribao", 9, 25, "累计.5高炉.燃料比");
            AddViewItem("ribao", 9, 26, "累计.5高炉.废铁单耗");
            AddViewItem("ribao", 9, 30, "累计.5高炉.TRT发电");
            AddViewItem("ribao", 9, 31, "累计.5高炉.返矿");
            AddViewItem("ribao", 9, 32, "累计.5高炉.矿槽");
            AddViewItem("ribao", 9, 33, "累计.5高炉.休风工艺");
            AddViewItem("ribao", 9, 34, "累计.5高炉.休风设备");
            AddViewItem("ribao", 9, 35, "累计.5高炉.休风电");
            AddViewItem("ribao", 9, 36, "累计.5高炉.休风外围");
            AddViewItem("ribao", 9, 37, "累计.5高炉.休风计划");
            AddViewItem("ribao", 9, 38, "累计.5高炉.休风");


            AddViewItem("ribao", 10, 3, "本日.6高炉.全铁产量");
            AddViewItem("ribao", 10, 4, "本日.6高炉.合格铁");
            AddViewItem("ribao", 10, 5, "本日.6高炉.号外铁");
            AddViewItem("ribao", 10, 6, "本日.6高炉.铸造铁");
            AddViewItem("ribao", 10, 7, "本日.6高炉.利用系数");
            AddViewItem("ribao", 10, 8, "本日.6高炉.原料矿单耗");
            AddViewItem("ribao", 10, 9, "本日.6高炉.干焦丁");
            AddViewItem("ribao", 10, 10, "本日.6高炉.综合焦炭单耗");
            AddViewItem("ribao", 10, 11, "本日.6高炉.入炉焦炭单耗");
            AddViewItem("ribao", 10, 12, "本日.6高炉.煤粉单耗");
            AddViewItem("ribao", 10, 13, "本日.6高炉.焦丁比");
            AddViewItem("ribao", 10, 14, "本日.6高炉.冶炼强度");
            AddViewItem("ribao", 10, 15, "本日.6高炉.合格率");
            AddViewItem("ribao", 10, 16, "本日.6高炉.休风率");
            AddViewItem("ribao", 10, 17, "本日.6高炉.平均风温");
            AddViewItem("ribao", 10, 18, "本日.6高炉.耗电");
            AddViewItem("ribao", 10, 19, "本日.6高炉.富氧率");
            AddViewItem("ribao", 10, 20, "本日.6高炉.生铁含硅");
            AddViewItem("ribao", 10, 21, "本日.6高炉.一级品率");
            AddViewItem("ribao", 10, 22, "本日.6高炉.氧比");
            AddViewItem("ribao", 10, 23, "本日.6高炉.氮比");
            AddViewItem("ribao", 10, 24, "本日.6高炉.焦炭负荷");
            AddViewItem("ribao", 10, 25, "本日.6高炉.燃料比");
            AddViewItem("ribao", 10, 26, "本日.6高炉.废铁单耗");
            AddViewItem("ribao", 10, 30, "本日.6高炉.TRT发电");
            AddViewItem("ribao", 10, 31, "本日.6高炉.返矿");
            AddViewItem("ribao", 10, 32, "本日.6高炉.矿槽");
            AddViewItem("ribao", 10, 33, "本日.6高炉.休风工艺");
            AddViewItem("ribao", 10, 34, "本日.6高炉.休风设备");
            AddViewItem("ribao", 10, 35, "本日.6高炉.休风电");
            AddViewItem("ribao", 10, 36, "本日.6高炉.休风外围");
            AddViewItem("ribao", 10, 37, "本日.6高炉.休风计划");
            AddViewItem("ribao", 10, 38, "本日.6高炉.休风");

            AddViewItem("ribao", 11, 3, "累计.6高炉.全铁产量");
            AddViewItem("ribao", 11, 4, "累计.6高炉.合格铁");
            AddViewItem("ribao", 11, 5, "累计.6高炉.号外铁");
            AddViewItem("ribao", 11, 6, "累计.6高炉.铸造铁");
            AddViewItem("ribao", 11, 7, "累计.6高炉.利用系数");
            AddViewItem("ribao", 11, 8, "累计.6高炉.原料矿单耗");
            AddViewItem("ribao", 11, 9, "累计.6高炉.干焦丁");
            AddViewItem("ribao", 11, 10, "累计.6高炉.综合焦炭单耗");
            AddViewItem("ribao", 11, 11, "累计.6高炉.入炉焦炭单耗");
            AddViewItem("ribao", 11, 12, "累计.6高炉.煤粉单耗");
            AddViewItem("ribao", 11, 13, "累计.6高炉.焦丁比");
            AddViewItem("ribao", 11, 14, "累计.6高炉.冶炼强度");
            AddViewItem("ribao", 11, 15, "累计.6高炉.合格率");
            AddViewItem("ribao", 11, 16, "累计.6高炉.休风率");
            AddViewItem("ribao", 11, 17, "累计.6高炉.平均风温");
            AddViewItem("ribao", 11, 18, "累计.6高炉.电单耗");
            AddViewItem("ribao", 11, 19, "累计.6高炉.富氧率");
            AddViewItem("ribao", 11, 20, "累计.6高炉.生铁含硅");
            AddViewItem("ribao", 11, 21, "累计.6高炉.一级品率");
            AddViewItem("ribao", 11, 22, "累计.6高炉.氧比");
            AddViewItem("ribao", 11, 23, "累计.6高炉.氮比");
            AddViewItem("ribao", 11, 24, "累计.6高炉.焦炭负荷");
            AddViewItem("ribao", 11, 25, "累计.6高炉.燃料比");
            AddViewItem("ribao", 11, 26, "累计.6高炉.废铁单耗");
            AddViewItem("ribao", 11, 30, "累计.6高炉.TRT发电");
            AddViewItem("ribao", 11, 31, "累计.6高炉.返矿");
            AddViewItem("ribao", 11, 32, "累计.6高炉.矿槽");
            AddViewItem("ribao", 11, 33, "累计.6高炉.休风工艺");
            AddViewItem("ribao", 11, 34, "累计.6高炉.休风设备");
            AddViewItem("ribao", 11, 35, "累计.6高炉.休风电");
            AddViewItem("ribao", 11, 36, "累计.6高炉.休风外围");
            AddViewItem("ribao", 11, 37, "累计.6高炉.休风计划");
            AddViewItem("ribao", 11, 38, "累计.6高炉.休风");

            AddViewItem("ribao", 14, 3, "本日.7高炉.全铁产量");
            AddViewItem("ribao", 14, 4, "本日.7高炉.合格铁");
            AddViewItem("ribao", 14, 5, "本日.7高炉.号外铁");
            AddViewItem("ribao", 14, 6, "本日.7高炉.铸造铁");
            AddViewItem("ribao", 14, 7, "本日.7高炉.利用系数");
            AddViewItem("ribao", 14, 8, "本日.7高炉.原料矿单耗");
            AddViewItem("ribao", 14, 9, "本日.7高炉.干焦丁");
            AddViewItem("ribao", 14, 10, "本日.7高炉.综合焦炭单耗");
            AddViewItem("ribao", 14, 11, "本日.7高炉.入炉焦炭单耗");
            AddViewItem("ribao", 14, 12, "本日.7高炉.煤粉单耗");
            AddViewItem("ribao", 14, 13, "本日.7高炉.焦丁比");
            AddViewItem("ribao", 14, 14, "本日.7高炉.冶炼强度");
            AddViewItem("ribao", 14, 15, "本日.7高炉.合格率");
            AddViewItem("ribao", 14, 16, "本日.7高炉.休风率");
            AddViewItem("ribao", 14, 17, "本日.7高炉.平均风温");
            AddViewItem("ribao", 14, 18, "本日.7高炉.耗电");
            AddViewItem("ribao", 14, 19, "本日.7高炉.富氧率");
            AddViewItem("ribao", 14, 20, "本日.7高炉.生铁含硅");
            AddViewItem("ribao", 14, 21, "本日.7高炉.一级品率");
            AddViewItem("ribao", 14, 22, "本日.7高炉.氧比");
            AddViewItem("ribao", 14, 23, "本日.7高炉.氮比");
            AddViewItem("ribao", 14, 24, "本日.7高炉.焦炭负荷");
            AddViewItem("ribao", 14, 25, "本日.7高炉.燃料比");
            AddViewItem("ribao", 14, 26, "本日.7高炉.废铁单耗");
            AddViewItem("ribao", 14, 30, "本日.7高炉.TRT发电");
            AddViewItem("ribao", 14, 31, "本日.7高炉.返矿");
            AddViewItem("ribao", 14, 32, "本日.7高炉.矿槽");
            AddViewItem("ribao", 14, 33, "本日.7高炉.休风工艺");
            AddViewItem("ribao", 14, 34, "本日.7高炉.休风设备");
            AddViewItem("ribao", 14, 35, "本日.7高炉.休风电");
            AddViewItem("ribao", 14, 36, "本日.7高炉.休风外围");
            AddViewItem("ribao", 14, 37, "本日.7高炉.休风计划");
            AddViewItem("ribao", 14, 38, "本日.7高炉.休风");

            AddViewItem("ribao", 15, 3, "累计.7高炉.全铁产量");
            AddViewItem("ribao", 15, 4, "累计.7高炉.合格铁");
            AddViewItem("ribao", 15, 5, "累计.7高炉.号外铁");
            AddViewItem("ribao", 15, 6, "累计.7高炉.铸造铁");
            AddViewItem("ribao", 15, 7, "累计.7高炉.利用系数");
            AddViewItem("ribao", 15, 8, "累计.7高炉.原料矿单耗");
            AddViewItem("ribao", 15, 9, "累计.7高炉.干焦丁");
            AddViewItem("ribao", 15, 10, "累计.7高炉.综合焦炭单耗");
            AddViewItem("ribao", 15, 11, "累计.7高炉.入炉焦炭单耗");
            AddViewItem("ribao", 15, 12, "累计.7高炉.煤粉单耗");
            AddViewItem("ribao", 15, 13, "累计.7高炉.焦丁比");
            AddViewItem("ribao", 15, 14, "累计.7高炉.冶炼强度");
            AddViewItem("ribao", 15, 15, "累计.7高炉.合格率");
            AddViewItem("ribao", 15, 16, "累计.7高炉.休风率");
            AddViewItem("ribao", 15, 17, "累计.7高炉.平均风温");
            AddViewItem("ribao", 15, 18, "累计.7高炉.电单耗");
            AddViewItem("ribao", 15, 19, "累计.7高炉.富氧率");
            AddViewItem("ribao", 15, 20, "累计.7高炉.生铁含硅");
            AddViewItem("ribao", 15, 21, "累计.7高炉.一级品率");
            AddViewItem("ribao", 15, 22, "累计.7高炉.氧比");
            AddViewItem("ribao", 15, 23, "累计.7高炉.氮比");
            AddViewItem("ribao", 15, 24, "累计.7高炉.焦炭负荷");
            AddViewItem("ribao", 15, 25, "累计.7高炉.燃料比");
            AddViewItem("ribao", 15, 26, "累计.7高炉.废铁单耗");
            AddViewItem("ribao", 15, 30, "累计.7高炉.TRT发电");
            AddViewItem("ribao", 15, 31, "累计.7高炉.返矿");
            AddViewItem("ribao", 15, 32, "累计.7高炉.矿槽");
            AddViewItem("ribao", 15, 33, "累计.7高炉.休风工艺");
            AddViewItem("ribao", 15, 34, "累计.7高炉.休风设备");
            AddViewItem("ribao", 15, 35, "累计.7高炉.休风电");
            AddViewItem("ribao", 15, 36, "累计.7高炉.休风外围");
            AddViewItem("ribao", 15, 37, "累计.7高炉.休风计划");
            AddViewItem("ribao", 15, 38, "累计.7高炉.休风");

            AddViewItem("ribao", 16, 3, "本日.8高炉.全铁产量");
            AddViewItem("ribao", 16, 4, "本日.8高炉.合格铁");
            AddViewItem("ribao", 16, 5, "本日.8高炉.号外铁");
            AddViewItem("ribao", 16, 6, "本日.8高炉.铸造铁");
            AddViewItem("ribao", 16, 7, "本日.8高炉.利用系数");
            AddViewItem("ribao", 16, 8, "本日.8高炉.原料矿单耗");
            AddViewItem("ribao", 16, 9, "本日.8高炉.干焦丁");
            AddViewItem("ribao", 16, 10, "本日.8高炉.综合焦炭单耗");
            AddViewItem("ribao", 16, 11, "本日.8高炉.入炉焦炭单耗");
            AddViewItem("ribao", 16, 12, "本日.8高炉.煤粉单耗");
            AddViewItem("ribao", 16, 13, "本日.8高炉.焦丁比");
            AddViewItem("ribao", 16, 14, "本日.8高炉.冶炼强度");
            AddViewItem("ribao", 16, 15, "本日.8高炉.合格率");
            AddViewItem("ribao", 16, 16, "本日.8高炉.休风率");
            AddViewItem("ribao", 16, 17, "本日.8高炉.平均风温");
            AddViewItem("ribao", 16, 18, "本日.8高炉.耗电");
            AddViewItem("ribao", 16, 19, "本日.8高炉.富氧率");
            AddViewItem("ribao", 16, 20, "本日.8高炉.生铁含硅");
            AddViewItem("ribao", 16, 21, "本日.8高炉.一级品率");
            AddViewItem("ribao", 16, 22, "本日.8高炉.氧比");
            AddViewItem("ribao", 16, 23, "本日.8高炉.氮比");
            AddViewItem("ribao", 16, 24, "本日.8高炉.焦炭负荷");
            AddViewItem("ribao", 16, 25, "本日.8高炉.燃料比");
            AddViewItem("ribao", 16, 26, "本日.8高炉.废铁单耗");
            AddViewItem("ribao", 16, 30, "本日.8高炉.TRT发电");
            AddViewItem("ribao", 16, 31, "本日.8高炉.返矿");
            AddViewItem("ribao", 16, 32, "本日.8高炉.矿槽");
            AddViewItem("ribao", 16, 33, "本日.8高炉.休风工艺");
            AddViewItem("ribao", 16, 34, "本日.8高炉.休风设备");
            AddViewItem("ribao", 16, 35, "本日.8高炉.休风电");
            AddViewItem("ribao", 16, 36, "本日.8高炉.休风外围");
            AddViewItem("ribao", 16, 37, "本日.8高炉.休风计划");
            AddViewItem("ribao", 16, 38, "本日.8高炉.休风");

            AddViewItem("ribao", 17, 3, "累计.8高炉.全铁产量");
            AddViewItem("ribao", 17, 4, "累计.8高炉.合格铁");
            AddViewItem("ribao", 17, 5, "累计.8高炉.号外铁");
            AddViewItem("ribao", 17, 6, "累计.8高炉.铸造铁");
            AddViewItem("ribao", 17, 7, "累计.8高炉.利用系数");
            AddViewItem("ribao", 17, 8, "累计.8高炉.原料矿单耗");
            AddViewItem("ribao", 17, 9, "累计.8高炉.干焦丁");
            AddViewItem("ribao", 17, 10, "累计.8高炉.综合焦炭单耗");
            AddViewItem("ribao", 17, 11, "累计.8高炉.入炉焦炭单耗");
            AddViewItem("ribao", 17, 12, "累计.8高炉.煤粉单耗");
            AddViewItem("ribao", 17, 13, "累计.8高炉.焦丁比");
            AddViewItem("ribao", 17, 14, "累计.8高炉.冶炼强度");
            AddViewItem("ribao", 17, 15, "累计.8高炉.合格率");
            AddViewItem("ribao", 17, 16, "累计.8高炉.休风率");
            AddViewItem("ribao", 17, 17, "累计.8高炉.平均风温");
            AddViewItem("ribao", 17, 18, "累计.8高炉.电单耗");
            AddViewItem("ribao", 17, 19, "累计.8高炉.富氧率");
            AddViewItem("ribao", 17, 20, "累计.8高炉.生铁含硅");
            AddViewItem("ribao", 17, 21, "累计.8高炉.一级品率");
            AddViewItem("ribao", 17, 22, "累计.8高炉.氧比");
            AddViewItem("ribao", 17, 23, "累计.8高炉.氮比");
            AddViewItem("ribao", 17, 24, "累计.8高炉.焦炭负荷");
            AddViewItem("ribao", 17, 25, "累计.8高炉.燃料比");
            AddViewItem("ribao", 17, 26, "累计.8高炉.废铁单耗");
            AddViewItem("ribao", 17, 30, "累计.8高炉.TRT发电");
            AddViewItem("ribao", 17, 31, "累计.8高炉.返矿");
            AddViewItem("ribao", 17, 32, "累计.8高炉.矿槽");
            AddViewItem("ribao", 17, 33, "累计.8高炉.休风工艺");
            AddViewItem("ribao", 17, 34, "累计.8高炉.休风设备");
            AddViewItem("ribao", 17, 35, "累计.8高炉.休风电");
            AddViewItem("ribao", 17, 36, "累计.8高炉.休风外围");
            AddViewItem("ribao", 17, 37, "累计.8高炉.休风计划");
            AddViewItem("ribao", 17, 38, "累计.8高炉.休风");

            AddViewItem("ribao", 22, 3, "本日.1高炉.全铁产量");
            AddViewItem("ribao", 22, 4, "本日.1高炉.合格铁");
            AddViewItem("ribao", 22, 5, "本日.1高炉.号外铁");
            AddViewItem("ribao", 22, 6, "本日.1高炉.铸造铁");
            AddViewItem("ribao", 22, 7, "本日.1高炉.利用系数");
            AddViewItem("ribao", 22, 8, "本日.1高炉.原料矿单耗");
            AddViewItem("ribao", 22, 9, "本日.1高炉.干焦丁");
            AddViewItem("ribao", 22, 10, "本日.1高炉.综合焦炭单耗");
            AddViewItem("ribao", 22, 11, "本日.1高炉.入炉焦炭单耗");
            AddViewItem("ribao", 22, 12, "本日.1高炉.煤粉单耗");
            AddViewItem("ribao", 22, 13, "本日.1高炉.焦丁比");
            AddViewItem("ribao", 22, 14, "本日.1高炉.冶炼强度");
            AddViewItem("ribao", 22, 15, "本日.1高炉.合格率");
            AddViewItem("ribao", 22, 16, "本日.1高炉.休风率");
            AddViewItem("ribao", 22, 17, "本日.1高炉.平均风温");
            AddViewItem("ribao", 22, 18, "本日.1高炉.耗电");
            AddViewItem("ribao", 22, 19, "本日.1高炉.富氧率");
            AddViewItem("ribao", 22, 20, "本日.1高炉.生铁含硅");
            AddViewItem("ribao", 22, 21, "本日.1高炉.一级品率");
            AddViewItem("ribao", 22, 22, "本日.1高炉.氧比");
            AddViewItem("ribao", 22, 23, "本日.1高炉.氮比");
            AddViewItem("ribao", 22, 24, "本日.1高炉.焦炭负荷");
            AddViewItem("ribao", 22, 25, "本日.1高炉.燃料比");
            AddViewItem("ribao", 22, 26, "本日.1高炉.废铁单耗");
            AddViewItem("ribao", 22, 30, "本日.1高炉.TRT发电");
            AddViewItem("ribao", 22, 31, "本日.1高炉.返矿");
            AddViewItem("ribao", 22, 32, "本日.1高炉.矿槽");
            AddViewItem("ribao", 22, 33, "本日.1高炉.休风工艺");
            AddViewItem("ribao", 22, 34, "本日.1高炉.休风设备");
            AddViewItem("ribao", 22, 35, "本日.1高炉.休风电");
            AddViewItem("ribao", 22, 36, "本日.1高炉.休风外围");
            AddViewItem("ribao", 22, 37, "本日.1高炉.休风计划");
            AddViewItem("ribao", 22, 38, "本日.1高炉.休风");

            AddViewItem("ribao", 23, 3, "累计.1高炉.全铁产量");
            AddViewItem("ribao", 23, 4, "累计.1高炉.合格铁");
            AddViewItem("ribao", 23, 5, "累计.1高炉.号外铁");
            AddViewItem("ribao", 23, 6, "累计.1高炉.铸造铁");
            AddViewItem("ribao", 23, 7, "累计.1高炉.利用系数");
            AddViewItem("ribao", 23, 8, "累计.1高炉.原料矿单耗");
            AddViewItem("ribao", 23, 9, "累计.1高炉.干焦丁");
            AddViewItem("ribao", 23, 10, "累计.1高炉.综合焦炭单耗");
            AddViewItem("ribao", 23, 11, "累计.1高炉.入炉焦炭单耗");
            AddViewItem("ribao", 23, 12, "累计.1高炉.煤粉单耗");
            AddViewItem("ribao", 23, 13, "累计.1高炉.焦丁比");
            AddViewItem("ribao", 23, 14, "累计.1高炉.冶炼强度");
            AddViewItem("ribao", 23, 15, "累计.1高炉.合格率");
            AddViewItem("ribao", 23, 16, "累计.1高炉.休风率");
            AddViewItem("ribao", 23, 17, "累计.1高炉.平均风温");
            AddViewItem("ribao", 23, 18, "累计.1高炉.电单耗");
            AddViewItem("ribao", 23, 19, "累计.1高炉.富氧率");
            AddViewItem("ribao", 23, 20, "累计.1高炉.生铁含硅");
            AddViewItem("ribao", 23, 21, "累计.1高炉.一级品率");
            AddViewItem("ribao", 23, 22, "累计.1高炉.氧比");
            AddViewItem("ribao", 23, 23, "累计.1高炉.氮比");
            AddViewItem("ribao", 23, 24, "累计.1高炉.焦炭负荷");
            AddViewItem("ribao", 23, 25, "累计.1高炉.燃料比");
            AddViewItem("ribao", 23, 26, "累计.1高炉.废铁单耗");
            AddViewItem("ribao", 23, 30, "累计.1高炉.TRT发电");
            AddViewItem("ribao", 23, 31, "累计.1高炉.返矿");
            AddViewItem("ribao", 23, 32, "累计.1高炉.矿槽");
            AddViewItem("ribao", 23, 33, "累计.1高炉.休风工艺");
            AddViewItem("ribao", 23, 34, "累计.1高炉.休风设备");
            AddViewItem("ribao", 23, 35, "累计.1高炉.休风电");
            AddViewItem("ribao", 23, 36, "累计.1高炉.休风外围");
            AddViewItem("ribao", 23, 37, "累计.1高炉.休风计划");
            AddViewItem("ribao", 23, 38, "累计.1高炉.休风");

            AddViewItem("ribao", 12, 3, "本日.456高炉合计.全铁产量");
            AddViewItem("ribao", 12, 4, "本日.456高炉合计.合格铁");
            AddViewItem("ribao", 12, 5, "本日.456高炉合计.号外铁");
            AddViewItem("ribao", 12, 6, "本日.456高炉合计.铸造铁");
            AddViewItem("ribao", 12, 7, "本日.456高炉合计.利用系数");
            AddViewItem("ribao", 12, 8, "本日.456高炉合计.原料矿单耗");
            AddViewItem("ribao", 12, 9, "本日.456高炉合计.干焦丁");
            AddViewItem("ribao", 12, 10, "本日.456高炉合计.综合焦炭单耗");
            AddViewItem("ribao", 12, 11, "本日.456高炉合计.入炉焦炭单耗");
            AddViewItem("ribao", 12, 12, "本日.456高炉合计.煤粉单耗");
            AddViewItem("ribao", 12, 13, "本日.456高炉合计.焦丁比");
            AddViewItem("ribao", 12, 14, "本日.456高炉合计.冶炼强度");
            AddViewItem("ribao", 12, 15, "本日.456高炉合计.合格率");
            AddViewItem("ribao", 12, 16, "本日.456高炉合计.休风率");
            AddViewItem("ribao", 12, 17, "本日.456高炉合计.平均风温");
            AddViewItem("ribao", 12, 18, "本日.456高炉合计.耗电");
            AddViewItem("ribao", 12, 19, "本日.456高炉合计.富氧率");
            AddViewItem("ribao", 12, 20, "本日.456高炉合计.生铁含硅");
            AddViewItem("ribao", 12, 21, "本日.456高炉合计.一级品率");
            AddViewItem("ribao", 12, 22, "本日.456高炉合计.氧比");
            AddViewItem("ribao", 12, 23, "本日.456高炉合计.氮比");
            AddViewItem("ribao", 12, 24, "本日.456高炉合计.焦炭负荷");
            AddViewItem("ribao", 12, 25, "本日.456高炉合计.燃料比");
            AddViewItem("ribao", 12, 26, "本日.456高炉合计.废铁单耗");
            AddViewItem("ribao", 12, 30, "本日.456高炉合计.TRT发电");
            AddViewItem("ribao", 12, 31, "本日.456高炉合计.返矿");
            AddViewItem("ribao", 12, 32, "本日.456高炉合计.矿槽");
            AddViewItem("ribao", 12, 33, "本日.456高炉合计.休风工艺");
            AddViewItem("ribao", 12, 34, "本日.456高炉合计.休风设备");
            AddViewItem("ribao", 12, 35, "本日.456高炉合计.休风电");
            AddViewItem("ribao", 12, 36, "本日.456高炉合计.休风外围");
            AddViewItem("ribao", 12, 37, "本日.456高炉合计.休风计划");
            AddViewItem("ribao", 12, 38, "本日.456高炉合计.休风");

            AddViewItem("ribao", 13, 3, "累计.456高炉合计.全铁产量");
            AddViewItem("ribao", 13, 4, "累计.456高炉合计.合格铁");
            AddViewItem("ribao", 13, 5, "累计.456高炉合计.号外铁");
            AddViewItem("ribao", 13, 6, "累计.456高炉合计.铸造铁");
            AddViewItem("ribao", 13, 7, "累计.456高炉合计.利用系数");
            AddViewItem("ribao", 13, 8, "累计.456高炉合计.原料矿单耗");
            AddViewItem("ribao", 13, 9, "累计.456高炉合计.干焦丁");
            AddViewItem("ribao", 13, 10, "累计.456高炉合计.综合焦炭单耗");
            AddViewItem("ribao", 13, 11, "累计.456高炉合计.入炉焦炭单耗");
            AddViewItem("ribao", 13, 12, "累计.456高炉合计.煤粉单耗");
            AddViewItem("ribao", 13, 13, "累计.456高炉合计.焦丁比");
            AddViewItem("ribao", 13, 14, "累计.456高炉合计.冶炼强度");
            AddViewItem("ribao", 13, 15, "累计.456高炉合计.合格率");
            AddViewItem("ribao", 13, 16, "累计.456高炉合计.休风率");
            AddViewItem("ribao", 13, 17, "累计.456高炉合计.平均风温");
            AddViewItem("ribao", 13, 18, "累计.456高炉合计.电单耗");
            AddViewItem("ribao", 13, 19, "累计.456高炉合计.富氧率");
            AddViewItem("ribao", 13, 20, "累计.456高炉合计.生铁含硅");
            AddViewItem("ribao", 13, 21, "累计.456高炉合计.一级品率");
            AddViewItem("ribao", 13, 22, "累计.456高炉合计.氧比");
            AddViewItem("ribao", 13, 23, "累计.456高炉合计.氮比");
            AddViewItem("ribao", 13, 24, "累计.456高炉合计.焦炭负荷");
            AddViewItem("ribao", 13, 25, "累计.456高炉合计.燃料比");
            AddViewItem("ribao", 13, 26, "累计.456高炉合计.废铁单耗");
            AddViewItem("ribao", 13, 30, "累计.456高炉合计.TRT发电");
            AddViewItem("ribao", 13, 31, "累计.456高炉合计.返矿");
            AddViewItem("ribao", 13, 32, "累计.456高炉合计.矿槽");
            AddViewItem("ribao", 13, 33, "累计.456高炉合计.休风工艺");
            AddViewItem("ribao", 13, 34, "累计.456高炉合计.休风设备");
            AddViewItem("ribao", 13, 35, "累计.456高炉合计.休风电");
            AddViewItem("ribao", 13, 36, "累计.456高炉合计.休风外围");
            AddViewItem("ribao", 13, 37, "累计.456高炉合计.休风计划");
            AddViewItem("ribao", 13, 38, "累计.456高炉合计.休风");

            AddViewItem("ribao", 18, 3, "本日.78高炉合计.全铁产量");
            AddViewItem("ribao", 18, 4, "本日.78高炉合计.合格铁");
            AddViewItem("ribao", 18, 5, "本日.78高炉合计.号外铁");
            AddViewItem("ribao", 18, 6, "本日.78高炉合计.铸造铁");
            AddViewItem("ribao", 18, 7, "本日.78高炉合计.利用系数");
            AddViewItem("ribao", 18, 8, "本日.78高炉合计.原料矿单耗");
            AddViewItem("ribao", 18, 9, "本日.78高炉合计.干焦丁");
            AddViewItem("ribao", 18, 10, "本日.78高炉合计.综合焦炭单耗");
            AddViewItem("ribao", 18, 11, "本日.78高炉合计.入炉焦炭单耗");
            AddViewItem("ribao", 18, 12, "本日.78高炉合计.煤粉单耗");
            AddViewItem("ribao", 18, 13, "本日.78高炉合计.焦丁比");
            AddViewItem("ribao", 18, 14, "本日.78高炉合计.冶炼强度");
            AddViewItem("ribao", 18, 15, "本日.78高炉合计.合格率");
            AddViewItem("ribao", 18, 16, "本日.78高炉合计.休风率");
            AddViewItem("ribao", 18, 17, "本日.78高炉合计.平均风温");
            AddViewItem("ribao", 18, 18, "本日.78高炉合计.耗电");
            AddViewItem("ribao", 18, 19, "本日.78高炉合计.富氧率");
            AddViewItem("ribao", 18, 20, "本日.78高炉合计.生铁含硅");
            AddViewItem("ribao", 18, 21, "本日.78高炉合计.一级品率");
            AddViewItem("ribao", 18, 22, "本日.78高炉合计.氧比");
            AddViewItem("ribao", 18, 23, "本日.78高炉合计.氮比");
            AddViewItem("ribao", 18, 24, "本日.78高炉合计.焦炭负荷");
            AddViewItem("ribao", 18, 25, "本日.78高炉合计.燃料比");
            AddViewItem("ribao", 18, 26, "本日.78高炉合计.废铁单耗");
            AddViewItem("ribao", 18, 30, "本日.78高炉合计.TRT发电");
            AddViewItem("ribao", 18, 31, "本日.78高炉合计.返矿");
            AddViewItem("ribao", 18, 32, "本日.78高炉合计.矿槽");
            AddViewItem("ribao", 18, 33, "本日.78高炉合计.休风工艺");
            AddViewItem("ribao", 18, 34, "本日.78高炉合计.休风设备");
            AddViewItem("ribao", 18, 35, "本日.78高炉合计.休风电");
            AddViewItem("ribao", 18, 36, "本日.78高炉合计.休风外围");
            AddViewItem("ribao", 18, 37, "本日.78高炉合计.休风计划");
            AddViewItem("ribao", 18, 38, "本日.78高炉合计.休风");

            AddViewItem("ribao", 19, 3, "累计.78高炉合计.全铁产量");
            AddViewItem("ribao", 19, 4, "累计.78高炉合计.合格铁");
            AddViewItem("ribao", 19, 5, "累计.78高炉合计.号外铁");
            AddViewItem("ribao", 19, 6, "累计.78高炉合计.铸造铁");
            AddViewItem("ribao", 19, 7, "累计.78高炉合计.利用系数");
            AddViewItem("ribao", 19, 8, "累计.78高炉合计.原料矿单耗");
            AddViewItem("ribao", 19, 9, "累计.78高炉合计.干焦丁");
            AddViewItem("ribao", 19, 10, "累计.78高炉合计.综合焦炭单耗");
            AddViewItem("ribao", 19, 11, "累计.78高炉合计.入炉焦炭单耗");
            AddViewItem("ribao", 19, 12, "累计.78高炉合计.煤粉单耗");
            AddViewItem("ribao", 19, 13, "累计.78高炉合计.焦丁比");
            AddViewItem("ribao", 19, 14, "累计.78高炉合计.冶炼强度");
            AddViewItem("ribao", 19, 15, "累计.78高炉合计.合格率");
            AddViewItem("ribao", 19, 16, "累计.78高炉合计.休风率");
            AddViewItem("ribao", 19, 17, "累计.78高炉合计.平均风温");
            AddViewItem("ribao", 19, 18, "累计.78高炉合计.电单耗");
            AddViewItem("ribao", 19, 19, "累计.78高炉合计.富氧率");
            AddViewItem("ribao", 19, 20, "累计.78高炉合计.生铁含硅");
            AddViewItem("ribao", 19, 21, "累计.78高炉合计.一级品率");
            AddViewItem("ribao", 19, 22, "累计.78高炉合计.氧比");
            AddViewItem("ribao", 19, 23, "累计.78高炉合计.氮比");
            AddViewItem("ribao", 19, 24, "累计.78高炉合计.焦炭负荷");
            AddViewItem("ribao", 19, 25, "累计.78高炉合计.燃料比");
            AddViewItem("ribao", 19, 26, "累计.78高炉合计.废铁单耗");
            AddViewItem("ribao", 19, 30, "累计.78高炉合计.TRT发电");
            AddViewItem("ribao", 19, 31, "累计.78高炉合计.返矿");
            AddViewItem("ribao", 19, 32, "累计.78高炉合计.矿槽");
            AddViewItem("ribao", 19, 33, "累计.78高炉合计.休风工艺");
            AddViewItem("ribao", 19, 34, "累计.78高炉合计.休风设备");
            AddViewItem("ribao", 19, 35, "累计.78高炉合计.休风电");
            AddViewItem("ribao", 19, 36, "累计.78高炉合计.休风外围");
            AddViewItem("ribao", 19, 37, "累计.78高炉合计.休风计划");
            AddViewItem("ribao", 19, 38, "累计.78高炉合计.休风");

            AddViewItem("ribao", 20, 3, "本日.4_8高炉合计.全铁产量");
            AddViewItem("ribao", 20, 4, "本日.4_8高炉合计.合格铁");
            AddViewItem("ribao", 20, 5, "本日.4_8高炉合计.号外铁");
            AddViewItem("ribao", 20, 6, "本日.4_8高炉合计.铸造铁");
            AddViewItem("ribao", 20, 7, "本日.4_8高炉合计.利用系数");
            AddViewItem("ribao", 20, 8, "本日.4_8高炉合计.原料矿单耗");
            AddViewItem("ribao", 20, 9, "本日.4_8高炉合计.干焦丁");
            AddViewItem("ribao", 20, 10, "本日.4_8高炉合计.综合焦炭单耗");
            AddViewItem("ribao", 20, 11, "本日.4_8高炉合计.入炉焦炭单耗");
            AddViewItem("ribao", 20, 12, "本日.4_8高炉合计.煤粉单耗");
            AddViewItem("ribao", 20, 13, "本日.4_8高炉合计.焦丁比");
            AddViewItem("ribao", 20, 14, "本日.4_8高炉合计.冶炼强度");
            AddViewItem("ribao", 20, 15, "本日.4_8高炉合计.合格率");
            AddViewItem("ribao", 20, 16, "本日.4_8高炉合计.休风率");
            AddViewItem("ribao", 20, 17, "本日.4_8高炉合计.平均风温");
            AddViewItem("ribao", 20, 18, "本日.4_8高炉合计.耗电");
            AddViewItem("ribao", 20, 19, "本日.4_8高炉合计.富氧率");
            AddViewItem("ribao", 20, 20, "本日.4_8高炉合计.生铁含硅");
            AddViewItem("ribao", 20, 21, "本日.4_8高炉合计.一级品率");
            AddViewItem("ribao", 20, 22, "本日.4_8高炉合计.氧比");
            AddViewItem("ribao", 20, 23, "本日.4_8高炉合计.氮比");
            AddViewItem("ribao", 20, 24, "本日.4_8高炉合计.焦炭负荷");
            AddViewItem("ribao", 20, 25, "本日.4_8高炉合计.燃料比");
            AddViewItem("ribao", 20, 26, "本日.4_8高炉合计.废铁单耗");
            AddViewItem("ribao", 20, 30, "本日.4_8高炉合计.TRT发电");
            AddViewItem("ribao", 20, 31, "本日.4_8高炉合计.返矿");
            AddViewItem("ribao", 20, 32, "本日.4_8高炉合计.矿槽");
            AddViewItem("ribao", 20, 33, "本日.4_8高炉合计.休风工艺");
            AddViewItem("ribao", 20, 34, "本日.4_8高炉合计.休风设备");
            AddViewItem("ribao", 20, 35, "本日.4_8高炉合计.休风电");
            AddViewItem("ribao", 20, 36, "本日.4_8高炉合计.休风外围");
            AddViewItem("ribao", 20, 37, "本日.4_8高炉合计.休风计划");
            AddViewItem("ribao", 20, 38, "本日.4_8高炉合计.休风");

            AddViewItem("ribao", 21, 3, "累计.4_8高炉合计.全铁产量");
            AddViewItem("ribao", 21, 4, "累计.4_8高炉合计.合格铁");
            AddViewItem("ribao", 21, 5, "累计.4_8高炉合计.号外铁");
            AddViewItem("ribao", 21, 6, "累计.4_8高炉合计.铸造铁");
            AddViewItem("ribao", 21, 7, "累计.4_8高炉合计.利用系数");
            AddViewItem("ribao", 21, 8, "累计.4_8高炉合计.原料矿单耗");
            AddViewItem("ribao", 21, 9, "累计.4_8高炉合计.干焦丁");
            AddViewItem("ribao", 21, 10, "累计.4_8高炉合计.综合焦炭单耗");
            AddViewItem("ribao", 21, 11, "累计.4_8高炉合计.入炉焦炭单耗");
            AddViewItem("ribao", 21, 12, "累计.4_8高炉合计.煤粉单耗");
            AddViewItem("ribao", 21, 13, "累计.4_8高炉合计.焦丁比");
            AddViewItem("ribao", 21, 14, "累计.4_8高炉合计.冶炼强度");
            AddViewItem("ribao", 21, 15, "累计.4_8高炉合计.合格率");
            AddViewItem("ribao", 21, 16, "累计.4_8高炉合计.休风率");
            AddViewItem("ribao", 21, 17, "累计.4_8高炉合计.平均风温");
            AddViewItem("ribao", 21, 18, "累计.4_8高炉合计.电单耗");
            AddViewItem("ribao", 21, 19, "累计.4_8高炉合计.富氧率");
            AddViewItem("ribao", 21, 20, "累计.4_8高炉合计.生铁含硅");
            AddViewItem("ribao", 21, 21, "累计.4_8高炉合计.一级品率");
            AddViewItem("ribao", 21, 22, "累计.4_8高炉合计.氧比");
            AddViewItem("ribao", 21, 23, "累计.4_8高炉合计.氮比");
            AddViewItem("ribao", 21, 24, "累计.4_8高炉合计.焦炭负荷");
            AddViewItem("ribao", 21, 25, "累计.4_8高炉合计.燃料比");
            AddViewItem("ribao", 21, 26, "累计.4_8高炉合计.废铁单耗");
            AddViewItem("ribao", 21, 30, "累计.4_8高炉合计.TRT发电");
            AddViewItem("ribao", 21, 31, "累计.4_8高炉合计.返矿");
            AddViewItem("ribao", 21, 32, "累计.4_8高炉合计.矿槽");
            AddViewItem("ribao", 21, 33, "累计.4_8高炉合计.休风工艺");
            AddViewItem("ribao", 21, 34, "累计.4_8高炉合计.休风设备");
            AddViewItem("ribao", 21, 35, "累计.4_8高炉合计.休风电");
            AddViewItem("ribao", 21, 36, "累计.4_8高炉合计.休风外围");
            AddViewItem("ribao", 21, 37, "累计.4_8高炉合计.休风计划");
            AddViewItem("ribao", 21, 38, "累计.4_8高炉合计.休风");

            AddViewItem("ribao", 24, 3, "本日.全铁产量");
            AddViewItem("ribao", 24, 4, "本日.合格铁");
            AddViewItem("ribao", 24, 5, "本日.号外铁");
            AddViewItem("ribao", 24, 6, "本日.铸造铁");
            AddViewItem("ribao", 24, 7, "本日.利用系数");
            AddViewItem("ribao", 24, 8, "本日.原料矿单耗");
            AddViewItem("ribao", 24, 9, "本日.干焦丁");
            AddViewItem("ribao", 24, 10, "本日.综合焦炭单耗");
            AddViewItem("ribao", 24, 11, "本日.入炉焦炭单耗");
            AddViewItem("ribao", 24, 12, "本日.煤粉单耗");
            AddViewItem("ribao", 24, 13, "本日.焦丁比");
            AddViewItem("ribao", 24, 14, "本日.冶炼强度");
            AddViewItem("ribao", 24, 15, "本日.合格率");
            AddViewItem("ribao", 24, 16, "本日.休风率");
            AddViewItem("ribao", 24, 17, "本日.平均风温");
            AddViewItem("ribao", 24, 18, "本日.耗电");
            AddViewItem("ribao", 24, 19, "本日.富氧率");
            AddViewItem("ribao", 24, 20, "本日.生铁含硅");
            AddViewItem("ribao", 24, 21, "本日.一级品率");
            AddViewItem("ribao", 24, 22, "本日.氧比");
            AddViewItem("ribao", 24, 23, "本日.氮比");
            AddViewItem("ribao", 24, 24, "本日.焦炭负荷");
            AddViewItem("ribao", 24, 25, "本日.燃料比");
            AddViewItem("ribao", 24, 26, "本日.废铁单耗");
            AddViewItem("ribao", 24, 30, "本日.TRT发电");
            AddViewItem("ribao", 24, 31, "本日.返矿");
            AddViewItem("ribao", 24, 32, "本日.矿槽");
            AddViewItem("ribao", 24, 33, "本日.休风工艺");
            AddViewItem("ribao", 24, 34, "本日.休风设备");
            AddViewItem("ribao", 24, 35, "本日.休风电");
            AddViewItem("ribao", 24, 36, "本日.休风外围");
            AddViewItem("ribao", 24, 37, "本日.休风计划");
            AddViewItem("ribao", 24, 38, "本日.休风");

            AddViewItem("ribao", 25, 3, "累计.全铁产量");
            AddViewItem("ribao", 25, 4, "累计.合格铁");
            AddViewItem("ribao", 25, 5, "累计.号外铁");
            AddViewItem("ribao", 25, 6, "累计.铸造铁");
            AddViewItem("ribao", 25, 7, "累计.利用系数");
            AddViewItem("ribao", 25, 8, "累计.原料矿单耗");
            AddViewItem("ribao", 25, 9, "累计.干焦丁");
            AddViewItem("ribao", 25, 10, "累计.综合焦炭单耗");
            AddViewItem("ribao", 25, 11, "累计.入炉焦炭单耗");
            AddViewItem("ribao", 25, 12, "累计.煤粉单耗");
            AddViewItem("ribao", 25, 13, "累计.焦丁比");
            AddViewItem("ribao", 25, 14, "累计.冶炼强度");
            AddViewItem("ribao", 25, 15, "累计.合格率");
            AddViewItem("ribao", 25, 16, "累计.休风率");
            AddViewItem("ribao", 25, 17, "累计.平均风温");
            AddViewItem("ribao", 25, 18, "累计.电单耗");
            AddViewItem("ribao", 25, 19, "累计.富氧率");
            AddViewItem("ribao", 25, 20, "累计.生铁含硅");
            AddViewItem("ribao", 25, 21, "累计.一级品率");
            AddViewItem("ribao", 25, 22, "累计.氧比");
            AddViewItem("ribao", 25, 23, "累计.氮比");
            AddViewItem("ribao", 25, 24, "累计.焦炭负荷");
            AddViewItem("ribao", 25, 25, "累计.燃料比");
            AddViewItem("ribao", 25, 26, "累计.废铁单耗");
            AddViewItem("ribao", 25, 30, "累计.TRT发电");
            AddViewItem("ribao", 25, 31, "累计.返矿");
            AddViewItem("ribao", 25, 32, "累计.矿槽");
            AddViewItem("ribao", 25, 33, "累计.休风工艺");
            AddViewItem("ribao", 25, 34, "累计.休风设备");
            AddViewItem("ribao", 25, 35, "累计.休风电");
            AddViewItem("ribao", 25, 36, "累计.休风外围");
            AddViewItem("ribao", 25, 37, "累计.休风计划");
            AddViewItem("ribao", 25, 38, "累计.休风");

            //中间分割
            AddViewItem("ribao", 28, 3, "本日.供应.烧结一车间");
            AddViewItem("ribao", 29, 3, "本日.供应.烧结二车间");
            AddViewItem("ribao", 30, 3, "本日.供应.230一车间");
            AddViewItem("ribao", 31, 3, "本日.供应.230二车间");
            AddViewItem("ribao", 32, 3, "本日.供应.260烧结");
            AddViewItem("ribao", 33, 3, "本日.供应.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 3, "本日.供应.球团一车间");
            AddViewItem("ribao", 35, 3, "本日.供应.球团二车间");
            AddViewItem("ribao", 36, 3, "本日.供应.北区球团");
            AddViewItem("ribao", 37, 3, "本日.供应.纽曼块");
            AddViewItem("ribao", 38, 3, "本日.供应.澳块");
            AddViewItem("ribao", 39, 3, "本日.供应.伊朗块");
            AddViewItem("ribao", 40, 3, "本日.供应.伊朗块");
            AddViewItem("ribao", 41, 3, "本日.供应.铁矿石");
            AddViewItem("ribao", 42, 3, "本日.供应.赛拉利昂块");
            AddViewItem("ribao", 43, 3, "本日.供应.钢粒");
            AddViewItem("ribao", 44, 3, "本日.供应.炉渣");
            AddViewItem("ribao", 45, 3, "本日.供应.水渣铁");
            AddViewItem("ribao", 46, 3, "本日.供应.破铁面");
            AddViewItem("ribao", 47, 3, "本日.供应.破铁块");
            AddViewItem("ribao", 48, 3, "本日.供应.微粉水渣铁");
            AddViewItem("ribao", 49, 3, "本日.供应.渣铁");
            AddViewItem("ribao", 50, 3, "本日.供应.外购焦粉");
            AddViewItem("ribao", 51, 3, "本日.供应.新焦化小焦");
            AddViewItem("ribao", 52, 3, "本日.供应.外购焦炭");
            AddViewItem("ribao", 53, 3, "本日.供应.自产焦炭");
            AddViewItem("ribao", 54, 3, "本日.供应.喷煤粉");
            AddViewItem("ribao", 55, 3, "本日.供应.烟煤粉");
            AddViewItem("ribao", 56, 3, "本日.供应.外购小焦");
            AddViewItem("ribao", 57, 3, "本日.供应.锰矿");
            AddViewItem("ribao", 58, 3, "本日.供应.萤石");
            AddViewItem("ribao", 59, 3, "本日.供应.硅石石子");

            AddViewItem("ribao", 28, 4, "累计.供应.烧结一车间");
            AddViewItem("ribao", 29, 4, "累计.供应.烧结二车间");
            AddViewItem("ribao", 30, 4, "累计.供应.230一车间");
            AddViewItem("ribao", 31, 4, "累计.供应.230二车间");
            AddViewItem("ribao", 32, 4, "累计.供应.260烧结");
            AddViewItem("ribao", 33, 4, "累计.供应.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 4, "累计.供应.球团一车间");
            AddViewItem("ribao", 35, 4, "累计.供应.球团二车间");
            AddViewItem("ribao", 36, 4, "累计.供应.北区球团");
            AddViewItem("ribao", 37, 4, "累计.供应.纽曼块");
            AddViewItem("ribao", 38, 4, "累计.供应.澳块");
            AddViewItem("ribao", 39, 4, "累计.供应.伊朗块");
            AddViewItem("ribao", 40, 4, "累计.供应.伊朗块");
            AddViewItem("ribao", 41, 4, "累计.供应.铁矿石");
            AddViewItem("ribao", 42, 4, "累计.供应.赛拉利昂块");
            AddViewItem("ribao", 43, 4, "累计.供应.钢粒");
            AddViewItem("ribao", 44, 4, "累计.供应.炉渣");
            AddViewItem("ribao", 45, 4, "累计.供应.水渣铁");
            AddViewItem("ribao", 46, 4, "累计.供应.破铁面");
            AddViewItem("ribao", 47, 4, "累计.供应.破铁块");
            AddViewItem("ribao", 48, 4, "累计.供应.微粉水渣铁");
            AddViewItem("ribao", 49, 4, "累计.供应.渣铁");
            AddViewItem("ribao", 50, 4, "累计.供应.外购焦粉");
            AddViewItem("ribao", 51, 4, "累计.供应.新焦化小焦");
            AddViewItem("ribao", 52, 4, "累计.供应.外购焦炭");
            AddViewItem("ribao", 53, 4, "累计.供应.自产焦炭");
            AddViewItem("ribao", 54, 4, "累计.供应.喷煤粉");
            AddViewItem("ribao", 55, 4, "累计.供应.烟煤粉");
            AddViewItem("ribao", 56, 4, "累计.供应.外购小焦");
            AddViewItem("ribao", 57, 4, "累计.供应.锰矿");
            AddViewItem("ribao", 58, 4, "累计.供应.萤石");
            AddViewItem("ribao", 59, 4, "累计.供应.硅石石子");

            AddViewItem("ribao", 28, 5, "本日.块矿面.烧结一车间");
            AddViewItem("ribao", 29, 5, "本日.块矿面.烧结二车间");
            AddViewItem("ribao", 30, 5, "本日.块矿面.230一车间");
            AddViewItem("ribao", 31, 5, "本日.块矿面.230二车间");
            AddViewItem("ribao", 32, 5, "本日.块矿面.260烧结");
            AddViewItem("ribao", 33, 5, "本日.块矿面.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 5, "本日.块矿面.球团一车间");
            AddViewItem("ribao", 35, 5, "本日.块矿面.球团二车间");
            AddViewItem("ribao", 36, 5, "本日.块矿面.北区球团");
            AddViewItem("ribao", 37, 5, "本日.块矿面.纽曼块");
            AddViewItem("ribao", 38, 5, "本日.块矿面.澳块");
            AddViewItem("ribao", 39, 5, "本日.块矿面.伊朗块");
            AddViewItem("ribao", 40, 5, "本日.块矿面.伊朗块");
            AddViewItem("ribao", 41, 5, "本日.块矿面.铁矿石");
            AddViewItem("ribao", 42, 5, "本日.块矿面.赛拉利昂块");
            AddViewItem("ribao", 43, 5, "本日.块矿面.钢粒");
            AddViewItem("ribao", 44, 5, "本日.块矿面.炉渣");
            AddViewItem("ribao", 45, 5, "本日.块矿面.水渣铁");
            AddViewItem("ribao", 46, 5, "本日.块矿面.破铁面");
            AddViewItem("ribao", 47, 5, "本日.块矿面.破铁块");
            AddViewItem("ribao", 48, 5, "本日.块矿面.微粉水渣铁");
            AddViewItem("ribao", 49, 5, "本日.块矿面.渣铁");
            AddViewItem("ribao", 50, 5, "本日.块矿面.外购焦粉");
            AddViewItem("ribao", 51, 5, "本日.块矿面.新焦化小焦");
            AddViewItem("ribao", 52, 5, "本日.块矿面.外购焦炭");
            AddViewItem("ribao", 53, 5, "本日.块矿面.自产焦炭");
            AddViewItem("ribao", 54, 5, "本日.块矿面.喷煤粉");
            AddViewItem("ribao", 55, 5, "本日.块矿面.烟煤粉");
            AddViewItem("ribao", 56, 5, "本日.块矿面.外购小焦");
            AddViewItem("ribao", 57, 5, "本日.块矿面.锰矿");
            AddViewItem("ribao", 58, 5, "本日.块矿面.萤石");
            AddViewItem("ribao", 59, 5, "本日.块矿面.硅石石子");

            AddViewItem("ribao", 28, 6, "累计.块矿面.烧结一车间");
            AddViewItem("ribao", 29, 6, "累计.块矿面.烧结二车间");
            AddViewItem("ribao", 30, 6, "累计.块矿面.230一车间");
            AddViewItem("ribao", 31, 6, "累计.块矿面.230二车间");
            AddViewItem("ribao", 32, 6, "累计.块矿面.260烧结");
            AddViewItem("ribao", 33, 6, "累计.块矿面.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 6, "累计.块矿面.球团一车间");
            AddViewItem("ribao", 35, 6, "累计.块矿面.球团二车间");
            AddViewItem("ribao", 36, 6, "累计.块矿面.北区球团");
            AddViewItem("ribao", 37, 6, "累计.块矿面.纽曼块");
            AddViewItem("ribao", 38, 6, "累计.块矿面.澳块");
            AddViewItem("ribao", 39, 6, "累计.块矿面.伊朗块");
            AddViewItem("ribao", 40, 6, "累计.块矿面.伊朗块");
            AddViewItem("ribao", 41, 6, "累计.块矿面.铁矿石");
            AddViewItem("ribao", 42, 6, "累计.块矿面.赛拉利昂块");
            AddViewItem("ribao", 43, 6, "累计.块矿面.钢粒");
            AddViewItem("ribao", 44, 6, "累计.块矿面.炉渣");
            AddViewItem("ribao", 45, 6, "累计.块矿面.水渣铁");
            AddViewItem("ribao", 46, 6, "累计.块矿面.破铁面");
            AddViewItem("ribao", 47, 6, "累计.块矿面.破铁块");
            AddViewItem("ribao", 48, 6, "累计.块矿面.微粉水渣铁");
            AddViewItem("ribao", 49, 6, "累计.块矿面.渣铁");
            AddViewItem("ribao", 50, 6, "累计.块矿面.外购焦粉");
            AddViewItem("ribao", 51, 6, "累计.块矿面.新焦化小焦");
            AddViewItem("ribao", 52, 6, "累计.块矿面.外购焦炭");
            AddViewItem("ribao", 53, 6, "累计.块矿面.自产焦炭");
            AddViewItem("ribao", 54, 6, "累计.块矿面.喷煤粉");
            AddViewItem("ribao", 55, 6, "累计.块矿面.烟煤粉");
            AddViewItem("ribao", 56, 6, "累计.块矿面.外购小焦");
            AddViewItem("ribao", 57, 6, "累计.块矿面.锰矿");
            AddViewItem("ribao", 58, 6, "累计.块矿面.萤石");
            AddViewItem("ribao", 59, 6, "累计.块矿面.硅石石子");

            AddViewItem("ribao", 28, 7, "本日.4高炉.烧结一车间");
            AddViewItem("ribao", 29, 7, "本日.4高炉.烧结二车间");
            AddViewItem("ribao", 30, 7, "本日.4高炉.230一车间");
            AddViewItem("ribao", 31, 7, "本日.4高炉.230二车间");
            AddViewItem("ribao", 32, 7, "本日.4高炉.260烧结");
            AddViewItem("ribao", 33, 7, "本日.4高炉.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 7, "本日.4高炉.球团一车间");
            AddViewItem("ribao", 35, 7, "本日.4高炉.球团二车间");
            AddViewItem("ribao", 36, 7, "本日.4高炉.北区球团");
            AddViewItem("ribao", 37, 7, "本日.4高炉.纽曼块");
            AddViewItem("ribao", 38, 7, "本日.4高炉.澳块");
            AddViewItem("ribao", 39, 7, "本日.4高炉.伊朗块");
            AddViewItem("ribao", 40, 7, "本日.4高炉.伊朗块");
            AddViewItem("ribao", 41, 7, "本日.4高炉.铁矿石");
            AddViewItem("ribao", 42, 7, "本日.4高炉.赛拉利昂块");
            AddViewItem("ribao", 43, 7, "本日.4高炉.钢粒");
            AddViewItem("ribao", 44, 7, "本日.4高炉.炉渣");
            AddViewItem("ribao", 45, 7, "本日.4高炉.水渣铁");
            AddViewItem("ribao", 46, 7, "本日.4高炉.破铁面");
            AddViewItem("ribao", 47, 7, "本日.4高炉.破铁块");
            AddViewItem("ribao", 48, 7, "本日.4高炉.微粉水渣铁");
            AddViewItem("ribao", 49, 7, "本日.4高炉.渣铁");
            AddViewItem("ribao", 50, 7, "本日.4高炉.外购焦粉");
            AddViewItem("ribao", 51, 7, "本日.4高炉.新焦化小焦");
            AddViewItem("ribao", 52, 7, "本日.4高炉.外购焦炭");
            AddViewItem("ribao", 53, 7, "本日.4高炉.自产焦炭");
            AddViewItem("ribao", 54, 7, "本日.4高炉.喷煤粉");
            AddViewItem("ribao", 55, 7, "本日.4高炉.烟煤粉");
            AddViewItem("ribao", 56, 7, "本日.4高炉.外购小焦");
            AddViewItem("ribao", 57, 7, "本日.4高炉.锰矿");
            AddViewItem("ribao", 58, 7, "本日.4高炉.萤石");
            AddViewItem("ribao", 59, 7, "本日.4高炉.硅石石子");

            AddViewItem("ribao", 28, 8, "累计.4高炉.烧结一车间");
            AddViewItem("ribao", 29, 8, "累计.4高炉.烧结二车间");
            AddViewItem("ribao", 30, 8, "累计.4高炉.230一车间");
            AddViewItem("ribao", 31, 8, "累计.4高炉.230二车间");
            AddViewItem("ribao", 32, 8, "累计.4高炉.260烧结");
            AddViewItem("ribao", 33, 8, "累计.4高炉.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 8, "累计.4高炉.球团一车间");
            AddViewItem("ribao", 35, 8, "累计.4高炉.球团二车间");
            AddViewItem("ribao", 36, 8, "累计.4高炉.北区球团");
            AddViewItem("ribao", 37, 8, "累计.4高炉.纽曼块");
            AddViewItem("ribao", 38, 8, "累计.4高炉.澳块");
            AddViewItem("ribao", 39, 8, "累计.4高炉.伊朗块");
            AddViewItem("ribao", 40, 8, "累计.4高炉.伊朗块");
            AddViewItem("ribao", 41, 8, "累计.4高炉.铁矿石");
            AddViewItem("ribao", 42, 8, "累计.4高炉.赛拉利昂块");
            AddViewItem("ribao", 43, 8, "累计.4高炉.钢粒");
            AddViewItem("ribao", 44, 8, "累计.4高炉.炉渣");
            AddViewItem("ribao", 45, 8, "累计.4高炉.水渣铁");
            AddViewItem("ribao", 46, 8, "累计.4高炉.破铁面");
            AddViewItem("ribao", 47, 8, "累计.4高炉.破铁块");
            AddViewItem("ribao", 48, 8, "累计.4高炉.微粉水渣铁");
            AddViewItem("ribao", 49, 8, "累计.4高炉.渣铁");
            AddViewItem("ribao", 50, 8, "累计.4高炉.外购焦粉");
            AddViewItem("ribao", 51, 8, "累计.4高炉.新焦化小焦");
            AddViewItem("ribao", 52, 8, "累计.4高炉.外购焦炭");
            AddViewItem("ribao", 53, 8, "累计.4高炉.自产焦炭");
            AddViewItem("ribao", 54, 8, "累计.4高炉.喷煤粉");
            AddViewItem("ribao", 55, 8, "累计.4高炉.烟煤粉");
            AddViewItem("ribao", 56, 8, "累计.4高炉.外购小焦");
            AddViewItem("ribao", 57, 8, "累计.4高炉.锰矿");
            AddViewItem("ribao", 58, 8, "累计.4高炉.萤石");
            AddViewItem("ribao", 59, 8, "累计.4高炉.硅石石子");


            AddViewItem("ribao", 28, 9, "本日.5高炉.烧结一车间");
            AddViewItem("ribao", 29, 9, "本日.5高炉.烧结二车间");
            AddViewItem("ribao", 30, 9, "本日.5高炉.230一车间");
            AddViewItem("ribao", 31, 9, "本日.5高炉.230二车间");
            AddViewItem("ribao", 32, 9, "本日.5高炉.260烧结");
            AddViewItem("ribao", 33, 9, "本日.5高炉.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 9, "本日.5高炉.球团一车间");
            AddViewItem("ribao", 35, 9, "本日.5高炉.球团二车间");
            AddViewItem("ribao", 36, 9, "本日.5高炉.北区球团");
            AddViewItem("ribao", 37, 9, "本日.5高炉.纽曼块");
            AddViewItem("ribao", 38, 9, "本日.5高炉.澳块");
            AddViewItem("ribao", 39, 9, "本日.5高炉.伊朗块");
            AddViewItem("ribao", 40, 9, "本日.5高炉.伊朗块");
            AddViewItem("ribao", 41, 9, "本日.5高炉.铁矿石");
            AddViewItem("ribao", 42, 9, "本日.5高炉.赛拉利昂块");
            AddViewItem("ribao", 43, 9, "本日.5高炉.钢粒");
            AddViewItem("ribao", 44, 9, "本日.5高炉.炉渣");
            AddViewItem("ribao", 45, 9, "本日.5高炉.水渣铁");
            AddViewItem("ribao", 46, 9, "本日.5高炉.破铁面");
            AddViewItem("ribao", 47, 9, "本日.5高炉.破铁块");
            AddViewItem("ribao", 48, 9, "本日.5高炉.微粉水渣铁");
            AddViewItem("ribao", 49, 9, "本日.5高炉.渣铁");
            AddViewItem("ribao", 50, 9, "本日.5高炉.外购焦粉");
            AddViewItem("ribao", 51, 9, "本日.5高炉.新焦化小焦");
            AddViewItem("ribao", 52, 9, "本日.5高炉.外购焦炭");
            AddViewItem("ribao", 53, 9, "本日.5高炉.自产焦炭");
            AddViewItem("ribao", 54, 9, "本日.5高炉.喷煤粉");
            AddViewItem("ribao", 55, 9, "本日.5高炉.烟煤粉");
            AddViewItem("ribao", 56, 9, "本日.5高炉.外购小焦");
            AddViewItem("ribao", 57, 9, "本日.5高炉.锰矿");
            AddViewItem("ribao", 58, 9, "本日.5高炉.萤石");
            AddViewItem("ribao", 59, 9, "本日.5高炉.硅石石子");

            AddViewItem("ribao", 28, 10, "累计.5高炉.烧结一车间");
            AddViewItem("ribao", 29, 10, "累计.5高炉.烧结二车间");
            AddViewItem("ribao", 30, 10, "累计.5高炉.230一车间");
            AddViewItem("ribao", 31, 10, "累计.5高炉.230二车间");
            AddViewItem("ribao", 32, 10, "累计.5高炉.260烧结");
            AddViewItem("ribao", 33, 10, "累计.5高炉.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 10, "累计.5高炉.球团一车间");
            AddViewItem("ribao", 35, 10, "累计.5高炉.球团二车间");
            AddViewItem("ribao", 36, 10, "累计.5高炉.北区球团");
            AddViewItem("ribao", 37, 10, "累计.5高炉.纽曼块");
            AddViewItem("ribao", 38, 10, "累计.5高炉.澳块");
            AddViewItem("ribao", 39, 10, "累计.5高炉.伊朗块");
            AddViewItem("ribao", 40, 10, "累计.5高炉.伊朗块");
            AddViewItem("ribao", 41, 10, "累计.5高炉.铁矿石");
            AddViewItem("ribao", 42, 10, "累计.5高炉.赛拉利昂块");
            AddViewItem("ribao", 43, 10, "累计.5高炉.钢粒");
            AddViewItem("ribao", 44, 10, "累计.5高炉.炉渣");
            AddViewItem("ribao", 45, 10, "累计.5高炉.水渣铁");
            AddViewItem("ribao", 46, 10, "累计.5高炉.破铁面");
            AddViewItem("ribao", 47, 10, "累计.5高炉.破铁块");
            AddViewItem("ribao", 48, 10, "累计.5高炉.微粉水渣铁");
            AddViewItem("ribao", 49, 10, "累计.5高炉.渣铁");
            AddViewItem("ribao", 50, 10, "累计.5高炉.外购焦粉");
            AddViewItem("ribao", 51, 10, "累计.5高炉.新焦化小焦");
            AddViewItem("ribao", 52, 10, "累计.5高炉.外购焦炭");
            AddViewItem("ribao", 53, 10, "累计.5高炉.自产焦炭");
            AddViewItem("ribao", 54, 10, "累计.5高炉.喷煤粉");
            AddViewItem("ribao", 55, 10, "累计.5高炉.烟煤粉");
            AddViewItem("ribao", 56, 10, "累计.5高炉.外购小焦");
            AddViewItem("ribao", 57, 10, "累计.5高炉.锰矿");
            AddViewItem("ribao", 58, 10, "累计.5高炉.萤石");
            AddViewItem("ribao", 59, 10, "累计.5高炉.硅石石子");

            AddViewItem("ribao", 28, 11, "本日.6高炉.烧结一车间");
            AddViewItem("ribao", 29, 11, "本日.6高炉.烧结二车间");
            AddViewItem("ribao", 30, 11, "本日.6高炉.230一车间");
            AddViewItem("ribao", 31, 11, "本日.6高炉.230二车间");
            AddViewItem("ribao", 32, 11, "本日.6高炉.260烧结");
            AddViewItem("ribao", 33, 11, "本日.6高炉.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 11, "本日.6高炉.球团一车间");
            AddViewItem("ribao", 35, 11, "本日.6高炉.球团二车间");
            AddViewItem("ribao", 36, 11, "本日.6高炉.北区球团");
            AddViewItem("ribao", 37, 11, "本日.6高炉.纽曼块");
            AddViewItem("ribao", 38, 11, "本日.6高炉.澳块");
            AddViewItem("ribao", 39, 11, "本日.6高炉.伊朗块");
            AddViewItem("ribao", 40, 11, "本日.6高炉.伊朗块");
            AddViewItem("ribao", 41, 11, "本日.6高炉.铁矿石");
            AddViewItem("ribao", 42, 11, "本日.6高炉.赛拉利昂块");
            AddViewItem("ribao", 43, 11, "本日.6高炉.钢粒");
            AddViewItem("ribao", 44, 11, "本日.6高炉.炉渣");
            AddViewItem("ribao", 45, 11, "本日.6高炉.水渣铁");
            AddViewItem("ribao", 46, 11, "本日.6高炉.破铁面");
            AddViewItem("ribao", 47, 11, "本日.6高炉.破铁块");
            AddViewItem("ribao", 48, 11, "本日.6高炉.微粉水渣铁");
            AddViewItem("ribao", 49, 11, "本日.6高炉.渣铁");
            AddViewItem("ribao", 50, 11, "本日.6高炉.外购焦粉");
            AddViewItem("ribao", 51, 11, "本日.6高炉.新焦化小焦");
            AddViewItem("ribao", 52, 11, "本日.6高炉.外购焦炭");
            AddViewItem("ribao", 53, 11, "本日.6高炉.自产焦炭");
            AddViewItem("ribao", 54, 11, "本日.6高炉.喷煤粉");
            AddViewItem("ribao", 55, 11, "本日.6高炉.烟煤粉");
            AddViewItem("ribao", 56, 11, "本日.6高炉.外购小焦");
            AddViewItem("ribao", 57, 11, "本日.6高炉.锰矿");
            AddViewItem("ribao", 58, 11, "本日.6高炉.萤石");
            AddViewItem("ribao", 59, 11, "本日.6高炉.硅石石子");

            AddViewItem("ribao", 28, 12, "累计.6高炉.烧结一车间");
            AddViewItem("ribao", 29, 12, "累计.6高炉.烧结二车间");
            AddViewItem("ribao", 30, 12, "累计.6高炉.230一车间");
            AddViewItem("ribao", 31, 12, "累计.6高炉.230二车间");
            AddViewItem("ribao", 32, 12, "累计.6高炉.260烧结");
            AddViewItem("ribao", 33, 12, "累计.6高炉.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 12, "累计.6高炉.球团一车间");
            AddViewItem("ribao", 35, 12, "累计.6高炉.球团二车间");
            AddViewItem("ribao", 36, 12, "累计.6高炉.北区球团");
            AddViewItem("ribao", 37, 12, "累计.6高炉.纽曼块");
            AddViewItem("ribao", 38, 12, "累计.6高炉.澳块");
            AddViewItem("ribao", 39, 12, "累计.6高炉.伊朗块");
            AddViewItem("ribao", 40, 12, "累计.6高炉.伊朗块");
            AddViewItem("ribao", 41, 12, "累计.6高炉.铁矿石");
            AddViewItem("ribao", 42, 12, "累计.6高炉.赛拉利昂块");
            AddViewItem("ribao", 43, 12, "累计.6高炉.钢粒");
            AddViewItem("ribao", 44, 12, "累计.6高炉.炉渣");
            AddViewItem("ribao", 45, 12, "累计.6高炉.水渣铁");
            AddViewItem("ribao", 46, 12, "累计.6高炉.破铁面");
            AddViewItem("ribao", 47, 12, "累计.6高炉.破铁块");
            AddViewItem("ribao", 48, 12, "累计.6高炉.微粉水渣铁");
            AddViewItem("ribao", 49, 12, "累计.6高炉.渣铁");
            AddViewItem("ribao", 50, 12, "累计.6高炉.外购焦粉");
            AddViewItem("ribao", 51, 12, "累计.6高炉.新焦化小焦");
            AddViewItem("ribao", 52, 12, "累计.6高炉.外购焦炭");
            AddViewItem("ribao", 53, 12, "累计.6高炉.自产焦炭");
            AddViewItem("ribao", 54, 12, "累计.6高炉.喷煤粉");
            AddViewItem("ribao", 55, 12, "累计.6高炉.烟煤粉");
            AddViewItem("ribao", 56, 12, "累计.6高炉.外购小焦");
            AddViewItem("ribao", 57, 12, "累计.6高炉.锰矿");
            AddViewItem("ribao", 58, 12, "累计.6高炉.萤石");
            AddViewItem("ribao", 59, 12, "累计.6高炉.硅石石子");

            AddViewItem("ribao", 28, 13, "本日.7高炉.烧结一车间");
            AddViewItem("ribao", 29, 13, "本日.7高炉.烧结二车间");
            AddViewItem("ribao", 30, 13, "本日.7高炉.230一车间");
            AddViewItem("ribao", 31, 13, "本日.7高炉.230二车间");
            AddViewItem("ribao", 32, 13, "本日.7高炉.260烧结");
            AddViewItem("ribao", 33, 13, "本日.7高炉.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 13, "本日.7高炉.球团一车间");
            AddViewItem("ribao", 35, 13, "本日.7高炉.球团二车间");
            AddViewItem("ribao", 36, 13, "本日.7高炉.北区球团");
            AddViewItem("ribao", 37, 13, "本日.7高炉.纽曼块");
            AddViewItem("ribao", 38, 13, "本日.7高炉.澳块");
            AddViewItem("ribao", 39, 13, "本日.7高炉.伊朗块");
            AddViewItem("ribao", 40, 13, "本日.7高炉.伊朗块");
            AddViewItem("ribao", 41, 13, "本日.7高炉.铁矿石");
            AddViewItem("ribao", 42, 13, "本日.7高炉.赛拉利昂块");
            AddViewItem("ribao", 43, 13, "本日.7高炉.钢粒");
            AddViewItem("ribao", 44, 13, "本日.7高炉.炉渣");
            AddViewItem("ribao", 45, 13, "本日.7高炉.水渣铁");
            AddViewItem("ribao", 46, 13, "本日.7高炉.破铁面");
            AddViewItem("ribao", 47, 13, "本日.7高炉.破铁块");
            AddViewItem("ribao", 48, 13, "本日.7高炉.微粉水渣铁");
            AddViewItem("ribao", 49, 13, "本日.7高炉.渣铁");
            AddViewItem("ribao", 50, 13, "本日.7高炉.外购焦粉");
            AddViewItem("ribao", 51, 13, "本日.7高炉.新焦化小焦");
            AddViewItem("ribao", 52, 13, "本日.7高炉.外购焦炭");
            AddViewItem("ribao", 53, 13, "本日.7高炉.自产焦炭");
            AddViewItem("ribao", 54, 13, "本日.7高炉.喷煤粉");
            AddViewItem("ribao", 55, 13, "本日.7高炉.烟煤粉");
            AddViewItem("ribao", 56, 13, "本日.7高炉.外购小焦");
            AddViewItem("ribao", 57, 13, "本日.7高炉.锰矿");
            AddViewItem("ribao", 58, 13, "本日.7高炉.萤石");
            AddViewItem("ribao", 59, 13, "本日.7高炉.硅石石子");

            AddViewItem("ribao", 28, 14, "累计.7高炉.烧结一车间");
            AddViewItem("ribao", 29, 14, "累计.7高炉.烧结二车间");
            AddViewItem("ribao", 30, 14, "累计.7高炉.230一车间");
            AddViewItem("ribao", 31, 14, "累计.7高炉.230二车间");
            AddViewItem("ribao", 32, 14, "累计.7高炉.260烧结");
            AddViewItem("ribao", 33, 14, "累计.7高炉.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 14, "累计.7高炉.球团一车间");
            AddViewItem("ribao", 35, 14, "累计.7高炉.球团二车间");
            AddViewItem("ribao", 36, 14, "累计.7高炉.北区球团");
            AddViewItem("ribao", 37, 14, "累计.7高炉.纽曼块");
            AddViewItem("ribao", 38, 14, "累计.7高炉.澳块");
            AddViewItem("ribao", 39, 14, "累计.7高炉.伊朗块");
            AddViewItem("ribao", 40, 14, "累计.7高炉.伊朗块");
            AddViewItem("ribao", 41, 14, "累计.7高炉.铁矿石");
            AddViewItem("ribao", 42, 14, "累计.7高炉.赛拉利昂块");
            AddViewItem("ribao", 43, 14, "累计.7高炉.钢粒");
            AddViewItem("ribao", 44, 14, "累计.7高炉.炉渣");
            AddViewItem("ribao", 45, 14, "累计.7高炉.水渣铁");
            AddViewItem("ribao", 46, 14, "累计.7高炉.破铁面");
            AddViewItem("ribao", 47, 14, "累计.7高炉.破铁块");
            AddViewItem("ribao", 48, 14, "累计.7高炉.微粉水渣铁");
            AddViewItem("ribao", 49, 14, "累计.7高炉.渣铁");
            AddViewItem("ribao", 50, 14, "累计.7高炉.外购焦粉");
            AddViewItem("ribao", 51, 14, "累计.7高炉.新焦化小焦");
            AddViewItem("ribao", 52, 14, "累计.7高炉.外购焦炭");
            AddViewItem("ribao", 53, 14, "累计.7高炉.自产焦炭");
            AddViewItem("ribao", 54, 14, "累计.7高炉.喷煤粉");
            AddViewItem("ribao", 55, 14, "累计.7高炉.烟煤粉");
            AddViewItem("ribao", 56, 14, "累计.7高炉.外购小焦");
            AddViewItem("ribao", 57, 14, "累计.7高炉.锰矿");
            AddViewItem("ribao", 58, 14, "累计.7高炉.萤石");
            AddViewItem("ribao", 59, 14, "累计.7高炉.硅石石子");


            AddViewItem("ribao", 28, 15, "本日.8高炉.烧结一车间");
            AddViewItem("ribao", 29, 15, "本日.8高炉.烧结二车间");
            AddViewItem("ribao", 30, 15, "本日.8高炉.230一车间");
            AddViewItem("ribao", 31, 15, "本日.8高炉.230二车间");
            AddViewItem("ribao", 32, 15, "本日.8高炉.260烧结");
            AddViewItem("ribao", 33, 15, "本日.8高炉.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 15, "本日.8高炉.球团一车间");
            AddViewItem("ribao", 35, 15, "本日.8高炉.球团二车间");
            AddViewItem("ribao", 36, 15, "本日.8高炉.北区球团");
            AddViewItem("ribao", 37, 15, "本日.8高炉.纽曼块");
            AddViewItem("ribao", 38, 15, "本日.8高炉.澳块");
            AddViewItem("ribao", 39, 15, "本日.8高炉.伊朗块");
            AddViewItem("ribao", 40, 15, "本日.8高炉.伊朗块");
            AddViewItem("ribao", 41, 15, "本日.8高炉.铁矿石");
            AddViewItem("ribao", 42, 15, "本日.8高炉.赛拉利昂块");
            AddViewItem("ribao", 43, 15, "本日.8高炉.钢粒");
            AddViewItem("ribao", 44, 15, "本日.8高炉.炉渣");
            AddViewItem("ribao", 45, 15, "本日.8高炉.水渣铁");
            AddViewItem("ribao", 46, 15, "本日.8高炉.破铁面");
            AddViewItem("ribao", 47, 15, "本日.8高炉.破铁块");
            AddViewItem("ribao", 48, 15, "本日.8高炉.微粉水渣铁");
            AddViewItem("ribao", 49, 15, "本日.8高炉.渣铁");
            AddViewItem("ribao", 50, 15, "本日.8高炉.外购焦粉");
            AddViewItem("ribao", 51, 15, "本日.8高炉.新焦化小焦");
            AddViewItem("ribao", 52, 15, "本日.8高炉.外购焦炭");
            AddViewItem("ribao", 53, 15, "本日.8高炉.自产焦炭");
            AddViewItem("ribao", 54, 15, "本日.8高炉.喷煤粉");
            AddViewItem("ribao", 55, 15, "本日.8高炉.烟煤粉");
            AddViewItem("ribao", 56, 15, "本日.8高炉.外购小焦");
            AddViewItem("ribao", 57, 15, "本日.8高炉.锰矿");
            AddViewItem("ribao", 58, 15, "本日.8高炉.萤石");
            AddViewItem("ribao", 59, 15, "本日.8高炉.硅石石子");

            AddViewItem("ribao", 28, 16, "累计.8高炉.烧结一车间");
            AddViewItem("ribao", 29, 16, "累计.8高炉.烧结二车间");
            AddViewItem("ribao", 30, 16, "累计.8高炉.230一车间");
            AddViewItem("ribao", 31, 16, "累计.8高炉.230二车间");
            AddViewItem("ribao", 32, 16, "累计.8高炉.260烧结");
            AddViewItem("ribao", 33, 16, "累计.8高炉.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 16, "累计.8高炉.球团一车间");
            AddViewItem("ribao", 35, 16, "累计.8高炉.球团二车间");
            AddViewItem("ribao", 36, 16, "累计.8高炉.北区球团");
            AddViewItem("ribao", 37, 16, "累计.8高炉.纽曼块");
            AddViewItem("ribao", 38, 16, "累计.8高炉.澳块");
            AddViewItem("ribao", 39, 16, "累计.8高炉.伊朗块");
            AddViewItem("ribao", 40, 16, "累计.8高炉.伊朗块");
            AddViewItem("ribao", 41, 16, "累计.8高炉.铁矿石");
            AddViewItem("ribao", 42, 16, "累计.8高炉.赛拉利昂块");
            AddViewItem("ribao", 43, 16, "累计.8高炉.钢粒");
            AddViewItem("ribao", 44, 16, "累计.8高炉.炉渣");
            AddViewItem("ribao", 45, 16, "累计.8高炉.水渣铁");
            AddViewItem("ribao", 46, 16, "累计.8高炉.破铁面");
            AddViewItem("ribao", 47, 16, "累计.8高炉.破铁块");
            AddViewItem("ribao", 48, 16, "累计.8高炉.微粉水渣铁");
            AddViewItem("ribao", 49, 16, "累计.8高炉.渣铁");
            AddViewItem("ribao", 50, 16, "累计.8高炉.外购焦粉");
            AddViewItem("ribao", 51, 16, "累计.8高炉.新焦化小焦");
            AddViewItem("ribao", 52, 16, "累计.8高炉.外购焦炭");
            AddViewItem("ribao", 53, 16, "累计.8高炉.自产焦炭");
            AddViewItem("ribao", 54, 16, "累计.8高炉.喷煤粉");
            AddViewItem("ribao", 55, 16, "累计.8高炉.烟煤粉");
            AddViewItem("ribao", 56, 16, "累计.8高炉.外购小焦");
            AddViewItem("ribao", 57, 16, "累计.8高炉.锰矿");
            AddViewItem("ribao", 58, 16, "累计.8高炉.萤石");
            AddViewItem("ribao", 59, 16, "累计.8高炉.硅石石子");


            AddViewItem("ribao", 28, 17, "本日.1高炉.烧结一车间");
            AddViewItem("ribao", 29, 17, "本日.1高炉.烧结二车间");
            AddViewItem("ribao", 30, 17, "本日.1高炉.230一车间");
            AddViewItem("ribao", 31, 17, "本日.1高炉.230二车间");
            AddViewItem("ribao", 32, 17, "本日.1高炉.260烧结");
            AddViewItem("ribao", 33, 17, "本日.1高炉.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 17, "本日.1高炉.球团一车间");
            AddViewItem("ribao", 35, 17, "本日.1高炉.球团二车间");
            AddViewItem("ribao", 36, 17, "本日.1高炉.北区球团");
            AddViewItem("ribao", 37, 17, "本日.1高炉.纽曼块");
            AddViewItem("ribao", 38, 17, "本日.1高炉.澳块");
            AddViewItem("ribao", 39, 17, "本日.1高炉.伊朗块");
            AddViewItem("ribao", 40, 17, "本日.1高炉.伊朗块");
            AddViewItem("ribao", 41, 17, "本日.1高炉.铁矿石");
            AddViewItem("ribao", 42, 17, "本日.1高炉.赛拉利昂块");
            AddViewItem("ribao", 43, 17, "本日.1高炉.钢粒");
            AddViewItem("ribao", 44, 17, "本日.1高炉.炉渣");
            AddViewItem("ribao", 45, 17, "本日.1高炉.水渣铁");
            AddViewItem("ribao", 46, 17, "本日.1高炉.破铁面");
            AddViewItem("ribao", 47, 17, "本日.1高炉.破铁块");
            AddViewItem("ribao", 48, 17, "本日.1高炉.微粉水渣铁");
            AddViewItem("ribao", 49, 17, "本日.1高炉.渣铁");
            AddViewItem("ribao", 50, 17, "本日.1高炉.外购焦粉");
            AddViewItem("ribao", 51, 17, "本日.1高炉.新焦化小焦");
            AddViewItem("ribao", 52, 17, "本日.1高炉.外购焦炭");
            AddViewItem("ribao", 53, 17, "本日.1高炉.自产焦炭");
            AddViewItem("ribao", 54, 17, "本日.1高炉.喷煤粉");
            AddViewItem("ribao", 55, 17, "本日.1高炉.烟煤粉");
            AddViewItem("ribao", 56, 17, "本日.1高炉.外购小焦");
            AddViewItem("ribao", 57, 17, "本日.1高炉.锰矿");
            AddViewItem("ribao", 58, 17, "本日.1高炉.萤石");
            AddViewItem("ribao", 59, 17, "本日.1高炉.硅石石子");

            AddViewItem("ribao", 28, 18, "累计.1高炉.烧结一车间");
            AddViewItem("ribao", 29, 18, "累计.1高炉.烧结二车间");
            AddViewItem("ribao", 30, 18, "累计.1高炉.230一车间");
            AddViewItem("ribao", 31, 18, "累计.1高炉.230二车间");
            AddViewItem("ribao", 32, 18, "累计.1高炉.260烧结");
            AddViewItem("ribao", 33, 18, "累计.1高炉.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 18, "累计.1高炉.球团一车间");
            AddViewItem("ribao", 35, 18, "累计.1高炉.球团二车间");
            AddViewItem("ribao", 36, 18, "累计.1高炉.北区球团");
            AddViewItem("ribao", 37, 18, "累计.1高炉.纽曼块");
            AddViewItem("ribao", 38, 18, "累计.1高炉.澳块");
            AddViewItem("ribao", 39, 18, "累计.1高炉.伊朗块");
            AddViewItem("ribao", 40, 18, "累计.1高炉.伊朗块");
            AddViewItem("ribao", 41, 18, "累计.1高炉.铁矿石");
            AddViewItem("ribao", 42, 18, "累计.1高炉.赛拉利昂块");
            AddViewItem("ribao", 43, 18, "累计.1高炉.钢粒");
            AddViewItem("ribao", 44, 18, "累计.1高炉.炉渣");
            AddViewItem("ribao", 45, 18, "累计.1高炉.水渣铁");
            AddViewItem("ribao", 46, 18, "累计.1高炉.破铁面");
            AddViewItem("ribao", 47, 18, "累计.1高炉.破铁块");
            AddViewItem("ribao", 48, 18, "累计.1高炉.微粉水渣铁");
            AddViewItem("ribao", 49, 18, "累计.1高炉.渣铁");
            AddViewItem("ribao", 50, 18, "累计.1高炉.外购焦粉");
            AddViewItem("ribao", 51, 18, "累计.1高炉.新焦化小焦");
            AddViewItem("ribao", 52, 18, "累计.1高炉.外购焦炭");
            AddViewItem("ribao", 53, 18, "累计.1高炉.自产焦炭");
            AddViewItem("ribao", 54, 18, "累计.1高炉.喷煤粉");
            AddViewItem("ribao", 55, 18, "累计.1高炉.烟煤粉");
            AddViewItem("ribao", 56, 18, "累计.1高炉.外购小焦");
            AddViewItem("ribao", 57, 18, "累计.1高炉.锰矿");
            AddViewItem("ribao", 58, 18, "累计.1高炉.萤石");
            AddViewItem("ribao", 59, 18, "累计.1高炉.硅石石子");



            AddViewItem("ribao", 28, 19, "本日.扣返矿.烧结一车间");
            AddViewItem("ribao", 29, 19, "本日.扣返矿.烧结二车间");
            AddViewItem("ribao", 30, 19, "本日.扣返矿.230一车间");
            AddViewItem("ribao", 31, 19, "本日.扣返矿.230二车间");
            AddViewItem("ribao", 32, 19, "本日.扣返矿.260烧结");
            AddViewItem("ribao", 33, 19, "本日.扣返矿.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 19, "本日.扣返矿.球团一车间");
            AddViewItem("ribao", 35, 19, "本日.扣返矿.球团二车间");
            AddViewItem("ribao", 36, 19, "本日.扣返矿.北区球团");
            AddViewItem("ribao", 37, 19, "本日.扣返矿.纽曼块");
            AddViewItem("ribao", 38, 19, "本日.扣返矿.澳块");
            AddViewItem("ribao", 39, 19, "本日.扣返矿.伊朗块");
            AddViewItem("ribao", 40, 19, "本日.扣返矿.伊朗块");
            AddViewItem("ribao", 41, 19, "本日.扣返矿.铁矿石");
            AddViewItem("ribao", 42, 19, "本日.扣返矿.赛拉利昂块");
            AddViewItem("ribao", 43, 19, "本日.扣返矿.钢粒");
            AddViewItem("ribao", 44, 19, "本日.扣返矿.炉渣");
            AddViewItem("ribao", 45, 19, "本日.扣返矿.水渣铁");
            AddViewItem("ribao", 46, 19, "本日.扣返矿.破铁面");
            AddViewItem("ribao", 47, 19, "本日.扣返矿.破铁块");
            AddViewItem("ribao", 48, 19, "本日.扣返矿.微粉水渣铁");
            AddViewItem("ribao", 49, 19, "本日.扣返矿.渣铁");
            AddViewItem("ribao", 50, 19, "本日.扣返矿.外购焦粉");
            AddViewItem("ribao", 51, 19, "本日.扣返矿.新焦化小焦");
            AddViewItem("ribao", 52, 19, "本日.扣返矿.外购焦炭");
            AddViewItem("ribao", 53, 19, "本日.扣返矿.自产焦炭");
            AddViewItem("ribao", 54, 19, "本日.扣返矿.喷煤粉");
            AddViewItem("ribao", 55, 19, "本日.扣返矿.烟煤粉");
            AddViewItem("ribao", 56, 19, "本日.扣返矿.外购小焦");
            AddViewItem("ribao", 57, 19, "本日.扣返矿.锰矿");
            AddViewItem("ribao", 58, 19, "本日.扣返矿.萤石");
            AddViewItem("ribao", 59, 19, "本日.扣返矿.硅石石子");

            AddViewItem("ribao", 28, 20, "累计.扣返矿.烧结一车间");
            AddViewItem("ribao", 29, 20, "累计.扣返矿.烧结二车间");
            AddViewItem("ribao", 30, 20, "累计.扣返矿.230一车间");
            AddViewItem("ribao", 31, 20, "累计.扣返矿.230二车间");
            AddViewItem("ribao", 32, 20, "累计.扣返矿.260烧结");
            AddViewItem("ribao", 33, 20, "累计.扣返矿.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 20, "累计.扣返矿.球团一车间");
            AddViewItem("ribao", 35, 20, "累计.扣返矿.球团二车间");
            AddViewItem("ribao", 36, 20, "累计.扣返矿.北区球团");
            AddViewItem("ribao", 37, 20, "累计.扣返矿.纽曼块");
            AddViewItem("ribao", 38, 20, "累计.扣返矿.澳块");
            AddViewItem("ribao", 39, 20, "累计.扣返矿.伊朗块");
            AddViewItem("ribao", 40, 20, "累计.扣返矿.伊朗块");
            AddViewItem("ribao", 41, 20, "累计.扣返矿.铁矿石");
            AddViewItem("ribao", 42, 20, "累计.扣返矿.赛拉利昂块");
            AddViewItem("ribao", 43, 20, "累计.扣返矿.钢粒");
            AddViewItem("ribao", 44, 20, "累计.扣返矿.炉渣");
            AddViewItem("ribao", 45, 20, "累计.扣返矿.水渣铁");
            AddViewItem("ribao", 46, 20, "累计.扣返矿.破铁面");
            AddViewItem("ribao", 47, 20, "累计.扣返矿.破铁块");
            AddViewItem("ribao", 48, 20, "累计.扣返矿.微粉水渣铁");
            AddViewItem("ribao", 49, 20, "累计.扣返矿.渣铁");
            AddViewItem("ribao", 50, 20, "累计.扣返矿.外购焦粉");
            AddViewItem("ribao", 51, 20, "累计.扣返矿.新焦化小焦");
            AddViewItem("ribao", 52, 20, "累计.扣返矿.外购焦炭");
            AddViewItem("ribao", 53, 20, "累计.扣返矿.自产焦炭");
            AddViewItem("ribao", 54, 20, "累计.扣返矿.喷煤粉");
            AddViewItem("ribao", 55, 20, "累计.扣返矿.烟煤粉");
            AddViewItem("ribao", 56, 20, "累计.扣返矿.外购小焦");
            AddViewItem("ribao", 57, 20, "累计.扣返矿.锰矿");
            AddViewItem("ribao", 58, 20, "累计.扣返矿.萤石");
            AddViewItem("ribao", 59, 20, "累计.扣返矿.硅石石子");



            AddViewItem("ribao", 28, 21, "本日.烧结一车间");
            AddViewItem("ribao", 29, 21, "本日.烧结二车间");
            AddViewItem("ribao", 30, 21, "本日.230一车间");
            AddViewItem("ribao", 31, 21, "本日.230二车间");
            AddViewItem("ribao", 32, 21, "本日.260烧结");
            AddViewItem("ribao", 33, 21, "本日.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 21, "本日.球团一车间");
            AddViewItem("ribao", 35, 21, "本日.球团二车间");
            AddViewItem("ribao", 36, 21, "本日.北区球团");
            AddViewItem("ribao", 37, 21, "本日.纽曼块");
            AddViewItem("ribao", 38, 21, "本日.澳块");
            AddViewItem("ribao", 39, 21, "本日.伊朗块");
            AddViewItem("ribao", 40, 21, "本日.伊朗块");
            AddViewItem("ribao", 41, 21, "本日.铁矿石");
            AddViewItem("ribao", 42, 21, "本日.赛拉利昂块");
            AddViewItem("ribao", 43, 21, "本日.钢粒");
            AddViewItem("ribao", 44, 21, "本日.炉渣");
            AddViewItem("ribao", 45, 21, "本日.水渣铁");
            AddViewItem("ribao", 46, 21, "本日.破铁面");
            AddViewItem("ribao", 47, 21, "本日.破铁块");
            AddViewItem("ribao", 48, 21, "本日.微粉水渣铁");
            AddViewItem("ribao", 49, 21, "本日.渣铁");
            AddViewItem("ribao", 50, 21, "本日.外购焦粉");
            AddViewItem("ribao", 51, 21, "本日.新焦化小焦");
            AddViewItem("ribao", 52, 21, "本日.外购焦炭");
            AddViewItem("ribao", 53, 21, "本日.自产焦炭");
            AddViewItem("ribao", 54, 21, "本日.喷煤粉");
            AddViewItem("ribao", 55, 21, "本日.烟煤粉");
            AddViewItem("ribao", 56, 21, "本日.外购小焦");
            AddViewItem("ribao", 57, 21, "本日.锰矿");
            AddViewItem("ribao", 58, 21, "本日.萤石");
            AddViewItem("ribao", 59, 21, "本日.硅石石子");

            AddViewItem("ribao", 28, 22, "累计.烧结一车间");
            AddViewItem("ribao", 29, 22, "累计.烧结二车间");
            AddViewItem("ribao", 30, 22, "累计.230一车间");
            AddViewItem("ribao", 31, 22, "累计.230二车间");
            AddViewItem("ribao", 32, 22, "累计.260烧结");
            AddViewItem("ribao", 33, 22, "累计.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 22, "累计.球团一车间");
            AddViewItem("ribao", 35, 22, "累计.球团二车间");
            AddViewItem("ribao", 36, 22, "累计.北区球团");
            AddViewItem("ribao", 37, 22, "累计.纽曼块");
            AddViewItem("ribao", 38, 22, "累计.澳块");
            AddViewItem("ribao", 39, 22, "累计.伊朗块");
            AddViewItem("ribao", 40, 22, "累计.伊朗块");
            AddViewItem("ribao", 41, 22, "累计.铁矿石");
            AddViewItem("ribao", 42, 22, "累计.赛拉利昂块");
            AddViewItem("ribao", 43, 22, "累计.钢粒");
            AddViewItem("ribao", 44, 22, "累计.炉渣");
            AddViewItem("ribao", 45, 22, "累计.水渣铁");
            AddViewItem("ribao", 46, 22, "累计.破铁面");
            AddViewItem("ribao", 47, 22, "累计.破铁块");
            AddViewItem("ribao", 48, 22, "累计.微粉水渣铁");
            AddViewItem("ribao", 49, 22, "累计.渣铁");
            AddViewItem("ribao", 50, 22, "累计.外购焦粉");
            AddViewItem("ribao", 51, 22, "累计.新焦化小焦");
            AddViewItem("ribao", 52, 22, "累计.外购焦炭");
            AddViewItem("ribao", 53, 22, "累计.自产焦炭");
            AddViewItem("ribao", 54, 22, "累计.喷煤粉");
            AddViewItem("ribao", 55, 22, "累计.烟煤粉");
            AddViewItem("ribao", 56, 22, "累计.外购小焦");
            AddViewItem("ribao", 57, 22, "累计.锰矿");
            AddViewItem("ribao", 58, 22, "累计.萤石");
            AddViewItem("ribao", 59, 22, "累计.硅石石子");



            AddViewItem("ribao", 28, 23, "本日库存.1高炉.烧结一车间");
            AddViewItem("ribao", 29, 23, "本日库存.1高炉.烧结二车间");
            AddViewItem("ribao", 30, 23, "本日库存.1高炉.230一车间");
            AddViewItem("ribao", 31, 23, "本日库存.1高炉.230二车间");
            AddViewItem("ribao", 32, 23, "本日库存.1高炉.260烧结");
            AddViewItem("ribao", 33, 23, "本日库存.1高炉.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 23, "本日库存.1高炉.球团一车间");
            AddViewItem("ribao", 35, 23, "本日库存.1高炉.球团二车间");
            AddViewItem("ribao", 36, 23, "本日库存.1高炉.北区球团");
            AddViewItem("ribao", 37, 23, "本日库存.1高炉.纽曼块");
            AddViewItem("ribao", 38, 23, "本日库存.1高炉.澳块");
            AddViewItem("ribao", 39, 23, "本日库存.1高炉.伊朗块");
            AddViewItem("ribao", 40, 23, "本日库存.1高炉.伊朗块");
            AddViewItem("ribao", 41, 23, "本日库存.1高炉.铁矿石");
            AddViewItem("ribao", 42, 23, "本日库存.1高炉.赛拉利昂块");
            AddViewItem("ribao", 43, 23, "本日库存.1高炉.钢粒");
            AddViewItem("ribao", 44, 23, "本日库存.1高炉.炉渣");
            AddViewItem("ribao", 45, 23, "本日库存.1高炉.水渣铁");
            AddViewItem("ribao", 46, 23, "本日库存.1高炉.破铁面");
            AddViewItem("ribao", 47, 23, "本日库存.1高炉.破铁块");
            AddViewItem("ribao", 48, 23, "本日库存.1高炉.微粉水渣铁");
            AddViewItem("ribao", 49, 23, "本日库存.1高炉.渣铁");
            AddViewItem("ribao", 50, 23, "本日库存.1高炉.外购焦粉");
            AddViewItem("ribao", 51, 23, "本日库存.1高炉.新焦化小焦");
            AddViewItem("ribao", 52, 23, "本日库存.1高炉.外购焦炭");
            AddViewItem("ribao", 53, 23, "本日库存.1高炉.自产焦炭");
            AddViewItem("ribao", 54, 23, "本日库存.1高炉.喷煤粉");
            AddViewItem("ribao", 55, 23, "本日库存.1高炉.烟煤粉");
            AddViewItem("ribao", 56, 23, "本日库存.1高炉.外购小焦");
            AddViewItem("ribao", 57, 23, "本日库存.1高炉.锰矿");
            AddViewItem("ribao", 58, 23, "本日库存.1高炉.萤石");
            AddViewItem("ribao", 59, 23, "本日库存.1高炉.硅石石子");

            AddViewItem("ribao", 28, 24, "本日库存.二车间.烧结一车间");
            AddViewItem("ribao", 29, 24, "本日库存.二车间.烧结二车间");
            AddViewItem("ribao", 30, 24, "本日库存.二车间.230一车间");
            AddViewItem("ribao", 31, 24, "本日库存.二车间.230二车间");
            AddViewItem("ribao", 32, 24, "本日库存.二车间.260烧结");
            AddViewItem("ribao", 33, 24, "本日库存.二车间.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 24, "本日库存.二车间.球团一车间");
            AddViewItem("ribao", 35, 24, "本日库存.二车间.球团二车间");
            AddViewItem("ribao", 36, 24, "本日库存.二车间.北区球团");
            AddViewItem("ribao", 37, 24, "本日库存.二车间.纽曼块");
            AddViewItem("ribao", 38, 24, "本日库存.二车间.澳块");
            AddViewItem("ribao", 39, 24, "本日库存.二车间.伊朗块");
            AddViewItem("ribao", 40, 24, "本日库存.二车间.伊朗块");
            AddViewItem("ribao", 41, 24, "本日库存.二车间.铁矿石");
            AddViewItem("ribao", 42, 24, "本日库存.二车间.赛拉利昂块");
            AddViewItem("ribao", 43, 24, "本日库存.二车间.钢粒");
            AddViewItem("ribao", 44, 24, "本日库存.二车间.炉渣");
            AddViewItem("ribao", 45, 24, "本日库存.二车间.水渣铁");
            AddViewItem("ribao", 46, 24, "本日库存.二车间.破铁面");
            AddViewItem("ribao", 47, 24, "本日库存.二车间.破铁块");
            AddViewItem("ribao", 48, 24, "本日库存.二车间.微粉水渣铁");
            AddViewItem("ribao", 49, 24, "本日库存.二车间.渣铁");
            AddViewItem("ribao", 50, 24, "本日库存.二车间.外购焦粉");
            AddViewItem("ribao", 51, 24, "本日库存.二车间.新焦化小焦");
            AddViewItem("ribao", 52, 24, "本日库存.二车间.外购焦炭");
            AddViewItem("ribao", 53, 24, "本日库存.二车间.自产焦炭");
            AddViewItem("ribao", 54, 24, "本日库存.二车间.喷煤粉");
            AddViewItem("ribao", 55, 24, "本日库存.二车间.烟煤粉");
            AddViewItem("ribao", 56, 24, "本日库存.二车间.外购小焦");
            AddViewItem("ribao", 57, 24, "本日库存.二车间.锰矿");
            AddViewItem("ribao", 58, 24, "本日库存.二车间.萤石");
            AddViewItem("ribao", 59, 24, "本日库存.二车间.硅石石子");


            AddViewItem("ribao", 28, 25, "本日库存.三车间.烧结一车间");
            AddViewItem("ribao", 29, 25, "本日库存.三车间.烧结二车间");
            AddViewItem("ribao", 30, 25, "本日库存.三车间.230一车间");
            AddViewItem("ribao", 31, 25, "本日库存.三车间.230二车间");
            AddViewItem("ribao", 32, 25, "本日库存.三车间.260烧结");
            AddViewItem("ribao", 33, 25, "本日库存.三车间.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 25, "本日库存.三车间.球团一车间");
            AddViewItem("ribao", 35, 25, "本日库存.三车间.球团二车间");
            AddViewItem("ribao", 36, 25, "本日库存.三车间.北区球团");
            AddViewItem("ribao", 37, 25, "本日库存.三车间.纽曼块");
            AddViewItem("ribao", 38, 25, "本日库存.三车间.澳块");
            AddViewItem("ribao", 39, 25, "本日库存.三车间.伊朗块");
            AddViewItem("ribao", 40, 25, "本日库存.三车间.伊朗块");
            AddViewItem("ribao", 41, 25, "本日库存.三车间.铁矿石");
            AddViewItem("ribao", 42, 25, "本日库存.三车间.赛拉利昂块");
            AddViewItem("ribao", 43, 25, "本日库存.三车间.钢粒");
            AddViewItem("ribao", 44, 25, "本日库存.三车间.炉渣");
            AddViewItem("ribao", 45, 25, "本日库存.三车间.水渣铁");
            AddViewItem("ribao", 46, 25, "本日库存.三车间.破铁面");
            AddViewItem("ribao", 47, 25, "本日库存.三车间.破铁块");
            AddViewItem("ribao", 48, 25, "本日库存.三车间.微粉水渣铁");
            AddViewItem("ribao", 49, 25, "本日库存.三车间.渣铁");
            AddViewItem("ribao", 50, 25, "本日库存.三车间.外购焦粉");
            AddViewItem("ribao", 51, 25, "本日库存.三车间.新焦化小焦");
            AddViewItem("ribao", 52, 25, "本日库存.三车间.外购焦炭");
            AddViewItem("ribao", 53, 25, "本日库存.三车间.自产焦炭");
            AddViewItem("ribao", 54, 25, "本日库存.三车间.喷煤粉");
            AddViewItem("ribao", 55, 25, "本日库存.三车间.烟煤粉");
            AddViewItem("ribao", 56, 25, "本日库存.三车间.外购小焦");
            AddViewItem("ribao", 57, 25, "本日库存.三车间.锰矿");
            AddViewItem("ribao", 58, 25, "本日库存.三车间.萤石");
            AddViewItem("ribao", 59, 25, "本日库存.三车间.硅石石子");


            AddViewItem("ribao", 28, 26, "本日库存.烧结一车间");
            AddViewItem("ribao", 29, 26, "本日库存.烧结二车间");
            AddViewItem("ribao", 30, 26, "本日库存.230一车间");
            AddViewItem("ribao", 31, 26, "本日库存.230二车间");
            AddViewItem("ribao", 32, 26, "本日库存.260烧结");
            AddViewItem("ribao", 33, 26, "本日库存.第二炼铁厂烧结");
            AddViewItem("ribao", 34, 26, "本日库存.球团一车间");
            AddViewItem("ribao", 35, 26, "本日库存.球团二车间");
            AddViewItem("ribao", 36, 26, "本日库存.北区球团");
            AddViewItem("ribao", 37, 26, "本日库存.纽曼块");
            AddViewItem("ribao", 38, 26, "本日库存.澳块");
            AddViewItem("ribao", 39, 26, "本日库存.伊朗块");
            AddViewItem("ribao", 40, 26, "本日库存.伊朗块");
            AddViewItem("ribao", 41, 26, "本日库存.铁矿石");
            AddViewItem("ribao", 42, 26, "本日库存.赛拉利昂块");
            AddViewItem("ribao", 43, 26, "本日库存.钢粒");
            AddViewItem("ribao", 44, 26, "本日库存.炉渣");
            AddViewItem("ribao", 45, 26, "本日库存.水渣铁");
            AddViewItem("ribao", 46, 26, "本日库存.破铁面");
            AddViewItem("ribao", 47, 26, "本日库存.破铁块");
            AddViewItem("ribao", 48, 26, "本日库存.微粉水渣铁");
            AddViewItem("ribao", 49, 26, "本日库存.渣铁");
            AddViewItem("ribao", 50, 26, "本日库存.外购焦粉");
            AddViewItem("ribao", 51, 26, "本日库存.新焦化小焦");
            AddViewItem("ribao", 52, 26, "本日库存.外购焦炭");
            AddViewItem("ribao", 53, 26, "本日库存.自产焦炭");
            AddViewItem("ribao", 54, 26, "本日库存.喷煤粉");
            AddViewItem("ribao", 55, 26, "本日库存.烟煤粉");
            AddViewItem("ribao", 56, 26, "本日库存.外购小焦");
            AddViewItem("ribao", 57, 26, "本日库存.锰矿");
            AddViewItem("ribao", 58, 26, "本日库存.萤石");
            AddViewItem("ribao", 59, 26, "本日库存.硅石石子");


            /////////////////////////////
            //下面左右分割
            AddViewItem("ribao", 28, 29, "入炉矿化验分析.全厂合计1.TFe");
            AddViewItem("ribao", 28, 30, "入炉矿化验分析.全厂合计1.FeO");
            AddViewItem("ribao", 28, 31, "入炉矿化验分析.全厂合计1.S");
            AddViewItem("ribao", 28, 32, "入炉矿化验分析.全厂合计1.R");
            AddViewItem("ribao", 28, 33, "入炉矿化验分析.全厂合计1.块矿品位");
            AddViewItem("ribao", 29, 29, "入炉矿化验分析.全厂合计2.TFe");
            AddViewItem("ribao", 29, 30, "入炉矿化验分析.全厂合计2.FeO");
            AddViewItem("ribao", 29, 31, "入炉矿化验分析.全厂合计2.S");
            AddViewItem("ribao", 29, 32, "入炉矿化验分析.全厂合计2.R");
            AddViewItem("ribao", 29, 33, "入炉矿化验分析.全厂合计2.块矿品位");
            AddViewItem("ribao", 30, 29, "入炉矿化验分析.二车间1.TFe");
            AddViewItem("ribao", 30, 30, "入炉矿化验分析.二车间1.FeO");
            AddViewItem("ribao", 30, 31, "入炉矿化验分析.二车间1.S");
            AddViewItem("ribao", 30, 32, "入炉矿化验分析.二车间1.R");
            AddViewItem("ribao", 30, 33, "入炉矿化验分析.二车间1.块矿品位");
            AddViewItem("ribao", 31, 29, "入炉矿化验分析.二车间2.TFe");
            AddViewItem("ribao", 31, 30, "入炉矿化验分析.二车间2.FeO");
            AddViewItem("ribao", 31, 31, "入炉矿化验分析.二车间2.S");
            AddViewItem("ribao", 31, 32, "入炉矿化验分析.二车间2.R");
            AddViewItem("ribao", 31, 33, "入炉矿化验分析.二车间2.块矿品位");
            AddViewItem("ribao", 32, 29, "入炉矿化验分析.78号炉1.TFe");
            AddViewItem("ribao", 32, 30, "入炉矿化验分析.78号炉1.FeO");
            AddViewItem("ribao", 32, 31, "入炉矿化验分析.78号炉1.S");
            AddViewItem("ribao", 32, 32, "入炉矿化验分析.78号炉1.R");
            AddViewItem("ribao", 32, 33, "入炉矿化验分析.78号炉1.块矿品位");
            AddViewItem("ribao", 33, 29, "入炉矿化验分析.78号炉2.TFe");
            AddViewItem("ribao", 33, 30, "入炉矿化验分析.78号炉2.FeO");
            AddViewItem("ribao", 33, 31, "入炉矿化验分析.78号炉2.S");
            AddViewItem("ribao", 33, 32, "入炉矿化验分析.78号炉2.R");
            AddViewItem("ribao", 33, 33, "入炉矿化验分析.78号炉2.块矿品位");
            AddViewItem("ribao", 34, 29, "入炉矿化验分析.1高炉1.TFe");
            AddViewItem("ribao", 34, 30, "入炉矿化验分析.1高炉1.FeO");
            AddViewItem("ribao", 34, 31, "入炉矿化验分析.1高炉1.S");
            AddViewItem("ribao", 34, 32, "入炉矿化验分析.1高炉1.R");
            AddViewItem("ribao", 34, 33, "入炉矿化验分析.1高炉1.块矿品位");
            AddViewItem("ribao", 35, 29, "入炉矿化验分析.1高炉2.TFe");
            AddViewItem("ribao", 35, 30, "入炉矿化验分析.1高炉2.FeO");
            AddViewItem("ribao", 35, 31, "入炉矿化验分析.1高炉2.S");
            AddViewItem("ribao", 35, 32, "入炉矿化验分析.1高炉2.R");
            AddViewItem("ribao", 35, 33, "入炉矿化验分析.1高炉2.块矿品位");
            AddViewItem("ribao", 28, 34, "焦炭入炉分析.全厂合计1.C固");
            AddViewItem("ribao", 28, 35, "焦炭入炉分析.全厂合计1.A");
            AddViewItem("ribao", 28, 36, "焦炭入炉分析.1高炉1.S");
            AddViewItem("ribao", 28, 37, "焦炭入炉分析.1高炉1.V");
            AddViewItem("ribao", 28, 38, "焦炭入炉分析.全厂合计1.H2O");
            AddViewItem("ribao", 29, 34, "焦炭入炉分析.全厂合计2.C固");
            AddViewItem("ribao", 29, 35, "焦炭入炉分析.全厂合计2.A");
            AddViewItem("ribao", 29, 36, "焦炭入炉分析.全厂合计2.S");
            AddViewItem("ribao", 29, 37, "焦炭入炉分析.全厂合计2.V");
            AddViewItem("ribao", 29, 38, "焦炭入炉分析.全厂合计2.H2O");
            AddViewItem("ribao", 30, 34, "焦炭入炉分析.二车间1.C固");
            AddViewItem("ribao", 30, 35, "焦炭入炉分析.二车间1.A");
            AddViewItem("ribao", 30, 36, "焦炭入炉分析.二车间1.S");
            AddViewItem("ribao", 30, 37, "焦炭入炉分析.二车间1.V");
            AddViewItem("ribao", 30, 38, "焦炭入炉分析.二车间1.H2O");
            AddViewItem("ribao", 31, 34, "焦炭入炉分析.二车间2.C固");
            AddViewItem("ribao", 31, 35, "焦炭入炉分析.二车间2.A");
            AddViewItem("ribao", 31, 36, "焦炭入炉分析.二车间2.S");
            AddViewItem("ribao", 31, 37, "焦炭入炉分析.二车间2.V");
            AddViewItem("ribao", 31, 38, "焦炭入炉分析.二车间2.H2O");
            AddViewItem("ribao", 32, 34, "焦炭入炉分析.78号炉1.C固");
            AddViewItem("ribao", 33, 34, "焦炭入炉分析.78号炉2.C固");
            AddViewItem("ribao", 32, 35, "焦炭入炉分析.78号炉1.A");
            AddViewItem("ribao", 33, 35, "焦炭入炉分析.78号炉2.A");
            AddViewItem("ribao", 32, 36, "焦炭入炉分析.78号炉2.S");
            AddViewItem("ribao", 33, 36, "焦炭入炉分析.78号炉2.S");
            AddViewItem("ribao", 32, 37, "焦炭入炉分析.78号炉1.V");
            AddViewItem("ribao", 33, 37, "焦炭入炉分析.78号炉2.V");
            AddViewItem("ribao", 32, 38, "焦炭入炉分析.78号炉1.H2O");
            AddViewItem("ribao", 33, 38, "焦炭入炉分析.78号炉2.H2O");
            AddViewItem("ribao", 34, 34, "焦炭入炉分析.1高炉1.C固");
            AddViewItem("ribao", 35, 34, "焦炭入炉分析.1高炉2.C固");
            AddViewItem("ribao", 34, 35, "焦炭入炉分析.1高炉1.A");
            AddViewItem("ribao", 35, 35, "焦炭入炉分析.1高炉2.A");
            AddViewItem("ribao", 34, 36, "焦炭入炉分析.1高炉1.S");
            AddViewItem("ribao", 35, 36, "焦炭入炉分析.1高炉2.S");
            AddViewItem("ribao", 34, 37, "焦炭入炉分析.1高炉1.V");
            AddViewItem("ribao", 35, 37, "焦炭入炉分析.1高炉2.V");
            AddViewItem("ribao", 34, 38, "焦炭入炉分析.1高炉1.H2O");
            AddViewItem("ribao", 35, 38, "焦炭入炉分析.1高炉2.H2O");
            AddViewItem("ribao", 41, 28, "外购焦炭质量.质量1.C固");
            AddViewItem("ribao", 41, 29, "外购焦炭质量.质量1.A");
            AddViewItem("ribao", 41, 30, "外购焦炭质量.质量1.S");
            AddViewItem("ribao", 41, 31, "外购焦炭质量.质量1.V");
            AddViewItem("ribao", 41, 32, "外购焦炭质量.质量1.M25");
            AddViewItem("ribao", 41, 33, "外购焦炭质量.质量1.M10");
            AddViewItem("ribao", 41, 34, "外购焦炭质量.质量1.H2O");
            AddViewItem("ribao", 42, 28, "外购焦炭质量.质量2.C固");
            AddViewItem("ribao", 42, 29, "外购焦炭质量.质量2.A");
            AddViewItem("ribao", 42, 30, "外购焦炭质量.质量2.S");
            AddViewItem("ribao", 42, 31, "外购焦炭质量.质量2.V");
            AddViewItem("ribao", 42, 32, "外购焦炭质量.质量2.M25");
            AddViewItem("ribao", 42, 33, "外购焦炭质量.质量2.M10");
            AddViewItem("ribao", 42, 34, "外购焦炭质量.质量2.H2O");
            AddViewItem("ribao", 41, 35, "烟煤化验分析.质量1.C固");
            AddViewItem("ribao", 41, 36, "烟煤化验分析.质量1.A");
            AddViewItem("ribao", 41, 37, "烟煤化验分析.质量1.S");
            AddViewItem("ribao", 41, 38, "烟煤化验分析.质量1.V");
            AddViewItem("ribao", 42, 35, "烟煤化验分析.质量2.C固");
            AddViewItem("ribao", 42, 36, "烟煤化验分析.质量2.A");
            AddViewItem("ribao", 42, 37, "烟煤化验分析.质量2.S");
            AddViewItem("ribao", 42, 38, "烟煤化验分析.质量2.V");
            AddViewItem("ribao", 44, 28, "新焦化焦炭质量.质量1.C固");
            AddViewItem("ribao", 44, 29, "新焦化焦炭质量.质量1.A");
            AddViewItem("ribao", 44, 30, "新焦化焦炭质量.质量1.S");
            AddViewItem("ribao", 44, 31, "新焦化焦炭质量.质量1.V");
            AddViewItem("ribao", 44, 32, "新焦化焦炭质量.质量1.M25");
            AddViewItem("ribao", 44, 33, "新焦化焦炭质量.质量1.M10");
            AddViewItem("ribao", 44, 34, "新焦化焦炭质量.质量1.H2O");
            AddViewItem("ribao", 45, 28, "新焦化焦炭质量.质量2.C固");
            AddViewItem("ribao", 45, 29, "新焦化焦炭质量.质量2.A");
            AddViewItem("ribao", 45, 30, "新焦化焦炭质量.质量2.S");
            AddViewItem("ribao", 45, 31, "新焦化焦炭质量.质量2.V");
            AddViewItem("ribao", 45, 32, "新焦化焦炭质量.质量2.M25");
            AddViewItem("ribao", 45, 33, "新焦化焦炭质量.质量2.M10");
            AddViewItem("ribao", 45, 34, "新焦化焦炭质量.质量2.H2O");
            AddViewItem("ribao", 44, 35, "喷煤化验分析.质量1.C固");
            AddViewItem("ribao", 44, 36, "喷煤化验分析.质量1.A");
            AddViewItem("ribao", 44, 37, "喷煤化验分析.质量1.S");
            AddViewItem("ribao", 44, 38, "喷煤化验分析.质量1.V");
            AddViewItem("ribao", 45, 35, "喷煤化验分析.质量2.C固");
            AddViewItem("ribao", 45, 36, "喷煤化验分析.质量2.A");
            AddViewItem("ribao", 45, 37, "喷煤化验分析.质量2.S");
            AddViewItem("ribao", 45, 38, "喷煤化验分析.质量2.V");
            AddViewItem("ribao", 57, 29, "本日.烧结矿.230烧结.二车间");
            AddViewItem("ribao", 57, 30, "本日.烧结矿.230烧结.三车间");
            AddViewItem("ribao", 57, 31, "本日.烧结矿.230烧结.1号炉");
            AddViewItem("ribao", 57, 32, "本日.烧结矿.230烧结.合计");
            AddViewItem("ribao", 58, 29, "累计.烧结矿.230烧结.二车间");
            AddViewItem("ribao", 58, 30, "累计.烧结矿.230烧结.三车间");
            AddViewItem("ribao", 58, 31, "累计.烧结矿.230烧结.1号炉");
            AddViewItem("ribao", 58, 32, "累计.烧结矿.230烧结.合计");
            AddViewItem("ribao", 57, 34, "本日.烧结矿.大炼铁烧结.二车间");
            AddViewItem("ribao", 57, 35, "本日.烧结矿.大炼铁烧结.三车间");
            AddViewItem("ribao", 57, 36, "本日.烧结矿.大炼铁烧结.1号炉");
            AddViewItem("ribao", 57, 37, "本日.烧结矿.大炼铁烧结.合计");
            AddViewItem("ribao", 58, 34, "累计.烧结矿.大炼铁烧结.二车间");
            AddViewItem("ribao", 58, 35, "累计.烧结矿.大炼铁烧结.三车间");
            AddViewItem("ribao", 58, 36, "累计.烧结矿.大炼铁烧结.1号炉");
            AddViewItem("ribao", 58, 37, "累计.烧结矿.大炼铁烧结.合计");
            AddViewItem("ribao", 57, 38, "本日.烧结矿");
            AddViewItem("ribao", 58, 38, "累计.烧结矿");
            AddViewItem("ribao", 1, 31, "本日.4高炉.全铁产量");
            AddViewItem("ribao", 2, 32, "本日.5高炉.全铁产量");
            AddViewItem("ribao", 2, 31, "本日.5高炉.全铁产量");
            AddViewItem("ribao", 3, 31, "本日.6高炉.全铁产量");
            AddViewItem("ribao", 1, 34, "本日.7高炉.全铁产量");
            AddViewItem("ribao", 2, 34, "本日.8高炉.全铁产量");
            AddViewItem("ribao", 3, 34, "本日.1高炉.全铁产量");
            AddViewItem("ribao", 3, 37, "本日.全铁产量");
        }

        public void LoadFinish()
        {
            List<MemTag> lst = new List<MemTag>();
            foreach (var item in this)
            {
                item.NeedCal = true;
                if (item.GetViewPoint("ribao").Count > 0)
                    lst.Add(item);
            }
       
        }

        public void getDataBy(DateTime rq)
        {
            OracleConnection conn = new OracleConnection();
            conn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
            conn.Open();

            if (this.RQ != rq)
                LoadByRq(rq, conn);


            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;
            cmd.Parameters.Add(":rq", OracleType.DateTime).Value = rq;

            #region 全铁产量
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where  trunc(zdsj)=:rq and feliang is not null group by gaolu";
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetInitValue(string.Format("本日.{0}高炉.全铁产量", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();
            #endregion

            #region 炼钢铁,铸造铁

            cmd.CommandText = "select gaolu,case when fesi<=1.25 then '炼钢铁' else '铸造铁' end,sum(feliang) from ddluci where trunc(zdsj)=:rq and feliang is not null group by gaolu,case when fesi<=1.25 then '炼钢铁' else '铸造铁' end";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    if (!dr.IsDBNull(1))
                        this.SetInitValue(string.Format("本日.{0}高炉.{1}", gaolu, dr.GetString(1)), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                }
            }
            dr.Close();
            #endregion

            #region 合格铁
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where 合格铁(fesi,fes)=1 and trunc(zdsj)=:rq and feliang is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetInitValue(string.Format("本日.{0}高炉.合格铁", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();
            #endregion

            #region 一级铁
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where FES<=0.03 and trunc(zdsj)=:rq and feliang is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetInitValue(string.Format("本日.{0}高炉.一级铁", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();
            #endregion

            #region 去炼钢,去铸造
            cmd.CommandText = "select gaolu,decode(quchu,'炼钢','去炼钢','去铸造'),sum(feliang) from ddluci where trunc(zdsj)=:rq and feliang is not null group by gaolu,decode(quchu,'炼钢','去炼钢','去铸造')";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetInitValue(string.Format("本日.{0}高炉.{1}", gaolu, dr.GetString(1)), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                }
            }
            dr.Close();
            #endregion

            //#region 折算产量
            //cmd.CommandText = "select gaolu,sum(折算产量(feliang,fesi,fes)) from ddluci where trunc(zdsj)=:rq and dksj is not null group by gaolu";
            //dr = cmd.ExecuteReader();
            //while (dr.Read())
            //{
            //    if (!dr.IsDBNull(0))
            //    {
            //        int gaolu = dr.GetInt32(0);
            //        this.SetInitValue(string.Format("本日.{0}高炉.折算产量", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
            //    }
            //}
            //dr.Close();
            //#endregion

            #region 冷风
            cmd.CommandText = "select gaolu,sum(LFLL),sum(FYLL) from rbcaozuo where trunc(sj)=:rq group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetInitValue(string.Format("本日.{0}高炉.冷风", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                    this.SetInitValue(string.Format("本日.{0}高炉.氧气", gaolu), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                }
            }
            dr.Close();
            #endregion

            #region 大中修
            cmd.CommandText = "select 高炉,sum(间隔) from 休风 where  trunc(时间)=:rq and (分类='大修' or 分类='中修')  and  间隔>0　and 间隔 is not null group by 高炉";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                int gaolu = dr.IsDBNull(0) ? 6 : dr.GetInt32(0);
                this.SetInitValue(string.Format("本日.{0}高炉.大中修", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));

            }
            dr.Close();
            #endregion\

            #region 生铁含硅
            cmd.CommandText = "select nvl(gaolu,9999),round(avg(FESI),2),round(avg(FEMN),2),round(avg(FEP),3),round(avg(FES),3) from ddluci where trunc(zdsj)=:rq and dksj is not null group by rollup(gaolu)";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    if (gaolu != 9999)
                    {
                        this.SetInitValue(string.Format("本日.{0}高炉.生铁含硅", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                        this.SetInitValue(string.Format("本日.{0}高炉.生铁含锰", gaolu), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                        this.SetInitValue(string.Format("本日.{0}高炉.生铁含磷", gaolu), dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                        this.SetInitValue(string.Format("本日.{0}高炉.生铁含硫", gaolu), dr.IsDBNull(4) ? 0 : dr.GetDouble(4));
                    }
                    else
                    {
                        this.SetInitValue("本日.生铁含硅", dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                        this.SetInitValue("本日.生铁含锰", dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                        this.SetInitValue("本日.生铁含磷", dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                        this.SetInitValue("本日.生铁含硫", dr.IsDBNull(4) ? 0 : dr.GetDouble(4));
                    }
                }
            }
            dr.Close();
            #endregion

            #region 炉渣碱度
            cmd.CommandText = "select  nvl(gaolu,9999),round(avg(ZHAR2),2) from ddluci where trunc(zdsj)=:rq and dksj is not null group by rollup(gaolu)";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    if (gaolu != 9999)
                    {
                        this.SetInitValue(string.Format("本日.{0}高炉.炉渣碱度", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                    }
                    else
                    {
                        this.SetInitValue(string.Format("本日.炉渣碱度"), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                    }
                }
            }
            dr.Close();
            #endregion

            #region 平均风温
            cmd.CommandText = "select  nvl(gaolu,9999),round(avg(RFWD),0) from rbcaozuo where trunc(sj)=:rq and RFWD is not null  group by rollup(gaolu)";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    if (gaolu != 9999)
                    {
                        this.SetInitValue(string.Format("本日.{0}高炉.平均风温", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                    }
                    else
                    {
                        this.SetInitValue(string.Format("本日.平均风温"), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                    }
                }
            }
            dr.Close();
            #endregion

            #region  消耗
            cmd.CommandText = "SELECT  GAOLU,MC,round(SUM(NVL(BAIBAN,0)+NVL(ZHONGBAN,0)+NVL(YEBAN,0))/1000,2) FROM RBXIAOHAO WHERE TRUNC(SJ)=:rq group by gaolu,mc";

            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                int gaolu = dr.GetInt32(0);
                if (!dr.IsDBNull(1))
                    this.SetInitValue(string.Format("本日.{0}高炉.{1}", gaolu, dr.GetString(1)), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
            }
            dr.Close();
            #endregion

            #region 休风
            Dictionary<string, string> map = new Dictionary<string, string>();
            //            map.Add("", "休风计划");
            //            map.Add("", "休风工艺");
            //            map.Add("", "休风电");
            //            map.Add("设备影响", "休风设备");
            //            map.Add("设备影响", "休风设备");
            //            map.Add("", "休风外围");
            map.Add("处理进风件", "休风设备");
            map.Add("大修", "休风计划");
            map.Add("计划检修", "休风计划");
            map.Add("工艺操作", "休风工艺");
            map.Add("其它设备故障", "休风设备");
            map.Add("设备影响", "休风设备");
            map.Add("坐料", "休风工艺");
            map.Add("工艺休风", "休风工艺");
            map.Add("其它", "休风计划");
            map.Add("易损件", "休风设备");

            cmd.CommandText = "select 高炉,分类,sum(间隔) from 休风 where trunc(时间)=:rq group by 高炉,分类";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    string fenlei = dr.IsDBNull(1) ? "" : dr.GetString(1);
                    double shijian = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                    if (map.ContainsKey(fenlei))
                        this.SetInitValue(string.Format("本日.{0}高炉.{1}", gaolu, map[fenlei]), shijian);
                    else
                        this.SetInitValue(string.Format("本日.{0}高炉.休风计划", gaolu), shijian);
                }
            }
            dr.Close();

            cmd.CommandText = "select 高炉,sum(间隔) from 休风 where trunc(时间)=:rq group by 高炉";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetInitValue(string.Format("本日.{0}高炉.休风时间", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();
            #endregion

            #region TRT发电
            cmd.CommandText = "select 高炉,sum(间隔) from 慢风 where trunc(时间)=:rq group by 高炉";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetInitValue(string.Format("本日.{0}高炉.慢风", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();
            #endregion

            #region 昨日累计部分

            if (rq.Day != 1)
            {
                OracleCommand yesterdayCmd = new OracleCommand();
                yesterdayCmd.Connection = conn;
                yesterdayCmd.CommandText = "select t.tagname,t.val from rd_ri t where t.rq=:RQ and t.tagname like '累计.%'";
                yesterdayCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = rq.AddDays(-1);
                OracleDataReader yesterdayDr = yesterdayCmd.ExecuteReader();
                while (yesterdayDr.Read())
                {
                    if (!yesterdayDr.IsDBNull(0) && !yesterdayDr.IsDBNull(1))
                    {
                        this.SetInitValue(yesterdayDr.GetString(0).Replace("累计.", "昨日累计."), yesterdayDr.GetDouble(1));
                    }
                }
                yesterdayDr.Close();
            }

            this.SetInitValue("累计.累计天数", rq.Day);

            #endregion

            #region 生铁含硅

            OracleCommand statisticCmd = new OracleCommand();
            statisticCmd.Connection = conn;
            statisticCmd.CommandText = "select nvl(gaolu,9999),round(avg(FESI),2),round(avg(FEMN),2),round(avg(FEP),3),round(avg(FES),3)  from ddluci where zdsj>=:rq1 and zdsj<:rq2 and dksj is not null group by rollup(gaolu)";
            statisticCmd.Parameters.Add(":rq1", OracleType.DateTime).Value = new DateTime(rq.Year, rq.Month, 1);
            statisticCmd.Parameters.Add(":rq2", OracleType.DateTime).Value = rq.AddDays(1);
            OracleDataReader statisticDr = statisticCmd.ExecuteReader();
            while (statisticDr.Read())
            {
                if (!statisticDr.IsDBNull(0))
                {
                    int gaolu = statisticDr.GetInt32(0);
                    if (gaolu != 9999)
                    {
                        this.SetInitValue(string.Format("累计.{0}高炉.生铁含硅", gaolu), statisticDr.IsDBNull(1) ? 0 : statisticDr.GetDouble(1));
                        this.SetInitValue(string.Format("累计.{0}高炉.生铁含锰", gaolu), statisticDr.IsDBNull(2) ? 0 : statisticDr.GetDouble(2));
                        this.SetInitValue(string.Format("累计.{0}高炉.生铁含磷", gaolu), statisticDr.IsDBNull(3) ? 0 : statisticDr.GetDouble(3));
                        this.SetInitValue(string.Format("累计.{0}高炉.生铁含硫", gaolu), statisticDr.IsDBNull(4) ? 0 : statisticDr.GetDouble(4));

                    }
                    else
                    {
                        this.SetInitValue("累计.生铁含硅", statisticDr.IsDBNull(1) ? 0 : statisticDr.GetDouble(1));
                        this.SetInitValue("累计.生铁含锰", statisticDr.IsDBNull(2) ? 0 : statisticDr.GetDouble(2));
                        this.SetInitValue("累计.生铁含磷", statisticDr.IsDBNull(3) ? 0 : statisticDr.GetDouble(3));
                        this.SetInitValue("累计.生铁含硫", statisticDr.IsDBNull(4) ? 0 : statisticDr.GetDouble(4));
                    }
                }
            }
            statisticDr.Close();
            #endregion

            #region 炉渣碱度
            statisticCmd.CommandText = "select  nvl(gaolu,9999),round(avg(ZHAR2),2) from ddluci where zdsj>=:rq1 and zdsj<:rq2 and dksj is not null group by rollup(gaolu)";
            statisticDr = statisticCmd.ExecuteReader();
            while (statisticDr.Read())
            {
                if (!statisticDr.IsDBNull(0))
                {
                    int gaolu = statisticDr.GetInt32(0);
                    if (gaolu != 9999)
                    {
                        this.SetInitValue(string.Format("累计.{0}高炉.炉渣碱度", gaolu), statisticDr.IsDBNull(1) ? 0 : statisticDr.GetDouble(1));
                    }
                    else
                    {
                        this.SetInitValue(string.Format("累计.炉渣碱度"), statisticDr.IsDBNull(1) ? 0 : statisticDr.GetDouble(1));
                    }
                }
            }
            dr.Close();
            #endregion

            #region 平均风温
            statisticCmd.CommandText = "select  nvl(gaolu,9999),round(avg(RFWD),0) from rbcaozuo where sj>=:rq1 and sj<:rq2  and RFWD is not null  group by rollup(gaolu)";
            statisticDr = statisticCmd.ExecuteReader();
            while (statisticDr.Read())
            {
                if (!statisticDr.IsDBNull(0))
                {
                    int gaolu = statisticDr.GetInt32(0);
                    if (gaolu != 9999)
                    {
                        this.SetInitValue(string.Format("累计.{0}高炉.平均风温", gaolu), statisticDr.IsDBNull(1) ? 0 : statisticDr.GetDouble(1));
                    }
                    else
                    {
                        this.SetInitValue(string.Format("累计.平均风温"), statisticDr.IsDBNull(1) ? 0 : statisticDr.GetDouble(1));
                    }
                }
            }
            statisticDr.Close();
            #endregion

            #region 原料成分

            //Dictionary<string, List<double>> shuju = new Dictionary<string, List<double>>();

            //shuju.Add("入炉矿化验分析.78号炉1.FeO", new List<double>());
            //shuju.Add("入炉矿化验分析.78号炉1.R", new List<double>());
            //shuju.Add("入炉矿化验分析.78号炉1.S", new List<double>());
            //shuju.Add("入炉矿化验分析.78号炉1.TFe", new List<double>());
            //shuju.Add("入炉矿化验分析.78号炉1.块矿品位", new List<double>());
            //shuju.Add("入炉矿化验分析.78号炉2.FeO", new List<double>());
            //shuju.Add("入炉矿化验分析.78号炉2.R", new List<double>());
            //shuju.Add("入炉矿化验分析.78号炉2.S", new List<double>());
            //shuju.Add("入炉矿化验分析.78号炉2.TFe", new List<double>());
            //shuju.Add("入炉矿化验分析.78号炉2.块矿品位", new List<double>());
            //shuju.Add("入炉矿化验分析.二车间1.FeO", new List<double>());
            //shuju.Add("入炉矿化验分析.二车间1.R", new List<double>());
            //shuju.Add("入炉矿化验分析.二车间1.S", new List<double>());
            //shuju.Add("入炉矿化验分析.二车间1.TFe", new List<double>());
            //shuju.Add("入炉矿化验分析.二车间1.块矿品位", new List<double>());
            //shuju.Add("入炉矿化验分析.二车间2.FeO", new List<double>());
            //shuju.Add("入炉矿化验分析.二车间2.R", new List<double>());
            //shuju.Add("入炉矿化验分析.二车间2.S", new List<double>());
            //shuju.Add("入炉矿化验分析.二车间2.TFe", new List<double>());
            //shuju.Add("入炉矿化验分析.二车间2.块矿品位", new List<double>());
            //shuju.Add("入炉矿化验分析.全厂合计1.FeO", new List<double>());
            //shuju.Add("入炉矿化验分析.全厂合计1.R", new List<double>());
            //shuju.Add("入炉矿化验分析.全厂合计1.S", new List<double>());
            //shuju.Add("入炉矿化验分析.全厂合计1.TFe", new List<double>());
            //shuju.Add("入炉矿化验分析.全厂合计1.块矿品位", new List<double>());
            //shuju.Add("入炉矿化验分析.全厂合计2.FeO", new List<double>());
            //shuju.Add("入炉矿化验分析.全厂合计2.R", new List<double>());
            //shuju.Add("入炉矿化验分析.全厂合计2.S", new List<double>());
            //shuju.Add("入炉矿化验分析.全厂合计2.TFe", new List<double>());
            //shuju.Add("入炉矿化验分析.全厂合计2.块矿品位", new List<double>());
            //shuju.Add("入炉矿化验分析.1高炉1.FeO", new List<double>());
            //shuju.Add("入炉矿化验分析.1高炉1.R", new List<double>());
            //shuju.Add("入炉矿化验分析.1高炉1.S", new List<double>());
            //shuju.Add("入炉矿化验分析.1高炉1.TFe", new List<double>());
            //shuju.Add("入炉矿化验分析.1高炉1.块矿品位", new List<double>());
            //shuju.Add("入炉矿化验分析.1高炉2.FeO", new List<double>());
            //shuju.Add("入炉矿化验分析.1高炉2.R", new List<double>());
            //shuju.Add("入炉矿化验分析.1高炉2.S", new List<double>());
            //shuju.Add("入炉矿化验分析.1高炉2.TFe", new List<double>());
            //shuju.Add("入炉矿化验分析.1高炉2.块矿品位", new List<double>());

            //shuju.Add("焦炭入炉分析.78号炉1.A", new List<double>());
            //shuju.Add("焦炭入炉分析.78号炉1.C固", new List<double>());
            //shuju.Add("焦炭入炉分析.78号炉1.H2O", new List<double>());
            //shuju.Add("焦炭入炉分析.78号炉1.S", new List<double>());
            //shuju.Add("焦炭入炉分析.78号炉1.V", new List<double>());
            //shuju.Add("焦炭入炉分析.78号炉2.A", new List<double>());
            //shuju.Add("焦炭入炉分析.78号炉2.C固", new List<double>());
            //shuju.Add("焦炭入炉分析.78号炉2.H2O", new List<double>());
            //shuju.Add("焦炭入炉分析.78号炉2.S", new List<double>());
            //shuju.Add("焦炭入炉分析.78号炉2.V", new List<double>());
            //shuju.Add("焦炭入炉分析.二车间1.A", new List<double>());
            //shuju.Add("焦炭入炉分析.二车间1.C固", new List<double>());
            //shuju.Add("焦炭入炉分析.二车间1.H2O", new List<double>());
            //shuju.Add("焦炭入炉分析.二车间1.S", new List<double>());
            //shuju.Add("焦炭入炉分析.二车间1.V", new List<double>());
            //shuju.Add("焦炭入炉分析.二车间2.A", new List<double>());
            //shuju.Add("焦炭入炉分析.二车间2.C固", new List<double>());
            //shuju.Add("焦炭入炉分析.二车间2.H2O", new List<double>());
            //shuju.Add("焦炭入炉分析.二车间2.S", new List<double>());
            //shuju.Add("焦炭入炉分析.二车间2.V", new List<double>());
            //shuju.Add("焦炭入炉分析.全厂合计1.A", new List<double>());
            //shuju.Add("焦炭入炉分析.全厂合计1.C固", new List<double>());
            //shuju.Add("焦炭入炉分析.全厂合计1.H2O", new List<double>());
            //shuju.Add("焦炭入炉分析.全厂合计1.S", new List<double>());
            //shuju.Add("焦炭入炉分析.全厂合计1.V", new List<double>());
            //shuju.Add("焦炭入炉分析.全厂合计2.A", new List<double>());
            //shuju.Add("焦炭入炉分析.全厂合计2.C固", new List<double>());
            //shuju.Add("焦炭入炉分析.全厂合计2.H2O", new List<double>());
            //shuju.Add("焦炭入炉分析.全厂合计2.S", new List<double>());
            //shuju.Add("焦炭入炉分析.全厂合计2.V", new List<double>());
            //shuju.Add("焦炭入炉分析.1高炉1.A", new List<double>());
            //shuju.Add("焦炭入炉分析.1高炉1.C固", new List<double>());
            //shuju.Add("焦炭入炉分析.1高炉1.H2O", new List<double>());
            //shuju.Add("焦炭入炉分析.1高炉1.S", new List<double>());
            //shuju.Add("焦炭入炉分析.1高炉1.V", new List<double>());
            //shuju.Add("焦炭入炉分析.1高炉2.A", new List<double>());
            //shuju.Add("焦炭入炉分析.1高炉2.C固", new List<double>());
            //shuju.Add("焦炭入炉分析.1高炉2.H2O", new List<double>());
            //shuju.Add("焦炭入炉分析.1高炉2.S", new List<double>());
            //shuju.Add("焦炭入炉分析.1高炉2.V", new List<double>());


            Shuju shuju = new Shuju(this);
            cmd.CommandText = "select MC,CANG,round(avg(TFE),2) as TFE, round(avg(FEO),2) as FEO, round(avg(S),3) as S,round(avg(JD),3) as JD from ddyl where trunc(sj)=:rq   group by MC,CANG";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                string mc = dr.IsDBNull(0) ? "" : dr.GetString(0);
                string cang = dr.IsDBNull(1) ? "" : dr.GetString(1);
                double tfe = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                double feo = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                double s = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                double jd = dr.IsDBNull(5) ? 0 : dr.GetDouble(5);
                if (mc == "烧结矿" || mc == "球团矿")
                {
                    if (cang.StartsWith("4#高炉") || cang.StartsWith("5#高炉") || cang.StartsWith("6#高炉") || cang.StartsWith("7#高炉") || cang.StartsWith("8#高炉") || cang.StartsWith("1#高炉"))
                    {
                        shuju.AddToDict("入炉矿化验分析.全厂合计1.FeO", feo);
                        shuju.AddToDict("入炉矿化验分析.全厂合计1.R", jd);
                        shuju.AddToDict("入炉矿化验分析.全厂合计1.S", s);
                        shuju.AddToDict("入炉矿化验分析.全厂合计1.TFe", tfe);
                    }

                    if (cang.StartsWith("4#高炉") || cang.StartsWith("5#高炉") || cang.StartsWith("6#高炉"))
                    {
                        shuju.AddToDict("入炉矿化验分析.二车间1.FeO", feo);
                        shuju.AddToDict("入炉矿化验分析.二车间1.R", jd);
                        shuju.AddToDict("入炉矿化验分析.二车间1.S", s);
                        shuju.AddToDict("入炉矿化验分析.二车间1.TFe", tfe);


                    }
                    else if (cang.StartsWith("7#高炉") || cang.StartsWith("8#高炉"))
                    {
                        shuju.AddToDict("入炉矿化验分析.78号炉1.FeO", feo);
                        shuju.AddToDict("入炉矿化验分析.78号炉1.R", jd);
                        shuju.AddToDict("入炉矿化验分析.78号炉1.S", s);
                        shuju.AddToDict("入炉矿化验分析.78号炉1.TFe", tfe);
                    } if (cang.StartsWith("1#高炉"))
                    {
                        shuju.AddToDict("入炉矿化验分析.1高炉1.FeO", feo);
                        shuju.AddToDict("入炉矿化验分析.1高炉1.R", jd);
                        shuju.AddToDict("入炉矿化验分析.1高炉1.S", s);
                        shuju.AddToDict("入炉矿化验分析.1高炉1.TFe", tfe);
                    }
                }
                else if (mc == "块矿")
                {
                    shuju.AddToDict("入炉矿化验分析.全厂合计1.块矿品位", tfe);
                    if (cang.StartsWith("4#高炉") || cang.StartsWith("5#高炉") || cang.StartsWith("6#高炉"))
                    {
                        shuju.AddToDict("入炉矿化验分析.二车间1.块矿品位", tfe);
                    }
                    else if (cang.StartsWith("7#高炉") || cang.StartsWith("8#高炉"))
                    {
                        shuju.AddToDict("入炉矿化验分析.78号炉1.块矿品位", tfe);
                    } if (cang.StartsWith("1#高炉"))
                    {
                        shuju.AddToDict("入炉矿化验分析.1高炉1.块矿品位", tfe);
                    }
                }
            }
            dr.Close();
            shuju.ExportDict();


            Shuju yllj = new Shuju(this);
            statisticCmd.CommandText = "select MC,CANG,round(avg(TFE),2) as TFE, round(avg(FEO),2) as FEO, round(avg(S),3) as S,round(avg(JD),3) as JD from ddyl where sj>=:rq1 and sj<:rq2   group by MC,CANG";
            dr = statisticCmd.ExecuteReader();
            while (dr.Read())
            {
                string mc = dr.IsDBNull(0) ? "" : dr.GetString(0);
                string cang = dr.IsDBNull(1) ? "" : dr.GetString(1);
                double tfe = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                double feo = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                double s = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                double jd = dr.IsDBNull(5) ? 0 : dr.GetDouble(5);
                if (mc == "烧结矿" || mc == "球团矿")
                {
                    if (cang.StartsWith("4#高炉") || cang.StartsWith("5#高炉") || cang.StartsWith("6#高炉") || cang.StartsWith("7#高炉") || cang.StartsWith("8#高炉") || cang.StartsWith("1#高炉"))
                    {
                        yllj.AddToDict("入炉矿化验分析.全厂合计2.FeO", feo);
                        yllj.AddToDict("入炉矿化验分析.全厂合计2.R", jd);
                        yllj.AddToDict("入炉矿化验分析.全厂合计2.S", s);
                        yllj.AddToDict("入炉矿化验分析.全厂合计2.TFe", tfe);
                    }

                    if (cang.StartsWith("4#高炉") || cang.StartsWith("5#高炉") || cang.StartsWith("6#高炉"))
                    {
                        yllj.AddToDict("入炉矿化验分析.二车间2.FeO", feo);
                        yllj.AddToDict("入炉矿化验分析.二车间2.R", jd);
                        yllj.AddToDict("入炉矿化验分析.二车间2.S", s);
                        yllj.AddToDict("入炉矿化验分析.二车间2.TFe", tfe);
                    }
                    else if (cang.StartsWith("7#高炉") || cang.StartsWith("8#高炉"))
                    {
                        yllj.AddToDict("入炉矿化验分析.78号炉2.FeO", feo);
                        yllj.AddToDict("入炉矿化验分析.78号炉2.R", jd);
                        yllj.AddToDict("入炉矿化验分析.78号炉2.S", s);
                        yllj.AddToDict("入炉矿化验分析.78号炉2.TFe", tfe);
                    } if (cang.StartsWith("1#高炉"))
                    {
                        yllj.AddToDict("入炉矿化验分析.1高炉2.FeO", feo);
                        yllj.AddToDict("入炉矿化验分析.1高炉2.R", jd);
                        yllj.AddToDict("入炉矿化验分析.1高炉2.S", s);
                        yllj.AddToDict("入炉矿化验分析.1高炉2.TFe", tfe);
                    }
                }
                else if (mc == "块矿")
                {
                    yllj.AddToDict("入炉矿化验分析.全厂合计2.块矿品位", tfe);
                    if (cang.StartsWith("4#高炉") || cang.StartsWith("5#高炉") || cang.StartsWith("6#高炉"))
                    {
                        yllj.AddToDict("入炉矿化验分析.二车间2.块矿品位", tfe);
                    }
                    else if (cang.StartsWith("7#高炉") || cang.StartsWith("8#高炉"))
                    {
                        yllj.AddToDict("入炉矿化验分析.78号炉2.块矿品位", tfe);
                    } if (cang.StartsWith("1#高炉"))
                    {
                        yllj.AddToDict("入炉矿化验分析.1高炉2.块矿品位", tfe);
                    }
                }
            }
            dr.Close();
            yllj.ExportDict();
            #endregion
            //

            #region 焦炭质量
            Shuju jtbr = new Shuju(this);
            cmd.CommandText = "select  substr(to_single_byte(trim(mc)),1,1) as gl,round(avg(c),2) as c,round(avg(shuif),2) as shuif,round(avg(huifen),2) as huifen,round(avg(huifa),2) as huifa,round(avg(s),2) as s,round(avg(m25),2) as m25,round(avg(m10),2) as m10,round(avg(qd),2) as qd from ddjt where trunc(sj)=:rq   group by substr(to_single_byte(trim(mc)),1,1)";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                string mc = dr.IsDBNull(0) ? "" : dr.GetString(0);
                if (mc == "1")
                {
                    if (!dr.IsDBNull(1))
                    {
                        jtbr.AddToDict("焦炭入炉分析.1高炉1.C固", dr.GetDouble(1));
                        jtbr.AddToDict("焦炭入炉分析.全厂合计1.C固", dr.GetDouble(1));
                    }
                    if (!dr.IsDBNull(2))
                    {
                        jtbr.AddToDict("焦炭入炉分析.1高炉1.H2O", dr.GetDouble(2));
                        jtbr.AddToDict("焦炭入炉分析.全厂合计1.H2O", dr.GetDouble(2));
                    }
                    if (!dr.IsDBNull(3))
                    {
                        jtbr.AddToDict("焦炭入炉分析.1高炉1.A", dr.GetDouble(3));
                        jtbr.AddToDict("焦炭入炉分析.全厂合计1.A", dr.GetDouble(3));
                    }
                    if (!dr.IsDBNull(4))
                    {
                        jtbr.AddToDict("焦炭入炉分析.1高炉1.V", dr.GetDouble(4));
                        jtbr.AddToDict("焦炭入炉分析.全厂合计1.V", dr.GetDouble(4));
                    }
                    if (!dr.IsDBNull(5))
                    {
                        jtbr.AddToDict("焦炭入炉分析.1高炉1.S", dr.GetDouble(5));
                        jtbr.AddToDict("焦炭入炉分析.全厂合计1.S", dr.GetDouble(5));
                    }

                }
                else if (mc == "4" || mc == "5" || mc == "6")
                {
                    if (!dr.IsDBNull(1))
                    {
                        jtbr.AddToDict("焦炭入炉分析.二车间1.C固", dr.GetDouble(1));
                        jtbr.AddToDict("焦炭入炉分析.全厂合计1.C固", dr.GetDouble(1));
                    }
                    if (!dr.IsDBNull(2))
                    {
                        jtbr.AddToDict("焦炭入炉分析.二车间1.H2O", dr.GetDouble(2));
                        jtbr.AddToDict("焦炭入炉分析.全厂合计1.H2O", dr.GetDouble(2));
                    }
                    if (!dr.IsDBNull(3))
                    {
                        jtbr.AddToDict("焦炭入炉分析.二车间1.A", dr.GetDouble(3));
                        jtbr.AddToDict("焦炭入炉分析.全厂合计1.A", dr.GetDouble(3));
                    }
                    if (!dr.IsDBNull(4))
                    {
                        jtbr.AddToDict("焦炭入炉分析.二车间1.V", dr.GetDouble(4));
                        jtbr.AddToDict("焦炭入炉分析.全厂合计1.V", dr.GetDouble(4));
                    }
                    if (!dr.IsDBNull(5))
                    {
                        jtbr.AddToDict("焦炭入炉分析.二车间1.S", dr.GetDouble(5));
                        jtbr.AddToDict("焦炭入炉分析.全厂合计1.S", dr.GetDouble(5));
                    }
                }
                else if (mc == "7" || mc == "8")
                {
                    if (!dr.IsDBNull(1))
                    {
                        jtbr.AddToDict("焦炭入炉分析.78号炉1.C固", dr.GetDouble(1));
                        jtbr.AddToDict("焦炭入炉分析.全厂合计1.C固", dr.GetDouble(1));
                    }
                    if (!dr.IsDBNull(2))
                    {
                        jtbr.AddToDict("焦炭入炉分析.78号炉1.H2O", dr.GetDouble(2));
                        jtbr.AddToDict("焦炭入炉分析.全厂合计1.H2O", dr.GetDouble(2));
                    }
                    if (!dr.IsDBNull(3))
                    {
                        jtbr.AddToDict("焦炭入炉分析.78号炉1.A", dr.GetDouble(3));
                        jtbr.AddToDict("焦炭入炉分析.全厂合计1.A", dr.GetDouble(3));
                    }
                    if (!dr.IsDBNull(4))
                    {
                        jtbr.AddToDict("焦炭入炉分析.78号炉1.V", dr.GetDouble(4));
                        jtbr.AddToDict("焦炭入炉分析.全厂合计1.V", dr.GetDouble(4));
                    }
                    if (!dr.IsDBNull(5))
                    {
                        jtbr.AddToDict("焦炭入炉分析.78号炉1.S", dr.GetDouble(5));
                        jtbr.AddToDict("焦炭入炉分析.全厂合计1.S", dr.GetDouble(5));
                    }
                }
                else if (mc == "新")
                {
                    if (!dr.IsDBNull(1)) jtbr.AddToDict("新焦化焦炭质量.质量1.C固", dr.GetDouble(1));
                    if (!dr.IsDBNull(2)) jtbr.AddToDict("新焦化焦炭质量.质量1.H2O", dr.GetDouble(2));
                    if (!dr.IsDBNull(3)) jtbr.AddToDict("新焦化焦炭质量.质量1.A", dr.GetDouble(3));
                    if (!dr.IsDBNull(4)) jtbr.AddToDict("新焦化焦炭质量.质量1.V", dr.GetDouble(4));
                    if (!dr.IsDBNull(5)) jtbr.AddToDict("新焦化焦炭质量.质量1.S", dr.GetDouble(5));
                    if (!dr.IsDBNull(6)) jtbr.AddToDict("新焦化焦炭质量.质量1.M25", dr.GetDouble(6));
                    if (!dr.IsDBNull(7)) jtbr.AddToDict("新焦化焦炭质量.质量1.M10", dr.GetDouble(7));
                }
                else if (mc == "外")
                {
                    if (!dr.IsDBNull(1)) jtbr.AddToDict("外购焦炭质量.质量1.C固", dr.GetDouble(1));
                    if (!dr.IsDBNull(2)) jtbr.AddToDict("外购焦炭质量.质量1.H2O", dr.GetDouble(2));
                    if (!dr.IsDBNull(3)) jtbr.AddToDict("外购焦炭质量.质量1.A", dr.GetDouble(3));
                    if (!dr.IsDBNull(4)) jtbr.AddToDict("外购焦炭质量.质量1.V", dr.GetDouble(4));
                    if (!dr.IsDBNull(5)) jtbr.AddToDict("外购焦炭质量.质量1.S", dr.GetDouble(5));
                    if (!dr.IsDBNull(6)) jtbr.AddToDict("外购焦炭质量.质量1.M25", dr.GetDouble(6));
                    if (!dr.IsDBNull(7)) jtbr.AddToDict("外购焦炭质量.质量1.M10", dr.GetDouble(7));
                }

            }
            dr.Close();
            jtbr.ExportDict();

            Shuju jtlj = new Shuju(this);
            statisticCmd.CommandText = "select  substr(to_single_byte(trim(mc)),1,1) as gl,round(avg(c),2) as c,round(avg(shuif),2) as shuif,round(avg(huifen),2) as huifen,round(avg(huifa),2) as huifa,round(avg(s),2) as s,round(avg(m25),2) as m25,round(avg(m10),2) as m10,round(avg(qd),2) as qd from ddjt where sj>=:rq1 and sj<:rq2   group by substr(to_single_byte(trim(mc)),1,1)";
            dr = statisticCmd.ExecuteReader();
            while (dr.Read())
            {
                string mc = dr.IsDBNull(0) ? "" : dr.GetString(0);
                if (mc == "1")
                {
                    if (!dr.IsDBNull(1))
                    {
                        jtlj.AddToDict("焦炭入炉分析.1高炉2.C固", dr.GetDouble(1));
                        jtlj.AddToDict("焦炭入炉分析.全厂合计2.C固", dr.GetDouble(1));
                    }
                    if (!dr.IsDBNull(2))
                    {
                        jtlj.AddToDict("焦炭入炉分析.1高炉2.H2O", dr.GetDouble(2));
                        jtlj.AddToDict("焦炭入炉分析.全厂合计2.H2O", dr.GetDouble(2));
                    }
                    if (!dr.IsDBNull(3))
                    {
                        jtlj.AddToDict("焦炭入炉分析.1高炉2.A", dr.GetDouble(3));
                        jtlj.AddToDict("焦炭入炉分析.全厂合计2.A", dr.GetDouble(3));
                    }
                    if (!dr.IsDBNull(4))
                    {
                        jtlj.AddToDict("焦炭入炉分析.1高炉2.V", dr.GetDouble(4));
                        jtlj.AddToDict("焦炭入炉分析.全厂合计2.V", dr.GetDouble(4));
                    }
                    if (!dr.IsDBNull(5))
                    {
                        jtlj.AddToDict("焦炭入炉分析.1高炉2.S", dr.GetDouble(5));
                        jtlj.AddToDict("焦炭入炉分析.全厂合计2.S", dr.GetDouble(5));
                    }

                }
                else if (mc == "4" || mc == "5" || mc == "6")
                {
                    if (!dr.IsDBNull(1))
                    {
                        jtlj.AddToDict("焦炭入炉分析.二车间2.C固", dr.GetDouble(1));
                        jtlj.AddToDict("焦炭入炉分析.全厂合计2.C固", dr.GetDouble(1));
                    }
                    if (!dr.IsDBNull(2))
                    {
                        jtlj.AddToDict("焦炭入炉分析.二车间2.H2O", dr.GetDouble(2));
                        jtlj.AddToDict("焦炭入炉分析.全厂合计2.H2O", dr.GetDouble(2));
                    }
                    if (!dr.IsDBNull(3))
                    {
                        jtlj.AddToDict("焦炭入炉分析.二车间2.A", dr.GetDouble(3));
                        jtlj.AddToDict("焦炭入炉分析.全厂合计2.A", dr.GetDouble(3));
                    }
                    if (!dr.IsDBNull(4))
                    {
                        jtlj.AddToDict("焦炭入炉分析.二车间2.V", dr.GetDouble(4));
                        jtlj.AddToDict("焦炭入炉分析.全厂合计2.V", dr.GetDouble(4));
                    }
                    if (!dr.IsDBNull(5))
                    {
                        jtlj.AddToDict("焦炭入炉分析.二车间2.S", dr.GetDouble(5));
                        jtlj.AddToDict("焦炭入炉分析.全厂合计2.S", dr.GetDouble(5));
                    }
                }
                else if (mc == "7" || mc == "8")
                {
                    if (!dr.IsDBNull(1))
                    {
                        jtlj.AddToDict("焦炭入炉分析.78号炉2.C固", dr.GetDouble(1));
                        jtlj.AddToDict("焦炭入炉分析.全厂合计2.C固", dr.GetDouble(1));
                    }
                    if (!dr.IsDBNull(2))
                    {
                        jtlj.AddToDict("焦炭入炉分析.78号炉2.H2O", dr.GetDouble(2));
                        jtlj.AddToDict("焦炭入炉分析.全厂合计2.H2O", dr.GetDouble(2));
                    }
                    if (!dr.IsDBNull(3))
                    {
                        jtlj.AddToDict("焦炭入炉分析.78号炉2.A", dr.GetDouble(3));
                        jtlj.AddToDict("焦炭入炉分析.全厂合计2.A", dr.GetDouble(3));
                    }
                    if (!dr.IsDBNull(4))
                    {
                        jtlj.AddToDict("焦炭入炉分析.78号炉2.V", dr.GetDouble(4));
                        jtlj.AddToDict("焦炭入炉分析.全厂合计2.V", dr.GetDouble(4));
                    }
                    if (!dr.IsDBNull(5))
                    {
                        jtlj.AddToDict("焦炭入炉分析.78号炉2.S", dr.GetDouble(5));
                        jtlj.AddToDict("焦炭入炉分析.全厂合计2.S", dr.GetDouble(5));
                    }
                }
                else if (mc == "新")
                {
                    if (!dr.IsDBNull(1)) jtlj.AddToDict("新焦化焦炭质量.质量2.C固", dr.GetDouble(1));
                    if (!dr.IsDBNull(2)) jtlj.AddToDict("新焦化焦炭质量.质量2.H2O", dr.GetDouble(2));
                    if (!dr.IsDBNull(3)) jtlj.AddToDict("新焦化焦炭质量.质量2.A", dr.GetDouble(3));
                    if (!dr.IsDBNull(4)) jtlj.AddToDict("新焦化焦炭质量.质量2.V", dr.GetDouble(4));
                    if (!dr.IsDBNull(5)) jtlj.AddToDict("新焦化焦炭质量.质量2.S", dr.GetDouble(5));
                    if (!dr.IsDBNull(6)) jtlj.AddToDict("新焦化焦炭质量.质量2.M25", dr.GetDouble(6));
                    if (!dr.IsDBNull(7)) jtlj.AddToDict("新焦化焦炭质量.质量2.M10", dr.GetDouble(7));
                }
                else if (mc == "外")
                {
                    if (!dr.IsDBNull(1)) jtlj.AddToDict("外购焦炭质量.质量2.C固", dr.GetDouble(1));
                    if (!dr.IsDBNull(2)) jtlj.AddToDict("外购焦炭质量.质量2.H2O", dr.GetDouble(2));
                    if (!dr.IsDBNull(3)) jtlj.AddToDict("外购焦炭质量.质量2.A", dr.GetDouble(3));
                    if (!dr.IsDBNull(4)) jtlj.AddToDict("外购焦炭质量.质量2.V", dr.GetDouble(4));
                    if (!dr.IsDBNull(5)) jtlj.AddToDict("外购焦炭质量.质量2.S", dr.GetDouble(5));
                    if (!dr.IsDBNull(6)) jtlj.AddToDict("外购焦炭质量.质量2.M25", dr.GetDouble(6));
                    if (!dr.IsDBNull(7)) jtlj.AddToDict("外购焦炭质量.质量2.M10", dr.GetDouble(7));
                }

            }
            dr.Close();
            jtlj.ExportDict();
            #endregion

            #region 煤粉质量

            Shuju mfbr = new Shuju(this);
            cmd.CommandText = "select  pz ,round(avg(c),2) as c,round(avg(sf),2) as shuif,round(avg(huifen),2) as huifen,round(avg(huifa),2) as huifa,round(avg(s),2) as s,round(avg(farezhi),2) as farezhi,round(avg(xidu),2) as xidu from ddmf where trunc(sj)=:rq   group by pz";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                string mc = dr.IsDBNull(0) ? "" : dr.GetString(0);
                if (mc == "烟煤")
                {
                    if (!dr.IsDBNull(1)) mfbr.AddToDict("烟煤化验分析.质量1.C固", dr.GetDouble(1));
                    if (!dr.IsDBNull(3)) mfbr.AddToDict("烟煤化验分析.质量1.A", dr.GetDouble(3));
                    if (!dr.IsDBNull(4)) mfbr.AddToDict("烟煤化验分析.质量1.V", dr.GetDouble(4));
                    if (!dr.IsDBNull(5)) mfbr.AddToDict("烟煤化验分析.质量1.S", dr.GetDouble(5));
                }
                else if (mc.Contains("喷煤") || mc.Contains("一期") || mc.Contains("二期"))
                {
                    if (!dr.IsDBNull(1)) mfbr.AddToDict("喷煤化验分析.质量1.C固", dr.GetDouble(1));
                    if (!dr.IsDBNull(3)) mfbr.AddToDict("喷煤化验分析.质量1.A", dr.GetDouble(3));
                    if (!dr.IsDBNull(4)) mfbr.AddToDict("喷煤化验分析.质量1.V", dr.GetDouble(4));
                    if (!dr.IsDBNull(5)) mfbr.AddToDict("喷煤化验分析.质量1.S", dr.GetDouble(5));
                }
            }
            dr.Close();
            mfbr.ExportDict();

            Shuju mflj = new Shuju(this);
            statisticCmd.CommandText = "select pz ,round(avg(c),2) as c,round(avg(sf),2) as shuif,round(avg(huifen),2) as huifen,round(avg(huifa),2) as huifa,round(avg(s),2) as s,round(avg(farezhi),2) as farezhi,round(avg(xidu),2) as xidu from ddmf where sj>=:rq1 and sj<:rq2   group by pz";
            dr = statisticCmd.ExecuteReader();
            while (dr.Read())
            {
                string mc = dr.IsDBNull(0) ? "" : dr.GetString(0);
                if (mc == "烟煤")
                {
                    if (!dr.IsDBNull(1)) mfbr.AddToDict("烟煤化验分析.质量2.C固", dr.GetDouble(1));
                    if (!dr.IsDBNull(3)) mfbr.AddToDict("烟煤化验分析.质量2.A", dr.GetDouble(3));
                    if (!dr.IsDBNull(4)) mfbr.AddToDict("烟煤化验分析.质量2.V", dr.GetDouble(4));
                    if (!dr.IsDBNull(5)) mfbr.AddToDict("烟煤化验分析.质量2.S", dr.GetDouble(5));
                }
                else if (mc.Contains("喷煤") || mc.Contains("一期") || mc.Contains("二期"))
                {
                    if (!dr.IsDBNull(1)) mfbr.AddToDict("喷煤化验分析.质量2.C固", dr.GetDouble(1));
                    if (!dr.IsDBNull(3)) mfbr.AddToDict("喷煤化验分析.质量2.A", dr.GetDouble(3));
                    if (!dr.IsDBNull(4)) mfbr.AddToDict("喷煤化验分析.质量2.V", dr.GetDouble(4));
                    if (!dr.IsDBNull(5)) mfbr.AddToDict("喷煤化验分析.质量2.S", dr.GetDouble(5));
                }
            }
            dr.Close();

            #endregion
            conn.Close();
            LoadFinish();
        }

        class Shuju
        {
            private MemTagModel parent = null;
            private Dictionary<string, List<double>> dict = new Dictionary<string, List<double>>();

            public Shuju(MemTagModel parent)
            {
                this.parent = parent;
            }

            public void AddToDict(string tag, double val)
            {
                if (dict.ContainsKey(tag))
                {
                    dict[tag].Add(val);
                }
                else
                {
                    dict.Add(tag, new List<double>() { val });
                }
            }
            public void ExportDict()
            {
                foreach (var item in dict)
                {
                    double result = 0;
                    if (item.Value.Count > 0)
                    {
                        foreach (var v in item.Value)
                        {
                            result += v;
                        }
                        result /= item.Value.Count;
                    }
                    this.parent.SetTagValue(item.Key, result);
                }
            }
        }

        private void LoadByRq(DateTime rq, OracleConnection conn)
        {
            this.RQ = rq;
            foreach (var item in this)
            {
                item.clear();
            }

            SetTagValue("报表日期", rq.ToString("yyyy年MM月dd日"));


            OracleCommand selectCmd = new OracleCommand();
            selectCmd.Connection = conn;
            selectCmd.CommandText = "SELECT TAGNAME,VAL,STRVAL,ISLOCK,ROWID FROM RD_RI WHERE RQ=:RQ";
            selectCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = rq;

            OracleDataReader dr = selectCmd.ExecuteReader();
            while (dr.Read())
            {
                string tag_name = "";
                double val = 0;
                string strval = "";
                if (dr.IsDBNull(0)) tag_name = ""; else tag_name = dr.GetString(0);
                if (dr.IsDBNull(1)) val = 0; else val = dr.GetDouble(1);
                if (dr.IsDBNull(2)) strval = ""; else strval = dr.GetString(2);

                this.SetDbValue(tag_name, val);
                if (strval != "")
                    this.SetDbValue(tag_name, strval);

            }
            dr.Close();

        }

        public void Load(DateTime rq)
        {

            OracleConnection conn = new OracleConnection();
            conn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
            conn.Open();
            LoadByRq(rq, conn);
            conn.Close();
            LoadFinish();
        }

        public DateTime? RQ
        {
            get;
            set;
        }

        public void Save()
        {
            OracleConnection conn = new OracleConnection();
            conn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
            conn.Open();
            //OracleTransaction trans = conn.BeginTransaction();
            foreach (var item in this.alltag)
            {
                MemTag tag = item.Value;
                if (tag.DataState == DataStateType.Add)
                {

                    OracleCommand insertCmd = new OracleCommand();
                    insertCmd.Connection = conn;
                    // insertCmd.Transaction = trans;
                    insertCmd.CommandText = "INSERT INTO RD_RI(RQ,TAGNAME,VAL,STRVAL,ISLOCK) VALUES(:RQ,:TAGNAME,:VAL,:STRVAL,:ISLOCK)";
                    insertCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = this.RQ;
                    insertCmd.Parameters.Add(":TAGNAME", OracleType.VarChar, 200).Value = tag.TagName;
                    insertCmd.Parameters.Add(":VAL", OracleType.Double).Value = (object)tag.TagVal ?? DBNull.Value;
                    insertCmd.Parameters.Add(":STRVAL", OracleType.VarChar, 500).Value = tag.StrVal;
                    insertCmd.Parameters.Add(":ISLOCK", OracleType.Byte).Value = tag.IsLock;
                    insertCmd.ExecuteNonQuery();
                    tag.DbInit(true);
                }
                else if (tag.DataState == DataStateType.Update)
                {
                    // OracleConnection con = trans.Connection;
                    OracleCommand updateCmd = new OracleCommand();
                    updateCmd.Connection = conn;
                    //updateCmd.Transaction = trans;
                    updateCmd.CommandText = "UPDATE RD_RI SET VAL=:VAL,STRVAL=:STRVAL,ISLOCK=:ISLOCK WHERE RQ=:RQ and TAGNAME=:TAGNAME";
                    updateCmd.Parameters.Add(":VAL", OracleType.Double).Value = (object)tag.TagVal ?? DBNull.Value;
                    updateCmd.Parameters.Add(":STRVAL", OracleType.VarChar, 500).Value = tag.StrVal;
                    updateCmd.Parameters.Add(":ISLOCK", OracleType.Double).Value = tag.IsLock;
                    updateCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = this.RQ;
                    updateCmd.Parameters.Add(":TAGNAME", OracleType.VarChar, 200).Value = tag.TagName;

                    int count = updateCmd.ExecuteNonQuery();
                    if (count > 0)
                    {
                        tag.DbInit(true);
                    }
                    else
                    {
                        tag.DbInit(false);
                    }
                } 
                if (tag.DataState == DataStateType.Delete)
                {
                    // OracleConnection con = trans.Connection;
                    OracleCommand deleteCmd = new OracleCommand();
                    deleteCmd.Connection = conn;
                    //  deleteCmd.Transaction = trans;
                    deleteCmd.CommandText = "DELETE FROM RD_RI WHERE RQ=:RQ and TAGNAME=:TAGNAME";
                    deleteCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = this.RQ;
                    deleteCmd.Parameters.Add(":TAGNAME", OracleType.VarChar, 200).Value = tag.TagName;
                    deleteCmd.ExecuteNonQuery();
                    tag.DbInit(false);
                }
            }
            conn.Close();
        }

    }


}
