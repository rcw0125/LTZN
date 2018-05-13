using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.CalFramework2
{
    /// <summary>
    /// IOper 的摘要说明。
    /// 计算符接口
    /// </summary>
    public interface IOper
    {
        /// <summary>
        /// 计算符计算接口计算方法
        /// </summary>
        /// <param name="o1">参数1</param>
        /// <param name="o2">参数2</param>
        /// <returns></returns>
        object Oper(object o1, object o2);
    }
}
