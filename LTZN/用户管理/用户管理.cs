using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.用户管理
{
    public class 用户管理
    {
        public string 当前用户 = "";
        public 权限 用户权限 = 权限.调度 | 权限.统计;

    }
    public enum 权限 { 调度, 统计 }

}
