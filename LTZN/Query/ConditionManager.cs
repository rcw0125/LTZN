using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.Query
{
    public class ConditionManager
    {
        public static bool CheckOpList(List<OpItem> opList)
        {
            if (opList.Count < 2)
            {
                return false;
            }
            if (opList[0].OpState != OpState.开始)
            {
                return false;
            }

            for (int i = 1; i < opList.Count; i++)
            {
                OpItem lastOpItem = opList[i - 1];
                OpItem curOpItem = opList[i];

                switch (lastOpItem.OpState)
                {
                    case OpState.开始:
                        if (curOpItem.OpState != OpState.左括号 && curOpItem.OpState != OpState.条件)
                        {
                            return false;
                        }
                        break;
                    case OpState.左括号:
                        if (curOpItem.OpState != OpState.左括号 && curOpItem.OpState != OpState.条件)
                        {
                            return false;
                        }
                        break;
                    case OpState.逻辑关系:
                        if (curOpItem.OpState != OpState.左括号 && curOpItem.OpState != OpState.条件)
                        {
                            return false;
                        }
                        break;
                    case OpState.条件:
                        if (curOpItem.OpState != OpState.右括号 && curOpItem.OpState != OpState.逻辑关系 && curOpItem.OpState != OpState.结束)
                        {
                            return false;
                        }
                        break;
                    case OpState.右括号:
                        if (curOpItem.OpState != OpState.右括号 && curOpItem.OpState != OpState.逻辑关系 && curOpItem.OpState != OpState.结束)
                        {
                            return false;
                        }
                        break;
                    case OpState.结束:
                        return true;
                        break;
                }
                if (curOpItem.OpState == OpState.结束)
                {
                    return true;
                }
            }
            return false;
        }

        public static Condition ConvertOpList(List<OpItem> opList)
        {
            Condition result = null;

            Stack<OpItem> opStack = new Stack<OpItem>();
            foreach (var item in opList)
            {
                OpItem lastOpItem = item;
                switch (lastOpItem.OpState)
                {
                    case OpState.开始:
                        opStack.Clear();
                        break;
                    case OpState.左括号:
                        opStack.Push(lastOpItem);
                        break;
                    case OpState.逻辑关系:
                        opStack.Push(lastOpItem);
                        break;
                    case OpState.条件:
                        while (opStack.Count > 0 && opStack.Peek().OpState == OpState.逻辑关系)
                        {
                            OpItem logicItem = opStack.Pop();
                            OpItem d1 = opStack.Pop();
                            if (d1.OpState != OpState.条件)
                            {
                                throw new Exception("查询条件不正确");
                            }
                            ComposeCondition c = new ComposeCondition((Condition)d1.OpValue, (LogicOpType)logicItem.OpValue, (Condition)lastOpItem.OpValue);
                            OpItem np = new OpItem(OpState.条件,c);
                            lastOpItem = np;
                        }
                        opStack.Push(lastOpItem);
                        break;
                    case OpState.右括号:
                        OpItem cond1 = opStack.Pop();
                        OpItem zkh = opStack.Pop();
                        if (cond1.OpState != OpState.条件 || zkh.OpState != OpState.左括号)
                        {
                            throw new Exception("查询条件不正确");
                        }
                        lastOpItem = cond1;
                        while (opStack.Count > 0 && opStack.Peek().OpState == OpState.逻辑关系)
                        {
                            OpItem logicItem = opStack.Pop();
                            OpItem d1 = opStack.Pop();
                            if (d1.OpState != OpState.条件)
                            {
                                throw new Exception("查询条件不正确");
                            }
                            ComposeCondition c = new ComposeCondition((Condition)d1.OpValue, (LogicOpType)logicItem.OpValue, (Condition)lastOpItem.OpValue);
                            OpItem np = new OpItem(OpState.条件, c);
                            lastOpItem = np;
                        }
                        opStack.Push(lastOpItem);
                        break;
                    case OpState.结束:
                        if (opStack.Count == 1)
                        {
                            OpItem resultItem = opStack.Pop();
                            if (resultItem.OpState == OpState.条件)
                            {
                                result = (Condition)resultItem.OpValue;
                            }
                        }
                        break;
                }
            } 
            return result;
        }
    }
}
