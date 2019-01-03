using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Dlubal.RFEM5;

namespace GH_RFEM
{
    public static class GetMaxNumber
    {
       public static int NodeNumber(IModelData data)
        {
            int totalNodesCount = data.GetNodes().Count();
            int lastNodeNo = data.GetNode(totalNodesCount-1, ItemAt.AtIndex).GetData().No;
            return lastNodeNo;
        }
    }
}
