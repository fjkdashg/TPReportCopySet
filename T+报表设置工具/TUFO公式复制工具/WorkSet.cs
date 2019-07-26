using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TUFO公式复制工具
{
    public class WorkSet
    {
        public WorkPath Work_Path;
        public List<ChildRep> ChildReps;
        public string BaseBNo;
        public string BaseJNo;

        public decimal ChildStartRow=3;
        public decimal BNoCol=4;
        public decimal JNoCol=6;
    }

    public class WorkPath
    {
        public string OuterPath;
        public string BaseRepTemp_Path;
        public string BCD_Path;
        public string JCD_Path;
        public string BZCFZ_Path;
        public string JZCFZ_Path;
        public string BLR_Path;
        public string JLR_Path;

    }

    public class ChildRep
    {
        public string ChildIndex;
        public string ShortTitle;
        public string RepNo;
        public string BNo;
        public string JNo;
    }
}
