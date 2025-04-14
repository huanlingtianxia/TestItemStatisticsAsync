using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestItemStatisticsAcync.ExcelOperation
{
    public class ParametersTestItem
    {
        // 测试SN个数
        public int NumSN { get; set; }

        // 数据源行开始
        public int StartRow { get; set; }

        // 数据源列开始
        public int StartCol { get; set; }

        // 数据源行结束
        public int EndRow { get; set; }

        // 数据源列结束
        public int EndtCol { get; set; }

        // 目标行开始
        public int StartRowDest { get; set; }

        // 目标列开始
        public int StartColDest { get; set; }

        // 单个SN的测试次数，即单个SN测试项跨度单元格数量
        public int Repeat { get; set; }

        // 测试项数量
        public int TotalItemCount { get; set; }

        // 来源Sheet名称
        public string FromSheet { get; set; } = "SortSelectTrans";

        // 目标Sheet名称
        public string ToSheet { get; set; } = "toSheetAll";
        // 源路径
        public string SourcePath { get; set; }
        // 目标路径
        public string TargetPath { get; set; }

        // sheet name
        public string[] SheetName { get; set; }
        // position sheet name
        public string PosSheetName { get; set; }

        // 删除sheet时保留sheet个数
        public int ReserveSheetCount { get; set; }
        // cell parameters
        public int[][] CellParams { get; set; }
        /* formula param*/
        // Low limit
        public string LowLimit { get; set; } = $"!G4";
        // High limit 
        public string HighLimit { get; set; } = $"!G3";
        // CP
        public string CPValue { get; set; } = $"!L20";
        // CPK 
        public string CPKValue { get; set; } = $"!N20";
        // GRR Value
        public string GRRValue { get; set; } = $"!L24";
    }

    public enum GeneralMode
    {
        CellCopyPaste = 0,
        CellDelete = 1,
        CellSummary = 2,
        SheetOperater = 3,
        
    }
}
