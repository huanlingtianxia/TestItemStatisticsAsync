using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestItem.Excel
{

    public class ParametersTestItem
    {
        // 测试SN个数
        public int NumSN { get; set; } = 8;

        // 数据源行开始
        public int StartRow { get; set; } = 9;

        // 数据源列开始
        public int StartCol { get; set; } = 1;

        // 数据源行结束
        public int EndRow { get; set; } = 9;

        // 数据源列结束
        public int EndtCol { get; set; } = 1;

        // 目标行开始
        public int StartRowDest { get; set; } = 1;

        // 目标列开始
        public int StartColDest { get; set; } = 2;

        // 单个SN的测试次数，即单个SN测试项跨度单元格数量
        public int Span { get; set; } = 9;

        // 测试项数量
        public int TotalItemCount { get; set; } = 229;

        // 来源Sheet名称
        public string FromSheet { get; set; } = "SortSelectTrans";

        // 目标Sheet名称
        public string ToSheet { get; set; } = "toSheetAll";
        // 源路径
        public string SourcePath { get; set; }
        // 目标路径
        public string TargetPath { get; set; }



        // 可以添加更多方法，根据需要执行操作。
    }
    public class ParametersSet
    {
        


    }
}
