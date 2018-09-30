using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 倒表工具
{
    class Program
    {
        static void Main(string[] args)
        {
            while(true)
            {
                Console.Write("请输入Excel文件路径或直接拖入Excel文件(输入exit退出)：");
                String filename = Console.ReadLine();
                if (filename.ToUpper() == "EXIT")
                    break;
                ExcelOperation eo = new ExcelOperation(filename);
                eo.Excel2Docx();
                Console.WriteLine("表格已转换完毕，请在原表格文件夹查看！");
            }
            
        }
    }
}
