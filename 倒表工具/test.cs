using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IronPython.Runtime;
using IronPython;
using IronPython.Hosting;
using Microsoft.Scripting.Hosting;
using Spire.Xls;

namespace 倒表工具
{
    public class Test
    {
        public Test()
        {

        }

        public Workbook TurnExcel(String filename)
        {
            Workbook wb = new Workbook();
            ScriptEngine pyEngine = Python.CreateEngine();
            dynamic py = pyEngine.ExecuteFile(@"\script\readexcel.py");
            wb = py.read_excel(filename);

            return wb;
        }
        
    }
}
