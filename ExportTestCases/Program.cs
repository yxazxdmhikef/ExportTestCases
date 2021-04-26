using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportTestCases
{
    class Program
    {
        static void Main(string[] args)
        {
            //укажем все необходимые значения
            String tfs = @"http://your-tfs-server/tfs";
            String project = "project-name"; 
            String plan = "Сценарий демонстрации";

            //экспорт
            ExportTest export = new ExportTest(tfs, project, plan);            
            export.ExportTestCases();
        }
    }
}
