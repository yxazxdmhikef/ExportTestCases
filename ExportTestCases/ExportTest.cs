using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExportTestCases
{

    public class ExportTest
    {
        /// <summary>
        /// Наименование файла Excel для тестовых случаев, он уже должен быть создан по указанному в filePath пути
        /// </summary>
        private String testFileName = "TestCases";
        /// <summary>
        /// Наименование htm файла для экспорта ошибок.
        /// </summary>
        private String bugsFileName = "Bugs";
        /// <summary>
        /// Локальный путь для экспорта файлов.
        /// </summary>
        private String filePath = "file\\path";
        /// <summary>
        /// Наименование проекта
        /// </summary>
        private String projectName = "project-name";
        /// <summary>
        /// Uri адрес TFS
        /// </summary>
        private String Uri = @"http://tfs-server/tfs";
        /// <summary>
        /// Наименование плана тестирования.
        /// </summary>
        private String testPlanName = "";

        /// <summary>
        /// Uri TFS.
        /// </summary>
        private Uri tfsUri;
        /// <summary>
        /// Коллекция командных проектов TFS.
        /// </summary>
        private TfsTeamProjectCollection tfsTeamProjectCollection;
        /// <summary>
        /// Хранилище рабочих элементов.
        /// </summary>
        private WorkItemStore workItemstore;
        /// <summary>
        /// Командный проект.
        /// </summary>
        private ITestManagementTeamProject teamProject;

        private Microsoft.Office.Interop.Excel.Application xlApp;
        private Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
        private object misValue = System.Reflection.Missing.Value;
        private Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
        private Microsoft.Office.Interop.Excel.Range chartRange;

        string upperBound = "a";
        string lowerBound = "a";
        string upperBound2 = "b";
        string lowerBound2 = "b";

        int testNumber = 1;
        int row = 1;



        private ITestPlanCollection plans;
        private ITestSuiteEntryCollection testSuites;
        private ITestSuiteEntry suite;


        /// <summary>
        /// Инициализирует новый экземпляр класса ExportTest.
        /// </summary>
        /// <param name="tfs">Uri адрес TFS.</param>
        /// <param name="project">Наименование проекта в TFS.</param>
        /// <param name="plan">Наименование плана тестирования.</param>       
        public ExportTest(String tfs, String project, String plan)
        {
            if (
                String.IsNullOrEmpty(tfs) ||
                String.IsNullOrEmpty(project) ||
                String.IsNullOrEmpty(plan)                
                )
                throw new Exception("Значение входных параметров не может быть пустым!");

            Uri = tfs;
            projectName = project;
            testPlanName = plan;
            FillVariables();
        }

        /// <summary>
        /// Наполняет переменные.
        /// </summary>
        private void FillVariables()
        {
            tfsUri = new Uri(Uri);
            tfsTeamProjectCollection = new TfsTeamProjectCollection(tfsUri);
            workItemstore = tfsTeamProjectCollection.GetService<WorkItemStore>();

            ITestManagementService service = (ITestManagementService)tfsTeamProjectCollection.GetService(typeof(ITestManagementService));
            ITestManagementTeamProject testManagementTeamProject = service.GetTeamProject(projectName);
            teamProject = testManagementTeamProject;
        }



        /// <summary>
        /// Создает заголовок для htm-файла ошибок.
        /// </summary>
        /// <returns>Html текст заголовка таблицы ошибок.</returns>
        private String CreateBeginContent()
        {
            String contents = "";
            contents += @"<html>";
            contents += @"<body>";
            contents += "<table border=\"1\">";
            contents += @"<tr>";
            contents += String.Format(@"<td>{0}</td>", "ID");
            contents += String.Format(@"<td>{0}</td>", "НАИМЕНОВАНИЕ");
            contents += String.Format(@"<td>{0}</td>", "ШАГИ ДЛЯ ВОСПРОИЗВЕДЕНИЯ");
            contents += String.Format(@"<td>{0}</td>", "ОПИСАНИЕ");
            contents += @"</tr>";
            return contents;
        }

        /// <summary>
        /// Создает окончание для htm-файла ошибок.
        /// </summary>
        /// <returns>Html текст окончания таблицы ошибок.</returns>
        private String CreateEndContent()
        {
            String contents = "";
            contents += @"</table>";
            contents += @"</body>";
            contents += @"</html>";
            return contents;
        }

        /// <summary>
        /// Создает строку ошибки.
        /// </summary>
        /// <param name="bug">Ошибка.</param>
        /// <returns>Html текст строки ошибки в таблице.</returns>
        private String CreateBugRow(WorkItem bug)
        {
            String row = "";
            row += @"<tr>";
            row += String.Format(@"<td>{0}</td>", bug.Id.ToString());
            row += String.Format(@"<td>{0}</td>", bug.Title);
            row += String.Format(@"<td>{0}</td>", ((bug.Fields["Шаги для воспроизведения"].Value != null) ? bug.Fields["Шаги для воспроизведения"].Value.ToString() : ""));
            row += String.Format(@"<td>{0}</td>", bug.Description);
            row += @"</tr>";
            return row;
        }

        /// <summary>
        /// Экспортирует ошибки в htm-файл.
        /// </summary>
        /// <param name="path">Путь для эексопрта.</param>
        /// <param name="fileName">Наименование файла.</param>
        public void ExportBugs(String path, String fileName)
        {
            if (String.IsNullOrEmpty(path) || String.IsNullOrEmpty(fileName))
                throw new Exception("Путь и/или имя файла не может быть пустым!");

            string wiql = "SELECT * FROM WorkItems WHERE [System.TeamProject] = '" + projectName + "' ORDER BY [System.Id] ";
            WorkItemCollection wic = workItemstore.Query(wiql);
            String contents = CreateBeginContent();
            foreach (WorkItem wi in wic)
            {
                if (wi.Type.Name == "Ошибка")
                    contents += CreateBugRow(wi);
            }
            contents += CreateEndContent();
            System.IO.File.WriteAllText(path + "\\" + fileName + ".htm", contents);
        }

        /// <summary>
        /// Экспортирует тестовые случаи в Excel файл.
        /// </summary>
        public void ExportTestCases()
        {
            try
            {                 
                CreateExcelTableSettings();
                //по тестовым ситуациям плана
                try
                {
                    IList<ITestCase> planCases = teamProject.TestCases.Query("Select * From TestCase").ToList();
                    if (planCases != null && planCases.Count > 0)
                        ByTestCases(planCases);
                }
                catch { }                   
               
                //по тестовым случаям
                ByTestSuites();

                EndingExcelSettings();

                xlWorkBook.SaveAs(filePath + "\\" + testFileName + ".xls", Excel.XlFileFormat.xlWorkbookNormal,
                        misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlShared, misValue, misValue, misValue, misValue, misValue);

                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();
                releaseObject(xlApp);
                releaseObject(xlWorkBook);
                releaseObject(xlWorkSheet);

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        /// <summary>
        /// Освобождает объект.
        /// </summary>
        /// <param name="obj"></param>
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        /// <summary>
        /// Задает натсройки для Excel таблицы по окончанию обрабоки.
        /// </summary>
        private void EndingExcelSettings()
        {
            lowerBound = "d";
            lowerBound += (row - 1);
            chartRange = xlWorkSheet.get_Range("a1", "d1");
            chartRange.Font.Bold = true;
            chartRange.Interior.Color = 18018018;

            chartRange = xlWorkSheet.get_Range("a1", lowerBound);
            chartRange.Cells.WrapText = true;
            chartRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            chartRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
        }


        /// <summary>
        /// Проходит по всем тестовым ситуациям тестового плана и заполняет Excel таблицу, 
        /// указывая подзаголовком категорию теста - наименования тестовой ситуации.
        /// </summary>
        private void ByTestSuites()
        {
            try
            {
                this.plans = teamProject.TestPlans.Query("Select * From TestPlan");
                foreach (ITestPlan plan in plans)
                {
                    if (plan.Name == testPlanName)
                        this.testSuites = plan.RootSuite.Entries;
                }
                foreach (ITestSuiteEntry suite_entry in testSuites)
                {
                    this.suite = suite_entry;
                    IStaticTestSuite newSuite = suite_entry.TestSuite as IStaticTestSuite;
                    if (newSuite != null)
                    {
                        xlWorkSheet.get_Range("a" + row, "d" + row).Merge(true);
                        xlWorkSheet.Cells[row, 1] = newSuite.Title;
                        xlWorkSheet.Cells[row, 1].Font.Bold = true;
                        xlWorkSheet.Cells[row, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        row = row + 1;
                        ITestCaseCollection cases = newSuite.AllTestCases;
                        ByTestCases(cases);
                    }
                    else
                    {
                        ITestCase test = suite_entry.TestCase as ITestCase;
                        if (test != null)
                            ByTest(test);
                    }
                }
            }
            catch { }
            
        }

        /// <summary>
        /// Заполняет Excel таблицу по тестовому случаю.
        /// </summary>
        /// <param name="testCase">Тестовый случай.</param>
        private void ByTest(ITestCase testCase)
        {
            xlWorkSheet.Cells[row, 1] = testNumber.ToString() + @".";
            xlWorkSheet.Cells[row, 2] = testCase.Title;

            string upperBound = "a";
            string lowerBound = "a";
            string upperBound2 = "b";
            string lowerBound2 = "b";

            upperBound += row;
            upperBound2 += row;

            TestActionCollection testActions = testCase.Actions;
            BySteps(testActions);

            lowerBound += (row - 1);
            lowerBound2 += (row - 1);

            xlWorkSheet.get_Range(upperBound, lowerBound).Merge(false);

            chartRange = xlWorkSheet.get_Range(upperBound, lowerBound);
            chartRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            chartRange.VerticalAlignment = 1;

            xlWorkSheet.get_Range(upperBound2, lowerBound2).Merge(false);
            chartRange = xlWorkSheet.get_Range(upperBound2, lowerBound2);
            chartRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            chartRange.VerticalAlignment = 1;
            testNumber++;
        }

        /// <summary>
        /// Заполняет Excel таблицу, проходясь по коллеции тестовых случаев.
        /// </summary>
        /// <param name="cases">Коллекция тестовых случаев.</param>
        private void ByTestCases(IList<ITestCase> cases)
        {            
            if (cases != null && cases.Count > 0)
            {
                foreach (ITestCase testCase in cases)
                {
                    ByTest(testCase);
                }
            }
        }

        /// <summary>
        /// Заполняет Excel таблицу данными по шагам тестового случая.
        /// </summary>
        /// <param name="testActions">Шаги теста.</param>
        private void BySteps(TestActionCollection testActions)
        {
            int stepId = 1;
            foreach (var stestStep in testActions)
            {
                ITestStep testStep = stestStep as ITestStep;
                if (testStep != null)
                {
                    xlWorkSheet.Cells[row, 3] = stepId.ToString() + @". " + testStep.Title.ToPlainText();
                    xlWorkSheet.Cells[row, 4] = testStep.ExpectedResult.ToPlainText();
                    stepId++;
                    row++;
                }
                else
                {
                    //расшаренный шаг
                    ISharedStep testStepS = stestStep as ISharedStep;
                    if (testStepS != null)
                    {
                        xlWorkSheet.Cells[row, 3] = testStepS.Title.ToString() + @". " + testStep.Title.ToPlainText();
                        stepId++;
                        row++;
                    }
                }
            }   
        }

        /// <summary>
        /// Создает Excel таблицу, указывает шапку  и др настройки.
        /// </summary>
        private void CreateExcelTableSettings()
        {  
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);            
            xlWorkSheet.Cells[1, 1] = "№";
            xlWorkSheet.Cells[1, 2] = "НАИМЕНОВАНИЕ";
            xlWorkSheet.Cells[1, 3] = "ДЕЙСТВИЯ";
            xlWorkSheet.Cells[1, 4] = "РЕЗУЛЬТАТ";


            xlWorkSheet.Cells[2, 1] = "1";
            xlWorkSheet.Cells[2, 2] = "2";
            xlWorkSheet.Cells[2, 3] = "3";
            xlWorkSheet.Cells[2, 4] = "4";

            xlWorkSheet.Cells[1, 1].Font.Bold = true;
            xlWorkSheet.Cells[1, 2].Font.Bold = true;
            xlWorkSheet.Cells[1, 3].Font.Bold = true;
            xlWorkSheet.Cells[1, 4].Font.Bold = true;
            xlWorkSheet.Cells[2, 1].Font.Bold = true;
            xlWorkSheet.Cells[2, 2].Font.Bold = true;
            xlWorkSheet.Cells[2, 3].Font.Bold = true;
            xlWorkSheet.Cells[2, 4].Font.Bold = true;
            xlWorkSheet.Cells[1, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xlWorkSheet.Cells[1, 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xlWorkSheet.Cells[1, 3].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xlWorkSheet.Cells[1, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xlWorkSheet.Cells[2, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xlWorkSheet.Cells[2, 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xlWorkSheet.Cells[2, 3].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            xlWorkSheet.Cells[2, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            (xlWorkSheet.Columns["A", Type.Missing]).ColumnWidth = 5;
            (xlWorkSheet.Columns["B", Type.Missing]).ColumnWidth = 25;
            (xlWorkSheet.Columns["C", Type.Missing]).ColumnWidth = 60;
            (xlWorkSheet.Columns["D", Type.Missing]).ColumnWidth = 60;
            row = 3;
        }
       
    }
}
