using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;

namespace VCO
{
    public partial class CostOptimization : Form
    {
        public List<Data> SimAndUsage { get; set; }
        private static readonly Regex ColumnNameRegex = new Regex("[A-Za-z]+");

        private List<int> planList = new List<int> {10240,5120,1024,500,250,3};
        private List<Tuple<List<Data>, double>> planAssignments = new List<Tuple<List<Data>, double>>();
        public CostOptimization()
        {
            InitializeComponent();
            SimAndUsage = new List<Data>();
        }
        
        private static string GetColumnName(string cellReference)
        {
            if (ColumnNameRegex.IsMatch(cellReference))
                return ColumnNameRegex.Match(cellReference).Value;

            throw new ArgumentOutOfRangeException(cellReference);
        }
        
        public void ReadFile(string path)
        {
            Cursor.Current = Cursors.WaitCursor;
            using (var document = SpreadsheetDocument.Open(path, true))
            {
                var sheets = document.WorkbookPart.Workbook.Descendants<Sheet>();
                foreach (Sheet sheet in sheets)
                {
                    WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id);
                    Worksheet worksheet = worksheetPart.Worksheet;
                    var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>();
                    int rowCount = 0;
                    foreach (var row in rows)
                    {
                        rowCount++;
                        if (rowCount == 1)
                        {
                            continue;
                        }
                        long simNum = -1;
                        double simUsage = 0;
                        var cells = row.Elements<Cell>();
                        bool insertValue = false;
                        foreach (var cell in cells)
                        {
                            if (GetColumnName(cell.CellReference) == "A")
                            {
                                var str = cell.CellValue.Text;
                                simNum = Convert.ToInt64(cell.CellValue.Text);
                                insertValue = true;
                            }
                            if (GetColumnName(cell.CellReference) == "M")
                            {
                                var str = cell.CellValue.Text;
                                simUsage = Convert.ToDouble(cell.CellValue.Text);
                                insertValue = true;
                            }
                        }
                        if (insertValue)
                        {
                            SimAndUsage.Add(new Data { Sim = simNum, Usage = simUsage });
                        }
                    }
                }
            }
            Cursor.Current = Cursors.Default;
            var planSubsets = FindSubsets(planList).ToList();
            planSubsets.RemoveAt(0); //Removing empty set
            int counter = 0;
            foreach (var plans in planSubsets)
            {
                if (plans.Any())
                {
                    var plansDesc = plans.OrderByDescending(x => x);
                    CalculatePlans(plansDesc.ToList());
                }
                counter++;
            }

            UpdateFile(path);
        }

        private void UpdateFile(string path)
        {
            FileInfo fileInfo = new FileInfo(path);
            ExcelPackage p = new ExcelPackage(fileInfo);
            ExcelWorksheet myWorksheet = p.Workbook.Worksheets["3290846DeDuped"];
            myWorksheet.Cells[5873, 26].Value = 1000000;
            p.Save();
        }

       

        private void CalculatePlans(List<int> plans )
        {
            double poolCommitment = 0;
            double accumulatedUsage = 0;
            var index = -1;
            var planIndex = 0;
            bool planTransition = false;
            double totalCost = 0;
            for (var i = 0; i < SimAndUsage.Count; i++)
            {
                while (planTransition && planIndex < plans.Count - 1 && plans[planIndex] >= SimAndUsage[i].Usage)
                {
                    planIndex++;
                    poolCommitment = 0;
                    accumulatedUsage = 0;
                }
                planTransition = false;
                if (planIndex == plans.Count - 1)
                {
                    AssignPlan(ref poolCommitment, ref accumulatedUsage, plans[plans.Count - 1], ref totalCost, i);
                }
                else
                {
                    AssignPlan(ref poolCommitment, ref accumulatedUsage, plans[planIndex], ref totalCost, i);
                    if (poolCommitment > accumulatedUsage * PlanInformation.GetInfoBySize(plans[planIndex]).Buffer)
                    {
                        planIndex++;
                        poolCommitment = 0;
                        accumulatedUsage = 0;
                        planTransition = true;
                    }
                }
            }
            if (accumulatedUsage > poolCommitment)
            {
                totalCost += (accumulatedUsage - poolCommitment) * PlanInformation.GetInfoBySize(3).OverageCost;
            }
            planAssignments.Add(new Tuple<List<Data>, double> (SimAndUsage, totalCost));
        }

        private void AssignPlan(ref double poolCommitment, ref double accumulatedUsage, int plan, ref double totalCost, int i)
        {
            accumulatedUsage += SimAndUsage[i].Usage;
            poolCommitment += plan;
            SimAndUsage[i].Plan = plan;
            SimAndUsage[i].Cost = PlanInformation.GetInfoBySize(plan).Cost;
            SimAndUsage[i].PlanAssigned = true;
            totalCost += SimAndUsage[i].Cost;
        }

        public IEnumerable<IEnumerable<T>> FindSubsets<T>(IEnumerable<T> source)
        {
            List<T> list = source.ToList();
            int length = list.Count;
            int max = (int)Math.Pow(2, list.Count);

            for (int count = 0; count < max; count++)
            {
                List<T> subset = new List<T>();
                uint rs = 0;
                while (rs < length)
                {
                    if ((count & (1u << (int)rs)) > 0)
                    {
                        subset.Add(list[(int)rs]);
                    }
                    rs++;
                }
                yield return subset;
            }
        }

        private void fileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Select File";

            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ReadFile(openFileDialog1.FileName);
            }
        }
    }
}
