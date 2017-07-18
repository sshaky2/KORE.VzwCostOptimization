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
using VCO;

namespace VCO
{
    public partial class CostOptimization : Form
    {
        public List<Data> SimAndUsage { get; set; }
        private static readonly Regex ColumnNameRegex = new Regex("[A-Za-z]+");

        private List<int> planList = new List<int> {10240,5120,1024,500,250,3};
        private List<Tuple<List<Data>, double, List<int>>> planAssignments = new List<Tuple<List<Data>, double, List<int>>>();
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

        public DataTable CreateTable()
        {
            // Here we create a DataTable with four columns.
            DataTable table = new DataTable();
            table.Columns.Add("Size", typeof(int));
            table.Columns.Add("Cost", typeof(double));
            table.Columns.Add("OverageCost", typeof(double));
            table.Columns.Add("Buffer", typeof(double));

            // Here we add five DataRows.
            table.Rows.Add(3, 1, 0.7, 1);
            table.Rows.Add(25, 7, 0.009, 1);
            table.Rows.Add(250, 8, 0.009, 1);
            table.Rows.Add(500, 10, 0.009, 1);
            table.Rows.Add(1024, 15, 0.009, 1);
            table.Rows.Add(5120, 35, 0.009, 1);
            table.Rows.Add(10240, 60, 0.009, 1);
            table.Rows.Add(20480, 125, 0.009, 1);
            table.Rows.Add(30720, 235, 0.009, 1);

            var dataSet = new DataSet();
            dataSet.Tables.Add(table);
            dataSet.WriteXml(@"C:\Users\sshakya\Documents\GitHub\VCO\file.xml");
            return table;

            
        }
        
        public void ReadFile(string path)
        {
            //CreateTable();
            var ds = new DataSet();
            ds.ReadXml(@"C:\Users\sshakya\Documents\GitHub\VCO\file.xml");
            var dt = ds.Tables[0];
            var abc = dt.Select("Size = 250")[0]["Cost"];

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
            var tempList = new List<List<int>>();
            foreach (var plans in planSubsets)
            {

                if (plans.Any())
                {
                    var plansDesc = plans.OrderByDescending(x => x);
                    tempList.Add(plansDesc.ToList());
                    
                }
            }
            var noDupes = tempList.Distinct();
            foreach (var plans in noDupes)
            {
                CalculatePlans(plans);
            }

            //CalculatePlans(new List<int> {10240,1024,500,250,25,3});

            var minCost = double.MaxValue;
            var optimalPlan = new Tuple<List<Data>, double, List<int>>(new List<Data>(), 0, new List<int>());
            foreach (var plan in planAssignments)
            {
                if (plan.Item2 < minCost)
                {
                    minCost = plan.Item2;
                    optimalPlan = plan;
                }
            }
            

            UpdateFile(path);
        }

        private void UpdateFile(string path)
        {
            // Open the document for editing.
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(path, true))
            {
                // Access the main Workbook part, which contains all references.
                WorkbookPart workbookPart = spreadSheet.WorkbookPart;
                // get sheet by name
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == "2973798DeDuped").FirstOrDefault();

                // get worksheetpart by sheet id
                WorksheetPart worksheetPart = workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;

                // The SheetData object will contain all the data.
                //SheetData sheetData = worksheetPart.Worksheet.GetFirstChild();

                Cell cell = GetCell(worksheetPart.Worksheet, "Z", 5867);

                cell.CellValue = new CellValue("10");
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);

                // Save the worksheet.
                worksheetPart.Worksheet.Save();

                // for recacluation of formula
                spreadSheet.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                spreadSheet.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;

            }
        }

        private Cell GetCell(Worksheet worksheet, string columnName, uint rowIndex)
        {
            Row row = GetRow(worksheet, rowIndex);

            if (row == null) return null;
            
            var FirstRow = row.Elements<Cell>().FirstOrDefault(c => string.Compare
            (c.CellReference.Value, columnName +
            rowIndex, true) == 0);

            if (FirstRow == null) return null;

            return FirstRow;
        }

        private Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            Row row = worksheet.GetFirstChild<SheetData>().
            Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row == null)
            {
                throw new ArgumentException($"No row with index {rowIndex} found in spreadsheet");
            }
            return row;
        }
    

        private void CalculatePlans(List<int> plans )
        {
            double poolCommitment = 0;
            double accumulatedUsage = 0;
            var index = -1;
            var planIndex = 0;
            bool planTransition = false;
            double totalCost = 0;
            var planUsed = new List<int>();
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
                    planUsed.Add(plans[plans.Count - 1]);
                    AssignPlan(ref poolCommitment, ref accumulatedUsage, plans[plans.Count - 1], ref totalCost, i);
                }
                else
                {
                    planUsed.Add(plans[planIndex]);
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
            var deduped = planUsed.Distinct().ToList();
            planAssignments.Add(new Tuple<List<Data>, double, List<int>> (SimAndUsage, totalCost, deduped));
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
