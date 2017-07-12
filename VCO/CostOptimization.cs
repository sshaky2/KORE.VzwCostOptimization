using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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

namespace VCO
{
    public partial class CostOptimization : Form
    {
        public List<Data> SimAndUsage { get; set; }
        private static readonly Regex ColumnNameRegex = new Regex("[A-Za-z]+");
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
            ApplyPlans();
        }

        private void ApplyPlans()
        {
            double totalPoolCommitment = 0;
            double accumulatedUsage = 0;
            for (var i = 0; i < SimAndUsage.Count; i++)
            {
                accumulatedUsage += SimAndUsage[i].Usage;
                totalPoolCommitment += 10240;
                SimAndUsage[i].Plan = 10240;
                if (totalPoolCommitment > accumulatedUsage)
                {
                    break;
                }
               
            }
        }

        private void fileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Select FIle";

            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ReadFile(openFileDialog1.FileName);
            }
        }
    }
}
