using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BitServiceExcelTest
{
    public partial class Form1 : Form
    {
        private const int SheetIndex = 0;

        // La liberia NPOI inizia a contare le rige e colonne da 0  ( A1 => { 0, 0 } )
        private const int D4_Row = 3;
        private const int D4_Col = 3;

        private const int E5_Row = 4;
        private const int E5_Col = 4;

        private const int F13_Row = 12;
        private const int F13_Col = 5;

        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Click del bottone "Esegui"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            // apertura file excel
            using (FileStream file = new FileStream("prova-excel.xlsx", FileMode.Open, FileAccess.ReadWrite))
            {
                // recupero foglio di lavoro
                XSSFWorkbook xssfworkbook = new XSSFWorkbook(file);
                ISheet sheet = xssfworkbook.GetSheetAt(SheetIndex);

                // impostazione dei valori letti dalla maschera
                sheet.GetRow(D4_Row).Cells[D4_Col].SetCellValue(this.numericUpDown1.Value.ToString());
                sheet.GetRow(E5_Row).Cells[E5_Col].SetCellValue(this.numericUpDown2.Value.ToString());

                // valutazione delle formule del file
                XSSFFormulaEvaluator.EvaluateAllFormulaCells(xssfworkbook);

                // lettura della cella risultato
                ICell cell = sheet.GetRow(F13_Row).GetCell(F13_Col);

                // aggiornamento maschera
                double value = cell.NumericCellValue;
                this.numericUpDown3.Text = value.ToString();

            }   // chiusura file excel

        }
    }
}
