using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest
{
    [TestClass]
    public class UsuwanieKolumnZObrazkiemTests
    {
        [TestMethod]
        public void UsuniecieKolumnPrzedObrazkiem()
        {
            var plik = new ExcelPackage();
            plik.Workbook.Worksheets.Add("Arkusz1");
            var arkusz = plik.Workbook.Worksheets[1];
            arkusz.Cells["C5:D8"].Merge = true;
            var obrazek = Image.FromFile("ObrazekTestowy.png");
            var obrazekExcel = arkusz.Drawings.AddPicture("obrazek", obrazek);
            var yObrazka = GetHeightInPixels(arkusz.Cells["A1"]) * 4 + 13;
            var xObrazka = GetWidthInPixels(arkusz.Cells["A1"]) * 2 + 23;
            obrazekExcel.SetPosition(yObrazka, xObrazka);
            obrazekExcel.EditAs = OfficeOpenXml.Drawing.eEditAs.TwoCell;
            arkusz.DeleteColumn(1, 2);

            if (System.Diagnostics.Debugger.IsAttached)
            {
                plik.SaveAs(new FileInfo("Rezultat.xlsx"));
                Process.Start("Rezultat.xlsx");
            }

            Assert.AreEqual(1, obrazekExcel.From.Column + 1);
            Assert.AreEqual(2, obrazekExcel.To.Column + 1);
        }

        [TestMethod]
        public void UsuniecieKolumnyWewnatrzObrazkaMogacegoZmieniacRozmiar()
        {
            var plik = new ExcelPackage();
            plik.Workbook.Worksheets.Add("Arkusz1");
            var arkusz = plik.Workbook.Worksheets[1];
            arkusz.Cells["C5:D8"].Merge = true;
            var obrazek = Image.FromFile("ObrazekTestowy.png");
            var obrazekExcel = arkusz.Drawings.AddPicture("obrazek", obrazek);
            var yObrazka = GetHeightInPixels(arkusz.Cells["A1"]) * 4 + 13;
            var xObrazka = GetWidthInPixels(arkusz.Cells["A1"]) * 2 + 23;
            obrazekExcel.SetPosition(yObrazka, xObrazka);
            obrazekExcel.EditAs = OfficeOpenXml.Drawing.eEditAs.TwoCell;
            arkusz.DeleteColumn(3);

            if (System.Diagnostics.Debugger.IsAttached)
            {
                plik.SaveAs(new FileInfo("Rezultat.xlsx"));
                Process.Start("Rezultat.xlsx");
            }

            Assert.AreEqual(3, obrazekExcel.To.Column + 1);
        }

        [TestMethod]
        public void UsuniecieKolumnyObejmujacejObrazekMajacyStalyRozmiar()
        {
            var plik = new ExcelPackage();
            plik.Workbook.Worksheets.Add("Arkusz1");
            var arkusz = plik.Workbook.Worksheets[1];
            arkusz.Cells["C5:D8"].Merge = true;
            var obrazek = Image.FromFile("ObrazekTestowy.png");
            var obrazekExcel = arkusz.Drawings.AddPicture("obrazek", obrazek);
            var yObrazka = GetHeightInPixels(arkusz.Cells["A1"]) * 4 + 13;
            var xObrazka = GetWidthInPixels(arkusz.Cells["A1"]) * 2 + 23;
            obrazekExcel.SetPosition(yObrazka, xObrazka);
            obrazekExcel.EditAs = OfficeOpenXml.Drawing.eEditAs.OneCell;
            arkusz.DeleteColumn(4);

            if (System.Diagnostics.Debugger.IsAttached)
            {
                plik.SaveAs(new FileInfo("Rezultat.xlsx"));
                Process.Start("Rezultat.xlsx");
            }

            Assert.AreEqual(4, obrazekExcel.To.Column + 1);
        }

        [TestMethod]
        public void UsuniecieKolumnyPrzedObrazkiemMajacymStalyRozmiarINieprzesuwanym()
        {
            var plik = new ExcelPackage();
            plik.Workbook.Worksheets.Add("Arkusz1");
            var arkusz = plik.Workbook.Worksheets[1];
            arkusz.Cells["C5:D8"].Merge = true;
            var obrazek = Image.FromFile("ObrazekTestowy.png");
            var obrazekExcel = arkusz.Drawings.AddPicture("obrazek", obrazek);
            var yObrazka = GetHeightInPixels(arkusz.Cells["A1"]) * 4 + 13;
            var xObrazka = GetWidthInPixels(arkusz.Cells["A1"]) * 2 + 23;
            obrazekExcel.SetPosition(yObrazka, xObrazka);
            obrazekExcel.EditAs = OfficeOpenXml.Drawing.eEditAs.Absolute;
            arkusz.DeleteColumn(1, 3);

            if (System.Diagnostics.Debugger.IsAttached)
            {
                plik.SaveAs(new FileInfo("Rezultat.xlsx"));
                Process.Start("Rezultat.xlsx");
            }

            Assert.AreEqual(3, obrazekExcel.From.Column + 1);
            Assert.AreEqual(4, obrazekExcel.To.Column + 1);
        }

        private int GetHeightInPixels(ExcelRange cell)
        {
            using (Graphics graphics = Graphics.FromHwnd(IntPtr.Zero))
            {
                float dpiY = graphics.DpiY;
                return (int)(cell.Worksheet.Row(cell.Start.Row).Height * (1 / 72.0) * dpiY);
            }
        }

        private int GetWidthInPixels(ExcelRange cell)
        {
            double columnWidth = cell.Worksheet.Column(cell.Start.Column).Width;
            Font font = new Font(cell.Style.Font.Name, cell.Style.Font.Size, FontStyle.Regular);

            double pxBaseline = Math.Round(MeasureString("1234567890", font) / 10);

            return (int)(columnWidth * pxBaseline);
        }

        public float MeasureString(string s, Font font)
        {
            using (var g = Graphics.FromHwnd(IntPtr.Zero))
            {
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

                return g.MeasureString(s, font, int.MaxValue, StringFormat.GenericTypographic).Width;
            }
        }
    }
}
