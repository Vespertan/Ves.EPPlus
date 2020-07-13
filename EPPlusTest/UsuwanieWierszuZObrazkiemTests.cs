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
    public class UsuwanieWierszuZObrazkiemTests
    {
        [TestMethod]
        public void PonadObrazkiem()
        {
            var plik = new ExcelPackage();
            plik.Workbook.Worksheets.Add("Arkusz1");
            var arkusz = plik.Workbook.Worksheets[1];
            arkusz.Cells["B5:C8"].Merge = true;
            var obrazek = Image.FromFile("ObrazekTestowy.png");
            var obrazekExcel = arkusz.Drawings.AddPicture("obrazek", obrazek);
            var yObrazka = GetHeightInPixels(arkusz.Cells["A1"]) * 4 + 13;
            var xObrazka = GetWidthInPixels(arkusz.Cells["A1"]) + 23;
            obrazekExcel.SetPosition(yObrazka, xObrazka);
            obrazekExcel.EditAs = OfficeOpenXml.Drawing.eEditAs.TwoCell;
            arkusz.DeleteRow(1, 2);

            if (System.Diagnostics.Debugger.IsAttached)
            {
                plik.SaveAs(new FileInfo("Rezultat.xlsx"));
                Process.Start("Rezultat.xlsx");
            }

            Assert.AreEqual(3, obrazekExcel.From.Row + 1);
            Assert.AreEqual(6, obrazekExcel.To.Row + 1);
        }

        [TestMethod]
        public void WewnatrzObrazkaMogacegoZmieniacRozmiar()
        {
            var plik = new ExcelPackage();
            plik.Workbook.Worksheets.Add("Arkusz1");
            var arkusz = plik.Workbook.Worksheets[1];
            arkusz.Cells["B5:C8"].Merge = true;
            var obrazek = Image.FromFile("ObrazekTestowy.png");
            var obrazekExcel = arkusz.Drawings.AddPicture("obrazek", obrazek);
            var yObrazka = GetHeightInPixels(arkusz.Cells["A1"]) * 4 + 13;
            var xObrazka = GetWidthInPixels(arkusz.Cells["A1"]) + 23;
            obrazekExcel.SetPosition(yObrazka, xObrazka);
            obrazekExcel.EditAs = OfficeOpenXml.Drawing.eEditAs.TwoCell;
            arkusz.DeleteRow(6, 2);

            if (System.Diagnostics.Debugger.IsAttached)
            {
                plik.SaveAs(new FileInfo("Rezultat.xlsx"));
                Process.Start("Rezultat.xlsx");

            }

            Assert.AreEqual(5, obrazekExcel.From.Row + 1);
            Assert.AreEqual(6, obrazekExcel.To.Row + 1);
        }

        [TestMethod]
        public void ObejmujacychObrazekMajacyStalyRozmiar()
        {
            var plik = new ExcelPackage();
            plik.Workbook.Worksheets.Add("Arkusz1");
            var arkusz = plik.Workbook.Worksheets[1];
            arkusz.Cells["B5:C8"].Merge = true;
            var obrazek = Image.FromFile("ObrazekTestowy.png");
            var obrazekExcel = arkusz.Drawings.AddPicture("obrazek", obrazek);
            var yObrazka = GetHeightInPixels(arkusz.Cells["A1"]) * 4 + 13;
            var xObrazka = GetWidthInPixels(arkusz.Cells["A1"]) + 23;
            obrazekExcel.SetPosition(yObrazka, xObrazka);
            obrazekExcel.EditAs = OfficeOpenXml.Drawing.eEditAs.OneCell;
            arkusz.DeleteRow(4, 2);

            if (System.Diagnostics.Debugger.IsAttached)
            {
                plik.SaveAs(new FileInfo("Rezultat.xlsx"));
                Process.Start("Rezultat.xlsx");

            }

            Assert.AreEqual(4, obrazekExcel.From.Row + 1);
            Assert.AreEqual(7, obrazekExcel.To.Row + 1);
        }

        [TestMethod]
        public void WewnatrzObrazkaMajacegoStalyRozmiarINieprzesuwanego()
        {
            var plik = new ExcelPackage();
            plik.Workbook.Worksheets.Add("Arkusz1");
            var arkusz = plik.Workbook.Worksheets[1];
            arkusz.Cells["B5:C8"].Merge = true;
            var obrazek = Image.FromFile("ObrazekTestowy.png");
            var obrazekExcel = arkusz.Drawings.AddPicture("obrazek", obrazek);
            var yObrazka = GetHeightInPixels(arkusz.Cells["A1"]) * 4 + 13;
            var xObrazka = GetWidthInPixels(arkusz.Cells["A1"]) + 23;
            obrazekExcel.SetPosition(yObrazka, xObrazka);
            obrazekExcel.EditAs = OfficeOpenXml.Drawing.eEditAs.Absolute;
            arkusz.DeleteRow(6, 2);

            if (System.Diagnostics.Debugger.IsAttached)
            {
                plik.SaveAs(new FileInfo("Rezultat.xlsx"));
                Process.Start("Rezultat.xlsx");

            }

            Assert.AreEqual(5, obrazekExcel.From.Row + 1);
            Assert.AreEqual(8, obrazekExcel.To.Row + 1);
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
