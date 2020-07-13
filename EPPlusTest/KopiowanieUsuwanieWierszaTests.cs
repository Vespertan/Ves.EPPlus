using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest
{
    [TestClass]
    public class KopiowanieUsuwanieWierszaTests
    {
        [TestMethod]
        public void UsuwanieKopiowanieJednegoWiersza()
        {
            var plik = new ExcelPackage();
            plik.Workbook.Worksheets.Add("Arkusz1");
            var arkusz = plik.Workbook.Worksheets[1];
            arkusz.Cells["A1:Z24"].Value = null;

            arkusz.Cells["A1:B1"].Merge = true;
            arkusz.Cells["A1:B1"].Value = "Jakiś tekst";
            arkusz.Cells["D1:E1"].Merge = true;
            arkusz.Cells["D1:E1"].Value = "Jakiś tekst2";

            arkusz.Cells["A2:C2"].Merge = true;
            arkusz.Cells["A2:C2"].Value = "Jakiś tekst3";
            arkusz.Cells["E2:G2"].Merge = true;
            arkusz.Cells["E2:G2"].Value = "Jakiś tekst4";

            arkusz.Cells["A3:D3"].Merge = true;
            arkusz.Cells["A3:D3"].Value = "Jakiś tekst5";
            arkusz.Cells["F3:I3"].Merge = true;
            arkusz.Cells["F3:I3"].Value = "Jakiś tekst6";

            arkusz.DeleteRow(2);
            arkusz.Cells["A2:I2"].Copy(arkusz.Cells["A3:I3"]);
            arkusz.DeleteRow(3);
            arkusz.Cells["A2:I2"].Copy(arkusz.Cells["A3:I3"]);
            arkusz.DeleteRow(3);
            arkusz.Cells["A2:I2"].Copy(arkusz.Cells["A3:I3"]);
            arkusz.DeleteRow(3);
            arkusz.Cells["A2:I2"].Copy(arkusz.Cells["A3:I3"]);
            arkusz.DeleteRow(3);
            arkusz.Cells["A2:I2"].Copy(arkusz.Cells["A3:I3"]);

            Assert.AreEqual("Jakiś tekst5", (string)arkusz.Cells["B2"].Value);
            Assert.AreEqual("Jakiś tekst6", (string)arkusz.Cells["F3"].Value);
            Assert.AreEqual(true, arkusz.Cells["F2:I3"].Merge);
            Assert.AreEqual(true, arkusz.Cells["A3:D3"].Merge);

            if (System.Diagnostics.Debugger.IsAttached)
            {
                plik.SaveAs(new FileInfo("Rezultat.xlsx"));
                Process.Start("Rezultat.xlsx");
            }
        }

        [TestMethod]
        public void UsuwanieKopiowanieDwochWierszy()
        {
            var plik = new ExcelPackage();
            plik.Workbook.Worksheets.Add("Arkusz1");
            var arkusz = plik.Workbook.Worksheets[1];
            arkusz.Cells["A1:Z24"].Value = null;

            arkusz.Cells["A1:B1"].Merge = true;
            arkusz.Cells["A1:B1"].Value = "Jakiś tekst";
            arkusz.Cells["C1:D1"].Merge = true;
            arkusz.Cells["C1:D1"].Value = "Jakiś tekst2";

            arkusz.Cells["A2:C2"].Merge = true;
            arkusz.Cells["A2:C2"].Value = "Jakiś tekst3";
            arkusz.Cells["D2:F2"].Merge = true;
            arkusz.Cells["D2:F2"].Value = "Jakiś tekst4";

            arkusz.Cells["A3:D3"].Merge = true;
            arkusz.Cells["A3:D3"].Value = "Jakiś tekst5";
            arkusz.Cells["E3:H3"].Merge = true;
            arkusz.Cells["E3:H3"].Value = "Jakiś tekst6";

            arkusz.DeleteRow(1);
            arkusz.Cells["A1:H2"].Copy(arkusz.Cells["A3:H4"]);
            arkusz.DeleteRow(2, 2);
            arkusz.Cells["A1:H2"].Copy(arkusz.Cells["A3:H4"]);
            arkusz.DeleteRow(2, 2);
            arkusz.Cells["A1:H2"].Copy(arkusz.Cells["A3:H4"]);
            arkusz.DeleteRow(2, 2);
            arkusz.Cells["A1:H2"].Copy(arkusz.Cells["A3:H4"]);

            Assert.AreEqual("Jakiś tekst3", (string)arkusz.Cells["B3"].Value);
            Assert.AreEqual("Jakiś tekst6", (string)arkusz.Cells["F4"].Value);
            Assert.AreEqual(true, arkusz.Cells["D3:F3"].Merge);
            Assert.AreEqual(true, arkusz.Cells["A4:D4"].Merge);

            if (System.Diagnostics.Debugger.IsAttached)
            {
                plik.SaveAs(new FileInfo("Rezultat.xlsx"));
                Process.Start("Rezultat.xlsx");
            }
        }

    }
}
