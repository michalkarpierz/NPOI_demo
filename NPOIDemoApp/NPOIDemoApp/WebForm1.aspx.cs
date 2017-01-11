using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace NPOIDemoApp
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            HSSFWorkbook workbook = CreateWorkbook();

            GenerateOutput(workbook);
        }

        HSSFWorkbook CreateWorkbook()
        {
            var workbook = new HSSFWorkbook();

            ISheet sheet = workbook.CreateSheet("NPOI demo sheet");

            //create styles
            IFont fontBold = workbook.CreateFont();
            fontBold.Boldweight = (short)FontBoldWeight.Bold;

            ICellStyle headerStyle = workbook.CreateCellStyle();
            headerStyle.Alignment = HorizontalAlignment.Center;
            headerStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightOrange.Index;
            headerStyle.FillPattern = FillPattern.SolidForeground;
            headerStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Medium; 
            headerStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Medium;
            headerStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Medium;
            headerStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Medium;

            ICellStyle itemStyle = workbook.CreateCellStyle();
            itemStyle.Alignment = HorizontalAlignment.Center;
            itemStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Medium;
            itemStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Medium;
            itemStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Medium;
            itemStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Medium;

            ICellStyle summaryStyle = workbook.CreateCellStyle();
            summaryStyle.Alignment = HorizontalAlignment.Center;
            summaryStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;
            summaryStyle.FillPattern = FillPattern.SolidForeground;
            summaryStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Medium;
            summaryStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Medium;
            summaryStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Medium;
            summaryStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Medium;
            summaryStyle.SetFont(fontBold);

            //create header row 
            var headerRow = sheet.CreateRow(1);

            var headerRow_Cell1 = headerRow.CreateCell(1);
            headerRow_Cell1.SetCellValue("Header 1");
            headerRow_Cell1.CellStyle = headerStyle;

            var headerRow_Cell2 = headerRow.CreateCell(2);
            headerRow_Cell2.SetCellValue("Header 2");
            headerRow_Cell2.CellStyle = headerStyle;

            var headerRow_Cell3 = headerRow.CreateCell(3);
            headerRow_Cell3.SetCellValue("Header 3");
            headerRow_Cell3.CellStyle = headerStyle;

            var headerRow_Cell4 = headerRow.CreateCell(4);
            headerRow_Cell4.SetCellValue("Header 4");
            headerRow_Cell4.CellStyle = headerStyle;

            int rowNo = 0;

            //create items rows
            for (var i = 1; i < 11; i++)
            {
                rowNo = i + 1;

                var itemRow = sheet.CreateRow(rowNo);

                var itemRow_Cell1 = itemRow.CreateCell(1);
                itemRow_Cell1.SetCellValue("Row [" + i + "] Cell [1]");
                itemRow_Cell1.CellStyle = itemStyle;

                var itemRow_Cell2 = itemRow.CreateCell(2);
                itemRow_Cell2.SetCellValue("Row [" + i + "] Cell [2]");
                itemRow_Cell2.CellStyle = itemStyle;

                var itemRow_Cell3 = itemRow.CreateCell(3);
                itemRow_Cell3.SetCellValue("Row [" + i + "] Cell [3]");
                itemRow_Cell3.CellStyle = itemStyle;

                var itemRow_Cell4 = itemRow.CreateCell(4);
                itemRow_Cell4.SetCellValue(10 * rowNo);
                itemRow_Cell4.CellStyle = itemStyle;
            }

            //create summary row
            var summaryRow = sheet.CreateRow(++rowNo);

            var summaryRowCell = summaryRow.CreateCell(4);
            summaryRowCell.SetCellType(CellType.Formula);
            summaryRowCell.CellFormula = "sum(e3:e12)";
            summaryRowCell.CellStyle = summaryStyle;

            //define column width
            sheet.SetColumnWidth(1, 4000);
            sheet.SetColumnWidth(2, 4000);
            sheet.SetColumnWidth(3, 4000);
            sheet.SetColumnWidth(4, 4000);

            return workbook;
        }

        void GenerateOutput(HSSFWorkbook workbook)
        {
            var fileData = new MemoryStream();

            workbook.Write(fileData);

            using (var exportData = new MemoryStream())
            {
                workbook.Write(exportData);

                string saveAsFileName = string.Format("MembershipExport-{0:d}.xls", DateTime.Now).Replace("/", "-");

                Response.ContentType = "application/vnd.ms-excel";

                Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", saveAsFileName));

                Response.Clear();

                Response.BinaryWrite(exportData.GetBuffer());

                Response.End();
            }
        }
    }
}