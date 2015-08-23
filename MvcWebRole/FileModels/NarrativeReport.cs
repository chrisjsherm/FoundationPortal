using FundEntities;
using MvcWebRole.Extensions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Web.Configuration;

namespace MvcWebRole.FileModels
{
    public class NarrativeReport : Report
    {
        private const int NUM_COLUMNS = 5;
        private int Row { get; set; }
        private IEnumerable<Fund> Funds { get; set; }

        public NarrativeReport(IEnumerable<Area> areas, IEnumerable<Fund> funds)
        {
            Funds = funds;

            ExcelPackage package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Funding Request Report");

            #region Table Labels
            Row++;
            sheet = WriteTableLabels(sheet);
            #endregion

            #region Area Data
            foreach (Area area in areas)
            {
                sheet = WriteAreaData(sheet, area);
            }
            #endregion

            sheet = PerformFinalFormatting(sheet);

            this.BinaryData = package.GetAsByteArray();
            this.FileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            this.FileName = "FoundationPortal_" + System.DateTime.Now.ToShortDateString() + ".xlsx";
        }

        private ExcelWorksheet WriteTableLabels(ExcelWorksheet sheet)
        {
            int column = 0;
            ExcelRange range_labels = sheet.Cells[Row, 1, Row, NUM_COLUMNS];
            range_labels.Style.Font.SetFromFont(new Font("Calibri", 11, FontStyle.Bold));
            range_labels.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;
            range_labels.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            range_labels.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(200, 200, 200));
            range_labels.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;

            sheet.Cells[Row, ++column].Value = DataAnnotationsHelper.GetPropertyName<Fund>(f => f.Number);
            sheet.Cells[Row, ++column].Value = DataAnnotationsHelper.GetPropertyName<Fund>(f => f.Title);
            sheet.Cells[Row, ++column].Value = DataAnnotationsHelper.GetPropertyName<Fund>(f => f.Description);
            sheet.Cells[Row, ++column].Value =
                DataAnnotationsHelper.GetPropertyName<Fund>(f => f.BudgetAdjustmentNote);
            sheet.Cells[Row, ++column].Value = "Variance";

            return sheet;
        }

        private ExcelWorksheet WriteAreaData(ExcelWorksheet sheet, Area area)
        {
            int column = 0;

            #region Area Name
            Row++;
            ExcelRange range_areaName = sheet.Cells[Row, 1, Row, NUM_COLUMNS];
            range_areaName.Merge = true;
            range_areaName.Style.Font.SetFromFont(new Font("Calibri", 11, FontStyle.Italic));
            range_areaName.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;
            range_areaName.Value = area.Name;
            #endregion

            #region Area Funds
            IEnumerable<Fund> areaFunds = Funds.Where(f => f.AreaId == area.Id);

            if (areaFunds.Count() == 0)
            {
                Row++;
                ExcelRange range_areaData = sheet.Cells[Row, 1, Row, NUM_COLUMNS];
                range_areaData.Merge = true;
                range_areaData.Style.Font.SetFromFont(new Font("Calibri", 11, FontStyle.Regular));
                range_areaData.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;
                range_areaData.Value = "No records";
            }
            else
            {
                foreach (Fund fund in areaFunds)
                {
                    Row++;
                    column = 0;
                    sheet.Cells[Row, ++column].Value = fund.Number;
                    sheet.Cells[Row, ++column].Value = fund.Title;
                    sheet.Cells[Row, ++column].Value = fund.Description;
                    sheet.Cells[Row, ++column].Value = fund.BudgetAdjustmentNote;
                    sheet.Cells[Row, ++column].Value = fund.BudgetAdjustment * -1;
                }
            }
            #endregion

            return sheet;
        }

        private ExcelWorksheet PerformFinalFormatting(ExcelWorksheet sheet)
        {
            //Header
            sheet.HeaderFooter.FirstHeader.LeftAlignedText = "VIRGINIA TECH FOUNDATION INC.\n"
                + "UNRESTRICTED BUDGET\n" +
                "FY " + WebConfigurationManager.AppSettings["FiscalYear"].ToString();

            //Footer
            sheet.HeaderFooter.FirstFooter.CenteredText = System.DateTime.Now.ToShortDateString() +
                " Narrative of VT Foundation Funding Request FY " +
                WebConfigurationManager.AppSettings["FiscalYear"].ToString();

            //Printing
            sheet.PrinterSettings.Orientation = eOrientation.Landscape;
            sheet.PrinterSettings.FitToPage = true;
            sheet.PrinterSettings.FitToWidth = 1;
            sheet.PrinterSettings.FitToHeight = 0;
            ExcelRange range_numberFormatting =
                sheet.Cells[1, NUM_COLUMNS, 100, NUM_COLUMNS];

            //Cell styling
            range_numberFormatting.Style.Numberformat.Format = "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)";

            sheet.Cells.AutoFitColumns();

            return sheet;
        }
    }
}