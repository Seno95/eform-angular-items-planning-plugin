/*
The MIT License (MIT)

Copyright (c) 2007 - 2021 Microting A/S

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

using System.Diagnostics;
using System.Reflection;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ItemsPlanning.Pn.Infrastructure.Models.Report;
using ItemsPlanning.Pn.Services.ItemsPlanningLocalizationService;
using Microting.eFormApi.BasePn.Abstractions;
using Microting.eFormApi.BasePn.Infrastructure.Models.API;
using Microting.ItemsPlanningBase.Infrastructure.Data;

namespace ItemsPlanning.Pn.Services.ExcelService;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Infrastructure.Consts;
using Infrastructure.Models.Import;
using Microsoft.Extensions.Logging;
using Microting.ItemsPlanningBase.Infrastructure.Enums;

public class PlanningExcelService(
    ILogger<PlanningExcelService> logger,
    IItemsPlanningLocalizationService localizationService,
    IEFormCoreService coreHelper,
    ItemsPlanningPnDbContext dbContext)
    : IPlanningExcelService
{
    private readonly IEFormCoreService _coreHelper = coreHelper;
    private readonly ItemsPlanningPnDbContext _dbContext = dbContext;

    public List<PlanningImportExcelModel> ParsePlanningImportFile(Stream excelStream)
    {
        try
        {
            var result = new List<PlanningImportExcelModel>();

            // Open the Excel document using OpenXML
            using (var spreadsheetDocument = SpreadsheetDocument.Open(excelStream, false))
            {
                var workbookPart = spreadsheetDocument.WorkbookPart;
                var sheets = workbookPart.Workbook.Sheets;

                // Get the specific sheet by index
                var sheet = sheets.ElementAt(PlanningImportExcelConsts.EformsWorksheet - 1) as Sheet;
                if (sheet == null)
                {
                    return result;
                }

                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                var rows = sheetData.Elements<Row>();

                foreach (var row in rows.Skip(1)) // Skip header
                {
                    var folders = new List<PlanningImportFolderModel>();
                    var item = new PlanningImportExcelModel
                    {
                        ExcelRow = (int)row.RowIndex.Value
                    };

                    // Folder 1 to 10
                    // Folder 1 to 10
                    AddFolderModel(workbookPart, folders, row, PlanningImportExcelConsts.Folder1Label, PlanningImportExcelConsts.Folder1Description, 1);
                    AddFolderModel(workbookPart, folders, row, PlanningImportExcelConsts.Folder2Label, PlanningImportExcelConsts.Folder2Description, 2);
                    AddFolderModel(workbookPart, folders, row, PlanningImportExcelConsts.Folder3Label, PlanningImportExcelConsts.Folder3Description, 3);
                    AddFolderModel(workbookPart, folders, row, PlanningImportExcelConsts.Folder4Label, PlanningImportExcelConsts.Folder4Description, 4);
                    AddFolderModel(workbookPart, folders, row, PlanningImportExcelConsts.Folder5Label, PlanningImportExcelConsts.Folder5Description, 5);
                    AddFolderModel(workbookPart, folders, row, PlanningImportExcelConsts.Folder6Label, PlanningImportExcelConsts.Folder6Description, 6);
                    AddFolderModel(workbookPart, folders, row, PlanningImportExcelConsts.Folder7Label, PlanningImportExcelConsts.Folder7Description, 7);
                    AddFolderModel(workbookPart, folders, row, PlanningImportExcelConsts.Folder8Label, PlanningImportExcelConsts.Folder8Description, 8);
                    AddFolderModel(workbookPart, folders, row, PlanningImportExcelConsts.Folder9Label, PlanningImportExcelConsts.Folder9Description, 9);
                    AddFolderModel(workbookPart, folders, row, PlanningImportExcelConsts.Folder10Label, PlanningImportExcelConsts.Folder10Description, 10);

                    item.Folders = folders;

                    // Planning info
                    item.PlanningName = GetCellValue(workbookPart, row, PlanningImportExcelConsts.PlanningItemNameCol);
                    // if (item.PlanningName == "")
                    // {
                    //     continue;
                    // }

                    if (item.Folders.Count == 0)
                    {
                        continue;
                    }

                    // Repeat information
                    var repeatEveryString = GetCellValue(workbookPart, row, PlanningImportExcelConsts.PlanningRepeatEveryCol);
                    if (int.TryParse(repeatEveryString, out var repeatEvery))
                    {
                        item.RepeatEvery = repeatEvery;
                    }

                    var repeatTypeString = GetCellValue(workbookPart, row, PlanningImportExcelConsts.PlanningRepeatTypeCol);
                    item.RepeatType = ParseRepeatType(repeatTypeString);

                    // Repeat until
                    var repeatUntilString = GetCellValue(workbookPart, row, PlanningImportExcelConsts.PlanningRepeatUntilCol);
                    if (DateTime.TryParseExact(repeatUntilString, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var repeatUntil))
                    {
                        item.RepeatUntil = repeatUntil;
                    }

                    // Day of week
                    var dayOfWeekString = GetCellValue(workbookPart, row, PlanningImportExcelConsts.PlanningDayOfWeekCol);
                    if (int.TryParse(dayOfWeekString, out var dayOfWeekValue))
                    {
                        item.DayOfWeek = (DayOfWeek)dayOfWeekValue;
                    }
                    else
                    {
                        item.DayOfWeek = ParseDayOfWeek(dayOfWeekString);
                    }

                    // Day of month
                    var dayOfMonthString = GetCellValue(workbookPart, row, PlanningImportExcelConsts.PlanningDayOfMonthCol);
                    if (int.TryParse(dayOfMonthString, out var dayOfMonth))
                    {
                        item.DayOfMonth = dayOfMonth;
                    }

                    // EForm name
                    item.EFormName = GetCellValue(workbookPart, row, PlanningImportExcelConsts.EformNameCol);

                    // EForm tags
                    item.Tags = new List<string>();
                    for (int i = PlanningImportExcelConsts.Tag1Col; i <= PlanningImportExcelConsts.Tag10Col; i++)
                    {
                        var tag = GetCellValue(workbookPart, row, i);
                        if (!string.IsNullOrEmpty(tag))
                        {
                            item.Tags.Add(tag);
                        }
                    }

                    result.Add(item);
                }
            }

            return result;
        }
        catch (Exception e)
        {
            logger.LogError(e, e.Message);
            throw;
        }
    }

// Helper Methods

    private void AddFolderModel(WorkbookPart workbookPart, List<PlanningImportFolderModel> folders, Row row, int labelIndex, int descIndex, int level)
    {
        var label = GetCellValue(workbookPart, row, labelIndex).Trim();
        var description = GetCellValue(workbookPart, row, descIndex).Trim();

        if (!string.IsNullOrEmpty(label))
        {
            folders.Add(new PlanningImportFolderModel
            {
                Label = label,
                Description = description,
                Level = level
            });
        }
    }

    private string GetCellValue(WorkbookPart workbookPart, Row row, int columnIndex)
    {
        // Get the column letter for the given columnIndex (e.g., A, B, C)
        var columnLetter = GetColumnLetter(columnIndex);

        // Create the cell reference (e.g., A1, B1, C1)
        var cellReference = columnLetter + row.RowIndex;

        // Find the cell with the matching CellReference
        var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference.Value == cellReference);

        if (cell == null || cell.CellValue == null)
        {
            return string.Empty; // Handle empty or missing cells
        }

        // If the cell is using a Shared String Table (most string values are stored here)
        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            var sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (sharedStringTablePart != null)
            {
                var sharedStringTable = sharedStringTablePart.SharedStringTable;
                return sharedStringTable.ElementAt(int.Parse(cell.CellValue.Text)).InnerText;
            }
        }

        // Return the cell value directly for non-shared strings
        return cell.CellValue.Text;
    }

    private string GetColumnLetter(int columnIndex)
    {
        string columnLetter = "";
        while (columnIndex > 0)
        {
            int modulo = (columnIndex - 1) % 26;
            columnLetter = Convert.ToChar(65 + modulo) + columnLetter;
            columnIndex = (columnIndex - modulo) / 26;
        }
        return columnLetter;
    }

    private RepeatType ParseRepeatType(string repeatTypeString)
    {
        return repeatTypeString.ToLower() switch
        {
            "month" => RepeatType.Month,
            "week" => RepeatType.Week,
            _ => RepeatType.Day,
        };
    }

    private DayOfWeek? ParseDayOfWeek(string dayOfWeekString)
    {
        return dayOfWeekString.ToLower() switch
        {
            "mon" => DayOfWeek.Monday,
            "tue" => DayOfWeek.Tuesday,
            "wed" => DayOfWeek.Wednesday,
            "thu" => DayOfWeek.Thursday,
            "fri" => DayOfWeek.Friday,
            "sat" => DayOfWeek.Saturday,
            "sun" => DayOfWeek.Sunday,
            _ => (DayOfWeek?)null,
        };
    }

    public async Task<OperationDataResult<Stream>> GenerateExcelDashboard(List<ReportEformModel> reportModel)
    {
        try
        {
            // Create a directory to store the result
            Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "results"));

            // Generate a timestamp for the result file
            var timeStamp = $"{DateTime.UtcNow:yyyyMMdd}_{DateTime.UtcNow:hhmmss}";

            // Create the path for the result Excel file
            var resultDocument = Path.Combine(Path.GetTempPath(), "results", $"{timeStamp}_.xlsx");

            // Create a new Excel file using OpenXml
            using (var spreadsheetDocument = SpreadsheetDocument.Create(resultDocument, SpreadsheetDocumentType.Workbook))
            {
                // Create a workbook and sheets
                var workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

                // Iterate through each eform report model to create a corresponding worksheet
                foreach (var eformModel in reportModel)
                {
                    if (eformModel.FromDate != null)
                    {
                        var sheetName = CleanSheetName(eformModel.TemplateName);
                        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                        worksheetPart.Worksheet = new Worksheet(new SheetData());
                        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                        // Add the new sheet to the workbook
                        var sheet = new Sheet
                        {
                            Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                            SheetId = (uint)(reportModel.IndexOf(eformModel) + 1),
                            Name = sheetName
                        };
                        sheets.Append(sheet);

                        // Add the headers row
                        var headerRow = new Row();
                        headerRow.Append(
                            ConstructCell(localizationService.GetString("Id"), CellValues.String),
                            ConstructCell(localizationService.GetString("CreatedAt"), CellValues.String),
                            ConstructCell(localizationService.GetString("DoneBy"), CellValues.String),
                            ConstructCell(localizationService.GetString("ItemName"), CellValues.String)
                        );

                        // Add dynamic headers for item fields
                        foreach (var itemHeader in eformModel.ItemHeaders)
                        {
                            headerRow.Append(ConstructCell(itemHeader.Value, CellValues.String));
                        }
                        sheetData.Append(headerRow);

                        // Add data rows
                        foreach (var dataModel in eformModel.Items)
                        {
                            var dataRow = new Row();
                            dataRow.Append(
                                ConstructCell(dataModel.MicrotingSdkCaseId.ToString(), CellValues.Number),
                                ConstructCell(dataModel.MicrotingSdkCaseDoneAt?.ToString("dd.MM.yyyy HH:mm:ss") ?? "", CellValues.String),
                                ConstructCell(dataModel.DoneBy, CellValues.String),
                                ConstructCell(dataModel.ItemName, CellValues.String)
                            );

                            // Add values for each field in the data model
                            foreach (var caseField in dataModel.CaseFields)
                            {
                                var cellValue = caseField.Value == "checked" ? "1" : caseField.Value == "unchecked" ? "0" : caseField.Value;

                                switch (caseField.Key)
                                {
                                    case "date":
                                        var date = DateTime.Parse(cellValue);
                                        dataRow.Append(ConstructCell(date.ToString("dd.MM.yyyy"), CellValues.Date));
                                        break;
                                    case "number":
                                        if (double.TryParse(cellValue, out var number))
                                        {
                                            dataRow.Append(ConstructCell(number.ToString(CultureInfo.InvariantCulture), CellValues.Number));
                                        }
                                        else
                                        {
                                            dataRow.Append(ConstructCell(cellValue, CellValues.String));
                                        }
                                        break;
                                    default:
                                        dataRow.Append(ConstructCell(cellValue, CellValues.String));
                                        break;
                                }
                            }
                            sheetData.Append(dataRow);
                        }
                    }
                }

                // Save the workbook
                workbookPart.Workbook.Save();
            }

            // Return the generated Excel file as a stream
            Stream result = File.Open(resultDocument, FileMode.Open, FileAccess.Read);
            return new OperationDataResult<Stream>(true, result);
        }
        catch (Exception e)
        {
            logger.LogError(e.Message);
            return new OperationDataResult<Stream>(false, localizationService.GetString("ErrorWhileCreatingExcelFile"));
        }
    }

// Helper method to create a cell
    private Cell ConstructCell(string value, CellValues dataType)
    {
        return new Cell
        {
            CellValue = new CellValue(value),
            DataType = new EnumValue<CellValues>(dataType)
        };
    }

// Helper method to clean and format sheet names
    private string CleanSheetName(string sheetName)
    {
        var invalidChars = new[] { ':', '\\', '/', '?', '*', '[', ']' };
        foreach (var invalidChar in invalidChars)
        {
            sheetName = sheetName.Replace(invalidChar.ToString(), string.Empty);
        }

        return sheetName.Length > 30 ? sheetName.Substring(0, 30) : sheetName;
    }

}