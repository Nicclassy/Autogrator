﻿using Excel = Microsoft.Office.Interop.Excel;

namespace Autogrator.OutlookAutomation;

public sealed class ExcelAllowedSenderList(Dictionary<string, string> folderNamesByAddress) : IAllowedSenderList {
    private const int StartingRowIndex = 2;

    public ExcelAllowedSenderList() : this(new()) { }

    public void Load(string filepath) {
        Excel.Application excel = new();
        Excel.Workbook workbook = excel.Workbooks.Open(filepath);
        Excel.Worksheet worksheet = workbook.Worksheets[1];
        
        int rowIndex = StartingRowIndex;
        Excel.Range cell = worksheet.Cells[rowIndex, 1];
        string cellValue = cell.Value;
        while (!string.IsNullOrEmpty(cellValue)) {
            (string folderName, List<string> emailAddresses) = ParseNonEmptyRowValues(worksheet, rowIndex);
            emailAddresses.ForEach(emailAddress => folderNamesByAddress[emailAddress] = folderName);

            cell = worksheet.Cells[rowIndex++, 1];
            cellValue = cell.Value;
        }
    }

    public bool IsAllowed(string emailAddress) => folderNamesByAddress.ContainsKey(emailAddress);

    public string GetSenderFolder(string emailAddress) => folderNamesByAddress[emailAddress];

    private static (string folderName, List<string> emailAddresses) ParseNonEmptyRowValues(Excel.Worksheet worksheet, int rowIndex) {
        List<string> emailAddresses = [];
        int columnIndex = 1;

        Excel.Range cell = worksheet.Cells[rowIndex, columnIndex++];
        string cellValue = cell.Value;
        while (!string.IsNullOrWhiteSpace(cellValue)) {
            emailAddresses.Add(cellValue);
            cell = worksheet.Cells[rowIndex, columnIndex++];
            cellValue = cell.Value;
        }

        string folderName = worksheet.Cells[rowIndex, 1].Value;
        return (folderName, emailAddresses);
    }
}