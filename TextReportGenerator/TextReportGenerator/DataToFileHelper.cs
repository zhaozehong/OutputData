using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;

namespace TextReportGenerator
{
  public abstract class DataToFile
  {
    public Boolean Save(String filePath, Dictionary<String, object> headerDictionary, object[,] dataArray)
    {
      TabelCellData[,] cellHeaderArray = null;
      if (headerDictionary != null && headerDictionary.Any())
      {
        cellHeaderArray = new TabelCellData[headerDictionary.Count, 2];
        for (var iHeader = 0; iHeader < headerDictionary.Count; iHeader++)
        {
          var excelHeaderIndex = iHeader + 1;
          cellHeaderArray[iHeader, 0] = new TabelCellData(headerDictionary.ElementAt(iHeader).Key, true) { ForegroundColor = Colors.Blue };
          cellHeaderArray[iHeader, 1] = new TabelCellData(headerDictionary.ElementAt(iHeader).Value) { ColumnSpan = Math.Max(2, dataArray.GetLength(1) - 1) };
        }
      }

      TabelCellData[,] cellDataArray = null;
      if (dataArray != null && dataArray.Length > 0)
      {
        cellDataArray = new TabelCellData[dataArray.GetLength(0), dataArray.GetLength(1)];
        for (int i = 0; i < dataArray.GetLength(0); i++)
        {
          for (int j = 0; j < dataArray.GetLength(1); j++)
          {
            cellDataArray[i, j] = new TabelCellData(dataArray[i, j]);
          }
        }
      }

      return Save(filePath, cellHeaderArray, cellDataArray);
    }

    public abstract Boolean Save(String filePath, TabelCellData[,] cellHeaderArray, TabelCellData[,] cellDataArray);

    public Dictionary<String, object> HeaderDictionary { get; set; }
    public object[,] DataArray { get; set; }
  }
  public class DataToExcel : DataToFile
  {
    public DataToExcel(int? decimals = null)
    {
      this.Decimals = decimals;
    }

    public static Boolean Save(String filePath, Dictionary<String, object> headerDictionary, List<ExcelSheetData> sheetDatas, int? decimals = null)
    {
      try
      {
        if (String.IsNullOrWhiteSpace(filePath))
          return false;

        using (var excel = new ExcelHelper())
        {
          excel.Initial(filePath, decimals);
          for (var iSheet = 0; iSheet < sheetDatas.Count; iSheet++)
          {
            if (!String.IsNullOrWhiteSpace(sheetDatas[iSheet].SheetName))
            {
              if (iSheet == 0)
                excel.SetSheetName(sheetDatas[iSheet].SheetName);
              else
                excel.AddSheet(sheetDatas[iSheet].SheetName);
            }

            var excelSheetIndex = iSheet + 1;
            // Header
            var formats = new List<ExcelRangeFormat>();
            if (headerDictionary != null && headerDictionary.Any())
            {
              var headerValues = new object[headerDictionary.Count, 2];
              for (var iHeader = 0; iHeader < headerDictionary.Count; iHeader++)
              {
                var excelHeaderIndex = iHeader + 1;
                headerValues[iHeader, 0] = headerDictionary.ElementAt(iHeader).Key;
                headerValues[iHeader, 1] = headerDictionary.ElementAt(iHeader).Value;
                formats.Add(new ExcelRangeFormat() { RowStartIndex = excelHeaderIndex, ColumnStartIndex = 1, Foreground = System.Windows.Media.Colors.Blue, Bold = true });
                formats.Add(new ExcelRangeFormat() { RowStartIndex = excelHeaderIndex, ColumnStartIndex = 2, ColumnCount = Math.Max(2, sheetDatas[iSheet].DataArray.GetLength(1) - 1) });
              }
              excel.WriteCells(excelSheetIndex, 1, 1, headerValues);
              excel.FormatCells(excelSheetIndex, formats);
            }

            // Content
            excel.WriteCells(excelSheetIndex, 1, 1, sheetDatas[iSheet].DataArray);
            excel.FormatCells(excelSheetIndex, sheetDatas[iSheet].RangeFormats);
          }

          excel.SaveWorkBook();
          excel.SetActiveSheet(1);
        }
        return true;
      }
      catch (Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
        return false;
      }
    }

    public override Boolean Save(String filePath, TabelCellData[,] cellHeaderArray, TabelCellData[,] cellDataArray)
    {
      try
      {
        if (String.IsNullOrWhiteSpace(filePath))
          return false;

        using (var excel = new ExcelHelper())
        {
          excel.Initial(filePath, this.Decimals);

          if (!OutputArray(excel, cellHeaderArray, 1, 1, 1))
            return false;

          if (!OutputArray(excel, cellDataArray, 1, cellHeaderArray.GetLength(0) + 2, 1))
            return false;

          excel.SaveWorkBook();
          excel.SetActiveSheet(1);
        }
        return true;
      }
      catch (Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
        return false;
      }
    }

    private Boolean OutputArray(ExcelHelper excel, TabelCellData[,] cellArray, int startSheet, int startRow, int startColumn)
    {
      try
      {
        if (excel == null)
          return false;
        if (cellArray == null)
          return true;

        if (startSheet < 1)
          startSheet = 1;
        if (startRow < 1)
          startRow = 1;
        if (startColumn < 1)
          startColumn = 1;

        var formats = new List<ExcelRangeFormat>();
        var rowCount = cellArray.GetLength(0);
        var columnCount = cellArray.GetLength(1);
        var dataArray = new object[rowCount, columnCount];
        for (var i = 0; i < rowCount; i++)
        {
          for (var j = 0; j < columnCount; j++)
          {
            dataArray[i, j] = cellArray[i, j].CellValue;
            formats.Add(cellArray[i, j].GetExcelRangeFormat(startRow + i, startColumn + j));
          }
        }
        excel.WriteCells(startSheet, startRow, startColumn, dataArray);
        excel.FormatCells(startSheet, formats.Where(p => p != null).ToList());
        return true;
      }
      catch (Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
        return false;
      }
    }

    public int? Decimals { get; private set; }
  }

  public class DataToPdf : DataToFile
  {
    public override Boolean Save(String filePath, TabelCellData[,] cellHeaderArray, TabelCellData[,] cellDataArray)
    {
      return true;
    }
  }

  public class TabelCellData
  {
    public TabelCellData(object cellValue, Boolean isBold = false)
    {
      this.CellValue = cellValue;
      this.IsBold = isBold;
      this.IsNumberFormat = false;
      this.ForegroundColor = Colors.Black;
      this.RowSpan = 1;
      this.ColumnSpan = 1;
    }

    public ExcelRangeFormat GetExcelRangeFormat(int iRow, int iColumn)
    {
      try
      {
        ExcelRangeFormat format = null;
        if (this.IsBold)
        {
          if (format == null)
            format = new ExcelRangeFormat() { RowStartIndex = iRow, ColumnStartIndex = iColumn };
          format.Bold = this.IsBold;
        }
        if (this.ForegroundColor != Colors.Black)
        {
          if (format == null)
            format = new ExcelRangeFormat() { RowStartIndex = iRow, ColumnStartIndex = iColumn };
          format.Foreground = this.ForegroundColor;
        }
        if (this.RowSpan != 1)
        {
          if (format == null)
            format = new ExcelRangeFormat() { RowStartIndex = iRow, ColumnStartIndex = iColumn };
          format.RowCount = RowSpan;
        }
        if (this.ColumnSpan != 1)
        {
          if (format == null)
            format = new ExcelRangeFormat() { RowStartIndex = iRow, ColumnStartIndex = iColumn };
          format.ColumnCount = this.ColumnSpan;
        }
        if (this.IsNumberFormat && this.CellValue is double)
        {
          if (format == null)
            format = new ExcelRangeFormat() { RowStartIndex = iRow, ColumnStartIndex = iColumn };
          format.NeedNumberFormat = true;
        }
        return format;
      }
      catch
      {
        return null;
      }
    }

    public object CellValue { get; private set; }
    public Boolean IsBold { get; set; }
    public Boolean IsNumberFormat { get; set; }
    public Color ForegroundColor { get; set; }
    public int RowSpan { get; set; }
    public int ColumnSpan { get; set; }
  }
}
