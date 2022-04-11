using ClosedXML.Excel;
using ExcelDataReader;
using ExcelStreamerLibrary.Attributes;
using ExcelStreamerLibrary.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

/*
*	ExcelStreamer v1.0
*	Updated On: 09/04/2022 (dd/MM/yyyy)
*	By Fatih KAZANCI
*   Licenced By GNU General Public License v3.0
*   https://github.com/fatihkazanci/excelstreamer/blob/main/LICENSE
*/

namespace ExcelStreamerLibrary
{
    public class ExcelStreamer : IDisposable
    {
        private XLWorkbook _xLWorkbook;
        private string _filePath;

        public ExcelStreamer(string excelFilePath)
        {
            _xLWorkbook = new XLWorkbook(excelFilePath);
            _filePath = excelFilePath;
        }

        public ExcelStreamer()
        {

        }

        public void SetFilePath(string filePath)
        {
            this._filePath = filePath;
        }
        public List<T> WorkSheet<T>(string worksheetName, int startRow, int endRow, params string[] columnLetterNames) where T : ExcelStreamerWorkSheetObject
        {
            return (List<T>)WorkSheet(typeof(List<T>), worksheetName, startRow, endRow, columnLetterNames);
        }
        public List<T> WorkSheet<T>(string worksheetName) where T : ExcelStreamerWorkSheetObject
        {
            return (List<T>)WorkSheet(typeof(List<T>), worksheetName);
        }
        public List<T> WorkSheet<T>(string worksheetName, params string[] columnLetterNames) where T : ExcelStreamerWorkSheetObject
        {
            return (List<T>)WorkSheet(typeof(List<T>), worksheetName, columnLetterNames);
        }
        private object WorkSheet(Type objectType, string worksheetName, int startRow, int endRow, params string[] columnLetterNames)
        {
            startRow = startRow <= 0 ? 1 : startRow;
            object newObjectList = Activator.CreateInstance(objectType);
            for (int c = startRow; c <= endRow; c++)
            {
                ExcelStreamerWorkSheetObject newObject = Activator.CreateInstance(objectType.GenericTypeArguments[0]) as ExcelStreamerWorkSheetObject;
                newObject._RowNumber = c;
                newObject._SheetName = worksheetName;
                newObjectList.GetType().GetMethod("Add").Invoke(newObjectList, new[] { newObject });
            }
            using (FileStream stream = File.Open(_filePath, FileMode.Open, FileAccess.Read))
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        while (reader.Read()) { }
                    } while (reader.NextResult());

                    DataSet result = reader.AsDataSet();
                    DataTable tables = result.Tables[worksheetName];
                    string[] alp = ExcelStreamerExtensions.Generate().Take(tables.Columns.Count).ToArray();
                    foreach (string letterName in columnLetterNames)
                    {
                        string letterNameUpper = letterName.ToUpper();
                        for (int i = startRow; i <= endRow; i++)
                        {
                            int columnStartIndex = Array.IndexOf(alp, letterNameUpper);
                            object currentObject = null;
                            int newObjectListCount = (int)newObjectList.GetType().GetTypeInfo().GetProperty("Count").GetValue(newObjectList);

                            for (int n = 0; n < newObjectListCount; n++)
                            {
                                object currentNObject = newObjectList.GetType().GetTypeInfo().GetProperty("Item").GetValue(newObjectList, new object[] { n });
                                int currentRowNumber = (int)currentNObject.GetType().GetProperty(nameof(ExcelStreamerWorkSheetObject._RowNumber)).GetValue(currentNObject);
                                if (currentRowNumber == i)
                                {
                                    currentObject = (ExcelStreamerWorkSheetObject)currentNObject;
                                    break;
                                }
                            }
                            if (columnStartIndex == -1)
                            {
                                string newLetterName = currentObject?.GetType().GetTypeInfo().GetProperty(letterName)?.GetCustomAttribute<ExcelStreamerColumnLetter>()?.ColumnLetterName.ToUpper();
                                if (!string.IsNullOrEmpty(newLetterName))
                                {
                                    columnStartIndex = Array.IndexOf(alp, newLetterName);
                                    letterNameUpper = newLetterName;
                                }
                                else
                                {
                                    return null;
                                }
                            }

                            PropertyInfo[] properties = currentObject?.GetType()?.GetTypeInfo()?.GetProperties();
                            if (properties is not null)
                            {
                                foreach (PropertyInfo item in properties)
                                {
                                    if (item.GetCustomAttribute<ExcelStreamerColumnLetter>()?.ColumnLetterName?.ToUpper() == letterNameUpper)
                                    {
                                        object currentItem = ((DataRow)tables.Rows[i - 1]).ItemArray[columnStartIndex];
                                        item.SetValue(currentObject, currentItem);
                                    }
                                }
                            }
                        }
                    }
                }
                return newObjectList;
            }
        }
        private object WorkSheet(Type objectType, string worksheetName)
        {
            object newObjectList = Activator.CreateInstance(objectType);
            using (FileStream stream = File.Open(_filePath, FileMode.Open, FileAccess.Read))
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        while (reader.Read()) { }
                    } while (reader.NextResult());

                    DataSet result = reader.AsDataSet();
                    DataTable tables = result.Tables[worksheetName];
                    for (int c = 1; c <= tables.Rows.Count; c++)
                    {
                        ExcelStreamerWorkSheetObject newObject = Activator.CreateInstance(objectType.GenericTypeArguments[0]) as ExcelStreamerWorkSheetObject;
                        newObject._RowNumber = c;
                        newObject._SheetName = worksheetName;
                        newObjectList.GetType().GetMethod("Add").Invoke(newObjectList, new[] { newObject });
                    }
                    string[] alp = ExcelStreamerExtensions.Generate().Take(tables.Columns.Count).ToArray();
                    for (int c = 1; c < tables.Columns.Count; c++)
                    {
                        string letterNameUpper = alp[c - 1];
                        for (int i = 1; i < tables.Rows.Count; i++)
                        {
                            object currentObject = null;
                            int newObjectListCount = (int)newObjectList.GetType().GetTypeInfo().GetProperty("Count").GetValue(newObjectList);

                            for (int n = 0; n < newObjectListCount; n++)
                            {
                                object currentNObject = newObjectList.GetType().GetTypeInfo().GetProperty("Item").GetValue(newObjectList, new object[] { n });
                                int currentRowNumber = (int)currentNObject.GetType().GetProperty(nameof(ExcelStreamerWorkSheetObject._RowNumber)).GetValue(currentNObject);
                                if (currentRowNumber == i)
                                {
                                    currentObject = (ExcelStreamerWorkSheetObject)currentNObject;
                                    break;
                                }
                            }
                            PropertyInfo[] properties = currentObject?.GetType()?.GetTypeInfo()?.GetProperties();
                            if (properties is not null)
                            {
                                foreach (PropertyInfo item in properties)
                                {
                                    if (item.GetCustomAttribute<ExcelStreamerColumnLetter>()?.ColumnLetterName?.ToUpper() == letterNameUpper)
                                    {
                                        object currentItem = ((DataRow)tables.Rows[i]).ItemArray[c];
                                        item.SetValue(currentObject, currentItem);
                                    }
                                }
                            }
                        }
                    }
                }
                return newObjectList;
            }
        }
        private object WorkSheet(Type objectType, string worksheetName, params string[] columnLetterNames)
        {
            object newObjectList = Activator.CreateInstance(objectType);
            using (FileStream stream = File.Open(_filePath, FileMode.Open, FileAccess.Read))
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        while (reader.Read()) { }
                    } while (reader.NextResult());

                    DataSet result = reader.AsDataSet();
                    DataTable tables = result.Tables[worksheetName];
                    for (int c = 1; c <= tables.Rows.Count; c++)
                    {
                        ExcelStreamerWorkSheetObject newObject = Activator.CreateInstance(objectType.GenericTypeArguments[0]) as ExcelStreamerWorkSheetObject;
                        newObject._RowNumber = c;
                        newObject._SheetName = worksheetName;
                        newObjectList.GetType().GetMethod("Add").Invoke(newObjectList, new[] { newObject });
                    }
                    string[] alp = ExcelStreamerExtensions.Generate().Take(tables.Columns.Count).ToArray();
                    for (int c = 1; c < tables.Columns.Count; c++)
                    {
                        foreach (string letterName in columnLetterNames)
                        {
                            string letterNameUpper = letterName.ToUpper();
                            int columnStartIndex = Array.IndexOf(alp, letterNameUpper);

                            for (int i = 1; i < tables.Rows.Count; i++)
                            {
                                object currentObject = null;
                                int newObjectListCount = (int)newObjectList.GetType().GetTypeInfo().GetProperty("Count").GetValue(newObjectList);

                                for (int n = 0; n < newObjectListCount; n++)
                                {
                                    object currentNObject = newObjectList.GetType().GetTypeInfo().GetProperty("Item").GetValue(newObjectList, new object[] { n });
                                    int currentRowNumber = (int)currentNObject.GetType().GetProperty(nameof(ExcelStreamerWorkSheetObject._RowNumber)).GetValue(currentNObject);
                                    if (currentRowNumber == i)
                                    {
                                        currentObject = (ExcelStreamerWorkSheetObject)currentNObject;
                                        break;
                                    }
                                }

                                if (columnStartIndex == -1)
                                {
                                    string newLetterName = currentObject?.GetType().GetTypeInfo().GetProperty(letterName)?.GetCustomAttribute<ExcelStreamerColumnLetter>()?.ColumnLetterName.ToUpper();
                                    if (!string.IsNullOrEmpty(newLetterName))
                                    {
                                        columnStartIndex = Array.IndexOf(alp, newLetterName);
                                        letterNameUpper = newLetterName;
                                    }
                                    else
                                    {
                                        return null;
                                    }
                                }

                                PropertyInfo[] properties = currentObject?.GetType()?.GetTypeInfo()?.GetProperties();
                                if (properties is not null)
                                {
                                    foreach (PropertyInfo item in properties)
                                    {
                                        if (item.GetCustomAttribute<ExcelStreamerColumnLetter>()?.ColumnLetterName?.ToUpper() == letterNameUpper)
                                        {
                                            object currentItem = ((DataRow)tables.Rows[i]).ItemArray[c];
                                            item.SetValue(currentObject, currentItem);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                return newObjectList;
            }
        }
        public T WorkSheets<T>(int startRow, int endRow, params string[] columnLetterNames) where T : ExcelStreamerObject
        {
            Type tType = typeof(T);
            T newTObject = Activator.CreateInstance<T>();
            IXLWorksheets xLWorksheets = _xLWorkbook.Worksheets;

            foreach (IXLWorksheet item in xLWorksheets)
            {
                PropertyInfo sheetProperty = tType.GetProperties().Where(i => i.GetCustomAttribute<ExcelStreamerWorkSheetName>()?.SheetName == item.Name).FirstOrDefault();
                if (sheetProperty is not null)
                {
                    object propertyObjectList = WorkSheet(sheetProperty.PropertyType, item.Name, startRow, endRow, columnLetterNames);
                    sheetProperty.SetValue(newTObject, propertyObjectList);
                }
            }
            return newTObject;
        }
        public T WorkSheets<T>() where T : ExcelStreamerObject
        {
            Type tType = typeof(T);
            T newTObject = Activator.CreateInstance<T>();
            IXLWorksheets xLWorksheets = _xLWorkbook.Worksheets;

            foreach (IXLWorksheet item in xLWorksheets)
            {
                PropertyInfo sheetProperty = tType.GetProperties().Where(i => i.GetCustomAttribute<ExcelStreamerWorkSheetName>()?.SheetName == item.Name).FirstOrDefault();
                if (sheetProperty is not null)
                {
                    object propertyObjectList = WorkSheet(sheetProperty.PropertyType, item.Name);
                    sheetProperty.SetValue(newTObject, propertyObjectList);
                }
            }
            return newTObject;
        }
        public T Get<T>(string worksheetName, int row, params string[] columnLetterNames) where T : ExcelStreamerWorkSheetObject
        {
            row = row == 0 ? 1 : row;
            T newObject = Activator.CreateInstance<T>();
            IXLWorksheet xLWorksheet = _xLWorkbook.Worksheet(worksheetName);
            int tableTotalColumn = 0;
            using (FileStream stream = File.Open(_filePath, FileMode.Open, FileAccess.Read))
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                {
                    tableTotalColumn = reader.FieldCount;
                }
            }
            string[] alp = ExcelStreamerExtensions.Generate().Take(tableTotalColumn).ToArray();
            foreach (var letterName in columnLetterNames)
            {
                string letterNameUpper = letterName.ToUpper();
                int columnStartIndex = Array.IndexOf(alp, letterNameUpper);
                if (columnStartIndex == -1)
                {
                    string newLetterName = newObject.GetType().GetTypeInfo().GetProperty(letterName)?.GetCustomAttribute<ExcelStreamerColumnLetter>()?.ColumnLetterName.ToUpper();
                    if (!string.IsNullOrEmpty(newLetterName))
                    {
                        columnStartIndex = Array.IndexOf(alp, newLetterName);
                        letterNameUpper = alp[columnStartIndex];
                    }
                    else
                    {
                        return null;
                    }
                }
                object currentItem = xLWorksheet.Cell($"{letterNameUpper}{row}")?.Value;
                newObject._RowNumber = row;
                PropertyInfo[] properties = newObject.GetType().GetTypeInfo().GetProperties();
                foreach (PropertyInfo item in properties)
                {
                    if (item.GetCustomAttribute<ExcelStreamerColumnLetter>()?.ColumnLetterName?.ToUpper() == letterNameUpper)
                    {
                        item.SetValue(newObject, currentItem);
                    }
                }
            }
            return newObject;
        }
        public T Get<ExcelStreamerSheet, T>(string worksheetName, string columnLetterName, int row) where ExcelStreamerSheet : ExcelStreamerWorkSheetObject
        {
            row = row == 0 ? 1 : row;
            ExcelStreamerSheet newSheetObject = Activator.CreateInstance<ExcelStreamerSheet>();
            IXLWorksheet xLWorksheet = _xLWorkbook.Worksheet(worksheetName);
            int tableTotalColumn = 0;
            using (FileStream stream = File.Open(_filePath, FileMode.Open, FileAccess.Read))
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                {
                    tableTotalColumn = reader.FieldCount;
                }
            }
            string[] alp = ExcelStreamerExtensions.Generate().Take(tableTotalColumn).ToArray();
            string letterNameUpper = columnLetterName.ToUpper();
            int columnStartIndex = Array.IndexOf(alp, letterNameUpper);
            if (columnStartIndex == -1)
            {
                string newLetterName = newSheetObject.GetType().GetTypeInfo().GetProperty(columnLetterName)?.GetCustomAttribute<ExcelStreamerColumnLetter>()?.ColumnLetterName.ToUpper();
                if (!string.IsNullOrEmpty(newLetterName))
                {
                    columnStartIndex = Array.IndexOf(alp, newLetterName);
                    letterNameUpper = alp[columnStartIndex];
                }
                else
                {
                    return default(T);
                }
            }
            object newObject = xLWorksheet.Cell($"{letterNameUpper}{row}").Value;
            return (T)Convert.ChangeType(newObject, typeof(T));
        }
        public T Get<T>(string worksheetName, string columnLetterName, int row)
        {
            row = row == 0 ? 1 : row;
            IXLWorksheet xLWorksheet = _xLWorkbook.Worksheet(worksheetName);
            object newObject = xLWorksheet.Cell($"{columnLetterName.ToUpper()}{row}").Value;
            return (T)Convert.ChangeType(newObject, typeof(T));
        }
        public ExcelStreamerResponse Update(ExcelStreamerWorkSheetObject updateObject)
        {
            ExcelStreamerResponse excelStreamerResponse = new();
            try
            {
                IXLWorksheet xLWorksheet = _xLWorkbook.Worksheet(updateObject._SheetName);
                PropertyInfo[] properties = updateObject.GetType().GetTypeInfo().GetProperties();
                foreach (PropertyInfo property in properties)
                {
                    string letterName = property.GetCustomAttribute<ExcelStreamerColumnLetter>()?.ColumnLetterName;
                    object propertyValue = property.GetValue(updateObject);
                    if (!string.IsNullOrEmpty(letterName) && propertyValue is not null)
                    {
                        letterName = letterName.ToUpper();
                        xLWorksheet.Cell($"{letterName}{updateObject._RowNumber}").Value = propertyValue;
                    }
                }
                _xLWorkbook.SaveAs(_filePath);
                excelStreamerResponse.Result = updateObject;
                return excelStreamerResponse;
            }
            catch (Exception ex)
            {
                excelStreamerResponse.Error(ex);
            }
            return excelStreamerResponse;
        }
        public ExcelStreamerResponse Update(object newValue, string worksheetName, string columnLetterName, int row)
        {
            ExcelStreamerResponse excelStreamerResponse = new();
            try
            {
                IXLWorksheet xLWorksheet = _xLWorkbook.Worksheet(worksheetName);
                xLWorksheet.Cell($"{columnLetterName}{row}").Value = newValue;
                excelStreamerResponse.Result = newValue;
                _xLWorkbook.SaveAs(_filePath);
            }
            catch (Exception ex)
            {
                excelStreamerResponse.Error(ex);
            }
            return excelStreamerResponse;
        }
        public ExcelStreamerWorkSheetCountResponse Count(string worksheetName)
        {
            ExcelStreamerWorkSheetCountResponse excelStreamerSheetCountResponse = new();

            using (FileStream stream = File.Open(_filePath, FileMode.Open, FileAccess.Read))
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                {
                    excelStreamerSheetCountResponse.ColumnCount = reader.FieldCount;
                    excelStreamerSheetCountResponse.RowCount = reader.RowCount;
                    excelStreamerSheetCountResponse.SheetName = worksheetName;
                }
            }
            return excelStreamerSheetCountResponse;
        }
        public ExcelStreamerCountResponse Count()
        {
            ExcelStreamerCountResponse excelStreamerCountResponse = new();
            IXLWorksheets xLWorksheets = _xLWorkbook.Worksheets;
            foreach (IXLWorksheet item in xLWorksheets)
            {
                excelStreamerCountResponse.Sheets.Add(Count(item.Name));
            }
            excelStreamerCountResponse.TotalSheet = xLWorksheets.Count;
            return excelStreamerCountResponse;
        }
        public ExcelStreamerResponse UpdateWorkSheetName(string currentSheetName, string newSheetName)
        {
            ExcelStreamerResponse excelStreamerResponse = new();
            try
            {
                IXLWorksheet xLWorksheet = _xLWorkbook.Worksheet(currentSheetName);
                xLWorksheet.Name = newSheetName;
                _xLWorkbook.SaveAs(_filePath);
            }
            catch (Exception ex)
            {
                excelStreamerResponse.Error(ex);
            }
            return excelStreamerResponse;
        }
        public void Dispose()
        {
            _xLWorkbook.Dispose();
        }
    }
}
