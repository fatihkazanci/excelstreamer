# ExcelStreamer
ExcelStreamer is a .NET library that enables you to code reading and updating Microsoft Excel (.xlsx) file in a simpler way by making use of ClosedXML and ExcelDataReader Libraries.

# Installation
You can start with creating a ExcelStreamer object. 

```csharp
string excelPath = $"<your filePath address>";
using (ExcelStreamer excelStreamer = new(excelPath))
 {
    
 }
```
For with Create Excel
```csharp
string willCreateExcelPath = $"<your filePath address>";
using (ExcelStreamer excelStreamer = new(willCreateExcelPath, newWorksheetNames:"WorkSheetName1","WorksheetName2"))
 {
    
 }
```
Or Manual Example 1
```csharp
using (ExcelStreamer excelStreamer = new())
 {
    string excelPath = $"<Your Microsoft Excel File Path>";
    excelStreamer.CreateExcelFile(excelPath, "Page1");
 }
```
Or Manual Example 2
```csharp
using (ExcelStreamer excelStreamer = new())
 {
    string excelPath = "<Your Microsoft Excel File Path>";
    excelStreamer.SetFilePath(excelPath);
    string defaultWorkSheetName = "<WorkSheet Name>";
    excelStreamer.SetDefaultWorkSheet(defaultWorkSheetName);
 }
```
Or Manual Example 3
```csharp
using (ExcelStreamer excelStreamer = new())
 {
    string excelPath = "<Your Microsoft Excel File Path>";
    string defaultWorkSheetName = "<WorkSheet Name>";
    excelStreamer.SetDefault(excelPath, defaultWorkSheetName);
 }
```

Or you can inject dependency if you are going to use it in ASP.Net Core projects. 

```csharp
public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
{
   ....
   services.AddExcelStreamer();
   ....
}
```

If the file path is not specified, you must specify the file path using the SetFilePath method. 

# Attributes
**ExcelStreamerColumnLetter:** Determines which Microsoft Excel Column a Property in the created Model points to. 
 ```csharp
public class ExampleExcelSheetModel: ExcelStreamerWorkSheetObject
 {
        [ExcelStreamerColumnLetter("a")]
        public string Name { get; set; }
        [ExcelStreamerColumnLetter("b")]
        public string Surname { get; set; }
 }
 ```
 
 **ExcelStreamerSheetName:** Determines which Microsoft Excel Workspace a Property in the created Model points to. 
  ```csharp
 public class ExampleExcelModel : ExcelStreamerObject
 {
    [ExcelStreamerWorkSheetName("Yapılacaklar Listesi")]
    public List<ExampleExcelSheetModel> ToDoList { get; set; }
 }
 ```

# Base Models
**ExcelStreamerObject:** This abstract class is used for listing all Workspaces. 

```csharp
public class ExampleExcelModel : ExcelStreamerObject
{
    [ExcelStreamerWorkSheetName("Yapılacaklar Listesi")]
    public List<ExampleExcelSheetModel> ToDoList { get; set; }
}
```

**ExcelStreamerWorkSheetObject:** This abstract is used for creating a model of a Workspace. 

```csharp
public class ExampleExcelSheetModel: ExcelStreamerWorkSheetObject
 {
        [ExcelStreamerColumnLetter("a")]
        public string Name { get; set; }
        [ExcelStreamerColumnLetter("b")]
        public string Surname { get; set; }
 }
```
 

# Methods
![image](https://user-images.githubusercontent.com/33206545/162427262-197f2fbe-6aef-491e-9c2c-812a71b41979.png)

**SetDefaultFilePath:** Determines the file path that ExcelStreamer will read. 

```csharp
   excelStreamer.SetDefaultFilePath("<Your Microsoft Excel File Path>");
```
**SetDefaultWorkSheet:** Determines the Worksheet that ExcelStreamer will read. The **SetDefaultFilePath** Method must be set.

```csharp
   excelStreamer.SetDefaultWorkSheet("<Your Worksheet Name>");
```

**SetDefault:** Single method to set both File path and WorkSheet.
```csharp
   excelStreamer.SetDefault("<Your Microsoft Excel File Path>","<Your Worksheet Name>");
```

**CreateExcelFile:** Create New Excel File. When you create new Excel file you don't need to use **SetDefaultFilePath** again.
```csharp
public static void ExampleCreateExcel()
        {
            using (ExcelStreamer excelStreamer = new())
            {
                string excelPath = $"{AppDomain.CurrentDomain.BaseDirectory}CreatedExampleExcel.xlsx";
                excelStreamer.CreateExcelFile(excelPath, "Page1");
            }
        }
 ```

**WorkSheet:**  Brings the determined work page’s table data as a list.
 
 ```csharp
public class ExampleExcelSheetModel: ExcelStreamerSheetObject
 {
        [ExcelStreamerColumnLetter("a")]
        public string Name { get; set; }
        [ExcelStreamerColumnLetter("b")]
        public string Surname { get; set; }
 }
 
 public class ExampleProject 
 {
    public void ExampleMethod()
   {
       List<ExampleExcelWorkSheetModel> exampleList = excelStreamer.WorkSheet<ExampleExcelWorkSheetModel>("Page1", 1, 5, nameof(ExampleExcelWorkSheetModel.Name),      nameof(ExampleExcelWorkSheetModel.Surname));
      //OR
      List<ExampleExcelWorkSheetModel> exampleList = excelStreamer.WorkSheet<ExampleExcelWorkSheetModel>("Page1", 1, 5, "a", "b");
      //OR
      List<ExampleExcelWorkSheetModel> exampleListNoLimit = excelStreamer.WorkSheet<ExampleExcelWorkSheetModel>("Page1");
      //OR
      List<ExampleExcelWorkSheetModel> exampleListNoLimitOnlyColumns = excelStreamer.WorkSheet<ExampleExcelWorkSheetModel>("Page1","a","b");
      //OR
      List<ExampleExcelWorkSheetModel> exampleListNoLimitOnlyColumns2 = excelStreamer.WorkSheet<ExampleExcelWorkSheetModel>("Page1", nameof(ExampleExcelWorkSheetModel.Name));
   }
 }
 ```
 
 **WorkSheets:** Brings the data of the tables in all existing Workspaces in the Microsoft Excel file to the appropriate determined model.
 
  ```csharp
 public class ExampleExcelSheetModel: ExcelStreamerWorkSheetObject
 {
        [ExcelStreamerColumnLetter("a")]
        public string Name { get; set; }
        [ExcelStreamerColumnLetter("b")]
        public string Surname { get; set; }
 }
 
 public class ExampleExcelModel : ExcelStreamerObject
 {
    [ExcelStreamerWorkSheetName("Yapılacaklar Listesi")]
    public List<ExampleExcelWorkSheetModel> ToDoList { get; set; }
 }
 
 public class ExampleProject 
 {
    public void ExampleMethod()
    {
        ExampleExcelModel exampleLetterList = excelStreamer.WorkSheets<ExampleExcelModel>(1, 5, "a", "b");
        //OR
        ExampleExcelModel exampleLetterList = excelStreamer.WorkSheets<ExampleExcelModel>(1, 5, nameof(ExampleExcelWorkSheetModel.Name), nameof(ExampleExcelWorkSheetModel.Surname));
        //OR
        ExampleExcelModel exampleLetterListNoLimit = excelStreamer.WorkSheets<ExampleExcelModel>();
    }
 }
  ```
 
**Get:** Brings a table data in the determined Workspace in the desired object type.  
```csharp
 public class ExampleExcelSheetModel: ExcelStreamerWorkSheetObject
 {
        [ExcelStreamerColumnLetter("a")]
        public string Name { get; set; }
        [ExcelStreamerColumnLetter("b")]
        public string Surname { get; set; }
 }
 
 public class ExampleProject 
 {
    public void ExampleMethod()
    {
        ExampleExcelWorkSheetModel exampleSheetData = excelStreamer.Get<ExampleExcelWorkSheetModel>("Page1", 1, nameof(ExampleExcelWorkSheetModel.Name));
        //OR
        ExampleExcelWorkSheetModel exampleSheetData = excelStreamer.Get<ExampleExcelWorkSheetModel>("Page1", 1, "a","b");
        //OR
        string exampleSheetDataName = excelStreamer.Get<ExampleExcelWorkSheetModel, string>("Page1", nameof(ExampleExcelWorkSheetModel.Name), 1);
        //OR
        string exampleSheetDataSurname = excelStreamer.Get<string>("Page1", "b", 1);
    }
 }
 ```
 
**Update:** Updates the determined field in the Microsoft Excel file according to the given ExcelStreamerSheetObject object. 
 
 ```csharp
 public class ExampleExcelSheetModel: ExcelStreamerWorkSheetObject
 {
        [ExcelStreamerColumnLetter("a")]
        public string Name { get; set; }
        [ExcelStreamerColumnLetter("b")]
        public string Surname { get; set; }
 }
 
 public class ExampleProject 
 {
    public void ExampleMethod()
    {
         List<ExampleExcelWorkSheetModel> exampleList = excelStreamer.Sheet<ExampleExcelWorkSheetModel>("Page1", 1, 5, nameof(ExampleExcelWorkSheetModel.Name), nameof(ExampleExcelWorkSheetModel.Surname));
         exampleList[1].Name = "Osman";
         excelStreamer.Update(exampleList[1]);
         //OR
         excelStreamer.Update("Kazım", "Page1", "a", 1);
         excelStreamer.SaveChanges(); // This is required to save changes.
    }
 }
 ```
 
 **SaveChanges**: Method used to save changes made.
```csharp
   excelStreamer.Update("Kazım", "Page1", "a", 1);
   excelStreamer.SaveChanges();
```
 
**Count:** Brings the number of lines and columns of the table in the specified Workspace. 
 
 ```csharp
 public class ExampleProject 
 {
    public void ExampleMethod()
    {
        ExcelStreamerSheetCountResponse exampleSheetCount = excelStreamer.Count("Page1");
        //OR
        ExcelStreamerCountResponse allSheetCount = excelStreamer.Count();
    }
 }
```
 
 **UpdateWorkSheetName:** Changes the name of the desired Workspace in the Microsoft Excel file. 
 
  ```csharp
 public class ExampleProject 
 {
    public void ExampleMethod()
    {
         excelStreamer.UpdateWorkSheetName("Page1", "ExampleSheetName");
    }
 }
```
 
