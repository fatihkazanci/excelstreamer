# ExcelStreamer
ExcelStreamer ClosedXML ve ExcelDataReader Kütüphanelerinden faydalanarak Microsoft Excel (.xlsx) dosyasını okuma ve güncelleme işlemlerini daha sade bir şekilde kodlamanızı sağlayan .NET kütüphanesidir.

# Installation
Yeni bir ExcelStreamer nesnesi oluşturarak başlayabilirsiniz

```csharp
string excelPath = $"<your filePath address>";
using (ExcelStreamer excelStreamer = new(excelPath))
 {
    
 }
```

veyahut eğer ASP.Net Core projelerinde kullanacaksanız dependency injection yapabilirsiniz.

```csharp
public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
{
   ....
   services.AddExcelStreamer();
   ....
}
```

Eğer dosya yolu belirtilmemiş ise **SetFilePath** methodunu kullanarak dosya yolu belirtmeniz gerekmektedir.

# Attributes
**ExcelStreamerColumnLetter:** Oluşturulan Modeldeki bir Property'in hangi Microsoft Excel Kolonuna işaret ettiğini belirler.
 ```csharp
public class ExampleExcelSheetModel: ExcelStreamerWorkSheetObject
 {
        [ExcelStreamerColumnLetter("a")]
        public string Name { get; set; }
        [ExcelStreamerColumnLetter("b")]
        public string Surname { get; set; }
 }
 ```
 
 **ExcelStreamerSheetName:** Oluşturulan Modeldeki bir Property'in hangi Microsoft Excel Çalışma Alanına işaret ettiğini belirler.
  ```csharp
 public class ExampleExcelModel : ExcelStreamerObject
 {
    [ExcelStreamerSheetName("Yapılacaklar Listesi")]
    public List<ExampleExcelSheetModel> ToDoList { get; set; }
 }
 ```

# Base Models
**ExcelStreamerObject:** Bütün Çalışma Alanlarını listeleyebilmek için bu abstract class'a ihtiyaç duyulur.

```csharp
public class ExampleExcelModel : ExcelStreamerObject
{
    [ExcelStreamerSheetName("Yapılacaklar Listesi")]
    public List<ExampleExcelSheetModel> ToDoList { get; set; }
}
```

**ExcelStreamerWorkSheetObject:** Bir Çalışma alanının modelini oluşturabilmekiçin bu abstract class'a ihtiyaç duyulur.

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


**SetFilePath(string filePath):** ExcelStreamer'in okuyacağı dosya yolunu belirler.

```csharp
   excelStreamer.SetFilePath("<Your Microsoft Excel File Path>");
```
**WorkSheet:** Belirlenen Çalışma Sayfasının tablo verilerini liste biçiminde getirir.
 
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
       List<ExampleExcelWorkSheetModel> exampleList = excelStreamer.Sheet<ExampleExcelWorkSheetModel>("Page1", 1, 5, nameof(ExampleExcelSheetModel.Name),      nameof(ExampleExcelSheetModel.Surname));
      //OR
      List<ExampleExcelWorkSheetModel> exampleList = excelStreamer.Sheet<ExampleExcelWorkSheetModel>("Page1", 1, 5, "a", "b");
      //OR
      List<ExampleExcelWorkSheetModel> exampleListNoLimit = excelStreamer.WorkSheet<ExampleExcelWorkSheetModel>("Page1");
   }
 }
 ```
 
 **WorkSheets:** Microsoft Excel dosyasındaki mevcut tüm Çalışma alanlarınıdaki tabloları verilerini uygun belirlenen modele getirir.
 
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
    [ExcelStreamerSheetName("Yapılacaklar Listesi")]
    public List<ExampleExcelSheetModel> ToDoList { get; set; }
 }
 
 public class ExampleProject 
 {
    public void ExampleMethod()
    {
        ExampleExcelModel exampleLetterList = excelStreamer.WorkSheets<ExampleExcelModel>(1, 5, "a", "b");
        //OR
        ExampleExcelModel exampleLetterList = excelStreamer.WorkSheets<ExampleExcelModel>(1, 5, nameof(ExampleExcelSheetModel.Name), nameof(ExampleExcelSheetModel.Surname));
        //OR
        ExampleExcelModel exampleLetterListNoLimit = excelStreamer.WorkSheets<ExampleExcelModel>();
    }
 }
  ```
 
**Get:** Belirlenen Çalışma alanınındaki bir tablo verisini istenilen nesne türünde getirir.
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
        ExampleExcelSheetModel exampleSheetData = excelStreamer.Get<ExampleExcelSheetModel>("Page1", 1, nameof(ExampleExcelSheetModel.Name));
        //OR
        ExampleExcelSheetModel exampleSheetData = excelStreamer.Get<ExampleExcelSheetModel>("Page1", 1, "a","b");
        //OR
        string exampleSheetDataName = excelStreamer.Get<ExampleExcelSheetModel, string>("Page1", nameof(ExampleExcelSheetModel.Name), 1);
        //OR
        string exampleSheetDataSurname = excelStreamer.Get<string>("Page1", "b", 1);
    }
 }
 ```
 
**Update(ExcelStreamerSheetObject updateObject):** Verilen ExcelStreamerSheetObject objesine göre Microsoft Excel dosyasındaki belirtilen alanı günceller.
 
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
         List<ExampleExcelSheetModel> exampleList = excelStreamer.Sheet<ExampleExcelSheetModel>("Page1", 1, 5, nameof(ExampleExcelSheetModel.Name), nameof(ExampleExcelSheetModel.Surname));
         exampleList[1].Name = "Osman";
         excelStreamer.Update(exampleList[1]);
    }
 }
 ```
 
**Update(object newValue, string worksheetName, string columnLetterName, int row):** Belirtilen Çalışma Alanına göre Microsoft Excel dosyasındaki alanı günceller.
 
```csharp
 public class ExampleProject 
 {
    public void ExampleMethod()
    {
       excelStreamer.Update("Kazım", "Page1", "a", 1);
    }
 }
```
 
**Count(string worksheetName):** Belirtilen Çalışma alanındaki tablonun satır ve sutun sayısını getirir.
 
 ```csharp
 public class ExampleProject 
 {
    public void ExampleMethod()
    {
        ExcelStreamerSheetCountResponse exampleSheetCount = excelStreamer.Count("Page1");
    }
 }
```
 
 **Count():** Bütün Çalışma alanlarındaki tabloların satır ve sutun sayısını getirir.
 
 ```csharp
 public class ExampleProject 
 {
    public void ExampleMethod()
    {
         ExcelStreamerCountResponse allSheetCount = excelStreamer.Count();
    }
 }
```
 
 **UpdateWorkSheetName(string currentSheetName, string newSheetName):** Microsoft Excel dosyasındaki istenilen bir Çalışma Alanı adını değiştirir.
 
  ```csharp
 public class ExampleProject 
 {
    public void ExampleMethod()
    {
         excelStreamer.UpdateSheetName("Page1", "ExampleSheetName");
    }
 }
```
 
