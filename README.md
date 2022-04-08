# ExcelStreamer.NET
ExcelStreamer ClosedXML ve ExcelDataReader Kütüphanelerinden faydalanarak Microsoft Excel (.xlsx) dosyasını okuma ve güncelleme işlemlerini daha sade bir şekilde kodlamanızı sağlayan kütüphanedir.

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

Eğer filePath yolu belirtilmemiş ise **SetFilePath** methodunu kullanarak dosya yolu belirtmeniz gerekmektedir.

# Methods

![image](https://user-images.githubusercontent.com/33206545/162419217-146890a5-6228-4117-b797-704617aa636c.png)


**SetFilePath(string filePath):** ExcelStreamer'in okuyacağı dosya yolunu belirler.

```csharp
   excelStreamer.SetFilePath("<Your Microsoft Excel File Path>");
```
**Sheet<T>(string worksheetName, int startRow, int endRow, params string[] columnLetterNames):** Belirlenen Çalışma Sayfasının tablo verilerini liste biçiminde getirir.
 
 ```csharp
 public class ExampleExcelModel : ExcelStreamerObject
 {
    [ExcelStreamerSheetName("Yapılacaklar Listesi")]
    public List<ExampleExcelSheetModel> ToDoList { get; set; }
 }
 
 List<ExampleExcelSheetModel> exampleList = excelStreamer.Sheet<ExampleExcelSheetModel>("Page1", 1, 5, nameof(ExampleExcelSheetModel.Name), nameof(ExampleExcelSheetModel.Surname));
 //OR
 List<ExampleExcelSheetModel> exampleList = excelStreamer.Sheet<ExampleExcelSheetModel>("Page1", 1, 5, "a", "b");
 ```
