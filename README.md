# ExcelHelper
這是一個自行撰寫的處理EXCEL檔案函式  
1. 主要用於將物件集合匯出為 Excel 檔案，並可自訂欄位名稱與順序。目前EXCEL檔的CELL格式是程式中預設格線與  
   內建的字型大小，未來可以擴充由外部傳入。  
2. 解析EXCEL檔，將EXCEL檔轉換為物件集合。  
3. 當時撰寫此函式時，主要是練習映射的觀念，所以使用了屬性映射（Attribute Mapping）來自動對應 Excel 欄位，  
   並且簡化產生既定格式EXCEL檔與解析EXCEL檔的程式碼。  
它基於 .NET Standard 2.1 與 NPOI 套件 與 [NPOI](https://github.com/tonyqus/npoi) 套件的 Excel 檔案操作函式庫，支援 xls/xlsx 格式，方便將物件集合匯出為 Excel 檔案，並可自訂欄位名稱與順序。  

## 特色

- 支援 .NET Standard 2.1，跨平台可用
- 以屬性映射（Attribute Mapping）自動對應 Excel 欄位
- 自訂欄位名稱與順序
- 自動產生標題列與資料列
- 支援 xls (Excel 97-2003) 及 xlsx (Excel 2007+) 格式
- 內建欄位自動寬度調整
- 可自訂表頭與內容樣式

## 安裝

請先安裝 [NPOI](https://www.nuget.org/packages/NPOI/)

## 使用方式

1. 定義資料模型，並以 `PropertyColumnNameAttribute` 與 `PropertySeqAttribute` 標註屬性：
```csharp
using ExcelNPOILib;
public class Person { [PropertySeq(1)] [PropertyColumnName("姓名")] public string Name { get; set; }
[PropertySeq(2)]
[PropertyColumnName("年齡")]
public int Age { get; set; }

[PropertySeq(3)]
[PropertyColumnName("生日")]
public DateTime Birthday { get; set; }
}
```
	
2. 匯出 Excel：
```csharp
var people = new List<Person> { new Person { Name = "王小明", Age = 30, Birthday = new DateTime(1993, 1, 1) }, new Person { Name = "李小華", Age = 25, Birthday = new DateTime(1998, 5, 20) } };
var service = new ExcelNPOIService(); 
service.CreateExcel("people.xlsx", "人員清單", people);
```


## 屬性標註說明

- `PropertySeqAttribute(int seq)`  
  指定欄位在 Excel 中的順序，數字越小越前面。
- `PropertyColumnNameAttribute(string columnName)`  
  指定欄位在 Excel 中的顯示名稱。

## 專案結構

- `ExcelNPOIService`：主要的 Excel 解析與匯出服務類別
- `PropertySeqAttribute`：屬性順序標註
- `PropertyColumnNameAttribute`：屬性欄位名稱標註
