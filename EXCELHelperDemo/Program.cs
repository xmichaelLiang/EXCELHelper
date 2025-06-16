using EXCELHelper;
using EXCELHelperDemo;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Reflection;
Console.WriteLine("Excel NPOI 函式庫示範程式");
Console.WriteLine("=========================");

// 準備示範資料
var students = new List<Student>
            {
                new Student { Id = 1, Name = "王大明", Age = 20, BirthDate = new DateTime(2004, 5, 15), Score = 85.5 },
                new Student { Id = 2, Name = "李小華", Age = 22, BirthDate = new DateTime(2002, 8, 21), Score = 92.0 },
                new Student { Id = 3, Name = "張三", Age = 19, BirthDate = new DateTime(2005, 3, 10), Score = 78.5 },
                new Student { Id = 4, Name = "陳小美", Age = 21, BirthDate = new DateTime(2003, 11, 5), Score = 95.0 },
                new Student { Id = 5, Name = "林志明", Age = 23, BirthDate = new DateTime(2001, 7, 30), Score = 88.0 }
            };

// 建立 ExcelNPOIService 實例
var excelService = new ExcelNPOIService();

// 指定 Excel 檔案路徑
string currentDir = Directory.GetCurrentDirectory();
string fileName = "學生資料.xlsx";
string filePath = Path.Combine(currentDir, fileName);

Console.WriteLine($"正在將資料匯出至 {filePath}...");

try
{
    // 將資料匯出為 Excel 檔案
    excelService.CreateExcel(filePath, "學生資料表", students);
    Console.WriteLine("匯出成功！");

    Console.WriteLine("\n正在讀取 Excel 檔案...");

    // 讀取 Excel 檔案並解析為物件列表
    var importedStudents = excelService.ParseExcelToList<Student>(filePath, "學生資料表");

    // 顯示解析後的資料
    Console.WriteLine("\n解析結果:");
    Console.WriteLine("ID\t姓名\t年齡\t出生日期\t分數");
    Console.WriteLine("----------------------------------------");
    foreach (var student in importedStudents)
    {
        Console.WriteLine($"{student.Id}\t{student.Name}\t{student.Age}\t{student.BirthDate:yyyy/MM/dd}\t{student.Score}");
    }
}
catch (Exception ex)
{
    Console.WriteLine($"發生錯誤: {ex.Message}");
    if (ex.InnerException != null)
        Console.WriteLine($"內部錯誤: {ex.InnerException.Message}");
}

Console.WriteLine("\n按任意鍵結束...");
Console.ReadKey();
        