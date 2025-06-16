# ExcelNPOILib
�o�O�@�Ӧۦ漶�g���B�zEXCEL�ɮר禡
1.�D�n�Ω�N���󶰦X�ץX�� Excel �ɮסA�åi�ۭq���W�ٻP���ǡC�ثeEXCEL�ɪ�CELL�榡�O�{�����w�]��u�P
   ���ت��r���j�p�A���ӥi�H�X�R�ѥ~���ǤJ�C
2.�ѪREXCEL�ɡA�NEXCEL���ഫ�����󶰦X�C
3.��ɼ��g���禡�ɡA�D�n�O�m�߬M�g���[���A�ҥH�ϥΤF�ݩʬM�g�]Attribute Mapping�^�Ӧ۰ʹ��� Excel ���A
  �åB²�Ʋ��ͬJ�w�榡EXCEL�ɻP�ѪREXCEL�ɪ��{���X�C
����� .NET Standard 2.1 �P NPOI �M�� �P [NPOI](https://github.com/tonyqus/npoi) �M�� Excel �ɮ׾ާ@�禡�w�A�䴩 xls/xlsx �榡�A��K�N���󶰦X�ץX�� Excel �ɮסA�åi�ۭq���W�ٻP���ǡC

## �S��

- �䴩 .NET Standard 2.1�A�󥭥x�i��
- �H�ݩʬM�g�]Attribute Mapping�^�۰ʹ��� Excel ���
- �ۭq���W�ٻP����
- �۰ʲ��ͼ��D�C�P��ƦC
- �䴩 xls (Excel 97-2003) �� xlsx (Excel 2007+) �榡
- �������۰ʼe�׽վ�
- �i�ۭq���Y�P���e�˦�

## �w��

�Х��w�� [NPOI](https://www.nuget.org/packages/NPOI/)

## �ϥΤ覡

1. �w�q��Ƽҫ��A�åH `PropertyColumnNameAttribute` �P `PropertySeqAttribute` �е��ݩʡG
```csharp
using ExcelNPOILib;
public class Person { [PropertySeq(1)] [PropertyColumnName("�m�W")] public string Name { get; set; }
[PropertySeq(2)]
[PropertyColumnName("�~��")]
public int Age { get; set; }

[PropertySeq(3)]
[PropertyColumnName("�ͤ�")]
public DateTime Birthday { get; set; }
}
```
	
2. �ץX Excel�G
```csharp
var people = new List<Person> { new Person { Name = "���p��", Age = 30, Birthday = new DateTime(1993, 1, 1) }, new Person { Name = "���p��", Age = 25, Birthday = new DateTime(1998, 5, 20) } };
var service = new ExcelNPOIService(); 
service.CreateExcel("people.xlsx", "�H���M��", people);
```


## �ݩʼе�����

- `PropertySeqAttribute(int seq)`  
  ���w���b Excel �������ǡA�Ʀr�V�p�V�e���C
- `PropertyColumnNameAttribute(string columnName)`  
  ���w���b Excel ������ܦW�١C

## �M�׵��c

- `ExcelNPOIService`�G�D�n�� Excel �ѪR�P�ץX�A�����O
- `PropertySeqAttribute`�G�ݩʶ��Ǽе�
- `PropertyColumnNameAttribute`�G�ݩ����W�ټе�