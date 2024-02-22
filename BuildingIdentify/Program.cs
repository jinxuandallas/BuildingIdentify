// See https://aka.ms/new-console-template for more information

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

FileStream file;
HSSFWorkbook wb;

var buildingaddresspath = @"txt\楼宇地址.txt";
var filepath = @"..\企业注册数据导出.xlsx";

Dictionary<string, string> buildingAddress = ReadBuildingAddress(buildingaddresspath);

Dictionary<string, string> ReadBuildingAddress(string path)
{
    Dictionary<string, string> result = new Dictionary<string, string>();

    StreamReader sr = new StreamReader(path);
    string? line = sr.ReadLine();
    while (line != null)
    {
        string[] item = line.Split('，');
        result.Add(item[1].Trim(), item[0].Trim());
        line = sr.ReadLine();
    }
    sr.Close();

    return result;

}

file = new FileStream(filepath, FileMode.Open, FileAccess.Read); 
wb = new HSSFWorkbook(file);
file.Close();

ISheet sheet= wb.GetSheetAt(0);

for (int i = 1; i < sheet.LastRowNum; i++)
{
    Console.WriteLine(sheet.GetRow(i).GetCell(10));
}


//foreach (var ba in buildingAddress)
//Console.WriteLine(ba);

