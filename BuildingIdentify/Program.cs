// See https://aka.ms/new-console-template for more information

using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Linq;
using System.Collections.Generic;

FileStream file;
XSSFWorkbook wb;

var buildingaddresspath = @"txt\楼宇地址.txt";
var filepath = @"企业注册数据导出.xlsx";

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
wb = new XSSFWorkbook(file);
file.Close();

ISheet sheet = wb.GetSheetAt(0);
string buildingadd;
int gzindex = (sheet.GetRow(0).Cells.Where(x => x.StringCellValue == "地址").First()).ColumnIndex;
//int LastColumnNum = sheet.GetRow(8).Cells.Count;
for (int i = 1; i < sheet.LastRowNum; i++)
{
    //buildingadd = GetItemRoad(sheet.GetRow(i).GetCell(9).ToString());

    //第9列是地址，此处可能会变
    buildingadd = GetItemAddress(sheet.GetRow(i).GetCell(gzindex).ToString());
    if (buildingadd != null)
        //最后一列，此处也可能会变
        sheet.GetRow(i).CreateCell(18).SetCellValue(buildingAddress.ContainsKey(buildingadd) ? buildingAddress[buildingadd] : "");
    //Console.WriteLine(sheet.GetRow(i).GetCell(9));
}



string GetItemAddress(string address)
{
    if (address.IndexOf("市北区") == -1)
        return null;
    int qu = address.IndexOf("区");
    int hao = address.IndexOf("号", qu);
    return hao < qu ? "" : address.Substring(qu + 1, hao - qu).Replace(" ", "");
}
//foreach (var ba in buildingAddress)
//Console.WriteLine(ba);

using (FileStream fs = new FileStream(@"E:\test\企业楼宇地址.xlsx", FileMode.Create, FileAccess.Write))
{
    fs.Seek(0, SeekOrigin.Begin);
    wb.Write(fs);
}

wb = null;

