// See https://aka.ms/new-console-template for more information

using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

FileStream file;
XSSFWorkbook wb;

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
wb = new XSSFWorkbook(file);
file.Close();

ISheet sheet = wb.GetSheetAt(0);
string buildingadd;
//int LastColumnNum = sheet.GetRow(8).Cells.Count;
for (int i = 1; i < sheet.LastRowNum; i++)
{
    buildingadd = GetItemRoad(sheet.GetRow(i).GetCell(9).ToString());
    sheet.GetRow(i).CreateCell(17).SetCellValue(buildingAddress.ContainsKey(buildingadd) ? buildingAddress[buildingadd] : "");
    //Console.WriteLine(sheet.GetRow(i).GetCell(9));
}


string GetItemRoad(string address)
{
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

