//  Pseudeo code
using Database; //  To use dbMethod method
using ExcelIO;  //  To use XSSWorkBook class
using System.UI;    //  To use Checkbox class

class Program
{
    private DataSet dataset = indb_flm051_05.ds_Data;    //  Source Data
    private CheckBox dwQ_m051_export_no_check = new CheckBox(); //  UI component

    public void Main()
    {
        make_worksheet3();
    }

    protected XSSFWorkbook make_worksheet3()
    {
        IRevtriever dataRetriever = new GeneralDataRetriever(dataset, dwQ_m051_export_no_check.Checked);
        ExcelExporter2 exporter = new ExcelExporter2(dataRetriever);

        exporter.AddColumn("製卡批號", "m051_export_no");
        exporter.AddColumn("安全時數證明編號", "m051_m05licno");
        exporter.AddColumn("中文姓名", "m05_cname");
        exporter.AddColumn("身分證／證件號碼", "m05_idno");
        exporter.AddColumn("QRcode資訊", "m051_qrcode");

        exporter.AddColumn("開班編號", "m05_remark");
        exporter.AddColumn("宣導機構", "m02_m01code");
        exporter.AddColumn("開班年度", "m02_yyy");
        exporter.AddColumn("宣導日期", "m02_sdate");
        exporter.AddColumn("場次名稱", "m02_name");

        exporter.AddColumn("時數", "m02_hrs");
        exporter.AddColumn("宣導活動項目", "m02_content1");
        exporter.AddColumn("體驗活動項目", "m02_content2");

        return exporter.GenerateWorkBook(dataset.Tables[0].Rows.Count);
    }

}

//--------------------------------------------------------------------------------------

//  建立欄位資訊
public class Column
{
    public string HeaderName { get; set; }
    public string ID { get; set; }

    public Column(string name, string id)
    {
        HeaderName = name;
        ID = id;
    }
}

public interface IRevtriever
{
    public string GetValue(string columnID, int rowIndex);
}

//  建立 General Retriever
public class GeneralDataRetriever : IRevtriever
{
    public bool _exportNoIsNull = false;

    private DataSet _dataset;

    public GeneralDataRetriever(DataSet dataset, bool exportNoIsNull)
    {
        _dataset = dataset;
        _exportNoIsNull = exportNoIsNull;
    }

    public string GetValue(string columnID, int rowIndex)
    {
        switch(columnID)
        {
            case "m051_export_no":
                return GetExportNo();
            case "m051_m05licno":
                return GetFromDataSet(columnID, rowIndex);
            case "m05_cname":
                return GetFromDataSet(columnID, rowIndex);
            case "m05_idno":
                return GetFromDataSet(columnID, rowIndex);
            case "m051_qrcode":
                return GetQRCode(columnID, rowIndex);
            case "m05_remark":
                return GetClassNo(columnID, rowIndex);
            case "m02_m01code":
                return GetPromotionalOrganization();
            case "m02_yyy":
                return OpeningYear(columnID, rowIndex);
            case "m02_sdate":
                return GetPromotionalDate(columnID, rowIndex);
            case "m02_name":
                return GetSessionName(columnID, rowIndex);
            case "m02_hrs":
                return GetHours(columnID, rowIndex);
            case "m02_content1":
                return GetContent1(columnID, rowIndex);
            case "m02_content2":
                return GetContent2(columnID, rowIndex);
            default:
                throw new ArgumentException($"Unknown columnID: {columnID}");

        }
    }

    private string GetFromDataSet(string columnID, int rowIndex)
    {
        return _dataset.Tables[0].Rows[rowIndex][columnID].ToString() ?? string.Empty;
    }

    private string GetExportNo()
    {
        string ls_addsql = "";
        if(_exportNoIsNull)
            ls_addsql += " and m051_export_no is null";
        else
            ls_addsql += " and m051_export_no is not null";

        string ls_sql_m051_export_no = @"select (select convert(varchar,convert(int,substring(convert(varchar, getdate(), 112),1,4))-1911)+substring(convert(varchar, getdate(), 112),5,8)
                                        + REPLICATE('0', 2 - LEN((select count(m051_export_no) + 1 from flm051 where substring(m051_export_no, 1, 8) = convert(varchar, getdate(), 112)))) +RTRIM(CAST((select count(m051_export_no) + 1 from flm051 where substring(m051_export_no, 2, 8) = convert(varchar, getdate(), 112)) AS CHAR)) )
                                        from flm051
                                        where (1 = 1)" + ls_addsql + get_addsql() + " group by m051_export_no";
        DbMethods.uf_ExecSQL(_SqlQuery, ref result);
        return result;
    }

    private string GetQRCode(string columnID, int rowIndex)
    {
        string ls_QRcode = "";
        string ls_sql = "SELECT dbo.uf_get_qrcode('" + Security.uf_SQL(_ds.Tables[0].Rows[rowIndex]["m051_m05licno"].ToString()) 
                                    + "', " + Security.uf_SQL(_ds.Tables[0].Rows[rowIndex]["m051_seq"].ToString()) 
                                    + ", '" + Security.uf_SQL(_ds.Tables[0].Rows[rowIndex]["m05_idno"].ToString()) + "')";
        DbMethods.uf_ExecSQL(ls_sql, ref ls_QRcode);
        return ls_QRcode;
    }

    private string GetClassNo(string columnID, int rowIndex)
    {
        string ls_remark = "";
        string ls_sql = @"select m051_remark from flm051 where m051_m05licno = '" 
                        + Security.uf_SQL(_ds.Tables[0].Rows[rowIndex]["m051_m05licno"])
                        + "' and m051_seq = '" + Security.uf_SQL(_ds.Tables[0].Rows[rowIndex]["m051_seq"]) + "'";
        DbMethods.uf_ExecSQL(ls_sql, ref ls_remark);
        
        string pattern = @"開班編號:(\d+)";
        Match match = Regex.Match(ls_remark, pattern);

        string result = "";
        if (match.Success)
        {
            // 獲取表達式中的捕獲組
            result = match.Groups[1].Value;
            Console.WriteLine("開班編號: " + result);
        }
        else
            Console.WriteLine("找不到開班編號");

        return result;
    }

    private string GetPromotionalOrganization()
    {
        string ls_promotionalOrganization = "";
        string ls_sql_m02_yyy = @"SELECT d.m01_cname FROM flm051 AS a JOIN flm02 AS b ON 
               SUBSTRING(a.m051_remark, 
               CHARINDEX('開班編號:', a.m051_remark) + 5,CHARINDEX('學員編號:', a.m051_remark) - CHARINDEX('開班編號:', a.m051_remark) - 5) = b.m02_classno
                JOIN flm01 AS d ON b.m02_m01code = d.m01_code
               WHERE
                    a.m051_m05licno = '" + Security.uf_SQL( _ds.Tables[0].Rows[rowIndex]["m051_m05licno"]) + "'AND a.m051_seq = '" + Security.uf_SQL(_ds.Tables[0].Rows[rowIndex]["m051_seq"]) + "'; ";

        DbMethods.uf_ExecSQL(ls_sql_m02_yyy, ref ls_promotionalOrganization);
        return ls_promotionalOrganization;
    }

    private string OpeningYear(string columnID, int rowIndex)
    {
        string ls_openingYear = "";
        //  先取得開班編號
        string ls_classno = _classnoRetriever.DataRetrieve(columnID, rowIndex);
        string ls_sql_m02_yyy = @"select m02_yyy + '年' from flm02 where m02_classno = '" + ls_classno + "' ";
        DbMethods.uf_ExecSQL(ls_sql_m02_yyy, ref ls_openingYear);
        return ls_openingYear;
    }

    private string GetPromotionalDate(string columnID, int rowIndex)
    {
        string ls_promotionDate = "";
        //  先取得開班編號
        string ls_classno = GetClassNo(columnID, rowIndex);
        
        string ls_sql_m02_sdate = @"select FORMAT(m02_sdate, 'yyyy-MM-dd') + '~' + FORMAT(m02_edate, 'yyyy-MM-dd')from flm02 where m02_classno = '" + ls_classno + "' ";
        DbMethods.uf_ExecSQL(ls_sql_m02_sdate, ref ls_promotionDate);
        return ls_promotionDate;
    }

    private string GetSessionName(string columnID, int rowIndex)
    {
        string ls_sessionName = "";
        //  先取得開班編號
        string ls_classno = GetClassNo(columnID, rowIndex);
        
        string ls_sql_m02_name = @"select m02_name from flm02 where m02_classno = '" + ls_classno + "' ";
        DbMethods.uf_ExecSQL(ls_sql_m02_name, ref ls_sessionName);
        return ls_sessionName;
    }

    private string GetHours(string columnID, int rowIndex)
    {
        string ls_hours = "";
        //  先取得開班編號
        string ls_classno = _classnoRetriever.DataRetrieve(columnID, rowIndex);
        
        string ls_sql_m02_hrs = @"select CONVERT(VARCHAR(50), m02_hrs) + '小時' from flm02 where m02_classno = '" + ls_classno + "' ";

        DbMethods.uf_ExecSQL(ls_sql_m02_hrs, ref ls_hours);
        return ls_hours;
    }

    private string GetContent1(string columnID, int rowIndex)
    {
        string ls_contents1 = "";
        //  先取得開班編號
        string ls_classno = GetClassNo(columnID, rowIndex);
        
        string ls_sql_m02_content1 = @"select m02_content1 from flm02 where m02_classno = '" + ls_classno + "' ";
        DbMethods.uf_ExecSQL(ls_sql_m02_content1, ref ls_contents1);
        return ls_contents1;        
    }

    private string GetContent2(string columnID, int rowIndex)
    {
        string ls_contents2 = "";
        //  先取得開班編號
        string ls_classno = GetClassNo(columnID, rowIndex);
        
        string ls_sql_m02_content2 = @"select m02_content2 from flm02 where m02_classno = '" + ls_classno + "' ";
        DbMethods.uf_ExecSQL(ls_sql_m02_content2, ref ls_contents2);
        return ls_contents2;
    }

    private string get_addsql()
    {
        return "A series of SQL commands need to be appended.";
    }

}


//  建立 Excel 匯出類別 (只使用一個 retriever，在GetValue()內判斷 欄位需要用哪一個 Get Data 的 method)
class ExcelExporter2
{
    private List<Column> _columns;
    private XSSFWorkbook _workbook;
    private GeneralDataRetriever _retriever;

    //  Constructor
    public ExcelExporter2(IRevtriever retriever)
    {
        _workbook = new XSSFWorkbook();
        _columns = new List<Column>();
        _retriever = retriever;
    }

    public void AddColumn(string headerName, string columnID)
    {
        _columns.Add(new Column(headerName, columnID));
    }

    public XSSFWorkbook GenerateWorkBook(int rowCount)
    {
        ISheet sheet = _workbook.CreateSheet("匯出卡片編號");
        CreateHeaderRow(sheet);

        //  Traverse each row except the header row.
        for(int rowIndex = 0; rowIndex < rowCount; rowIndex++)
        {
            IRow dataRow;
            dataRow = sheet.CreateRow(rowIndex + 1);
            
            //  Assign Data to each column in the row
            for(int idx = 0; idx < _columns.Count; idx++)
            {
                ICell cell = dataRow.CreateCell(_retriever.GetValue(_columns[idx].HeaderName, _columns[idx].ID));
            }
        }
    }

    //  建立首欄標題名稱
    private void CreateHeaderRow(ISheet sheet)
    {
        IRow row = sheet.CreateRow(0);
        for (int idx=0; idx < _columns.Count; idx++)
        {
            ICell cell = row.CreateCell(idx);
            cell.SetCellValue(_columns[idx].HeaderName);
            sheet.AutoSizeColumn(idx);
        }
    }
}