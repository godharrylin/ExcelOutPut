//  Pseudeo code
using Database; //  To use dbMethod method
using ExcelIO;  //  To use XSSWorkBook class
using System.UI;    //  To use Checkbox class

class Program
{
    private DataSet ds = indb_flm051_05.ds_Data;    //  Source Data
    private CheckBox dwQ_m051_export_no_check = new CheckBox(); //  UI component

    public void Main()
    {
        make_worksheet3();
    } 

    protected XSSFWorkbook make_worksheet2()
    {
        DataSet ds = indb_flm051_05.ds_Data;
        string ls_addsql = "";
        if (dwQ_m051_export_no_check.Checked == true)
            ls_addsql += " and m051_export_no is null";
        else
            ls_addsql += " and m051_export_no is not null";


        string ls_sql_m051_export_no = @"select (select convert(varchar,convert(int,substring(convert(varchar, getdate(), 112),1,4))-1911)+substring(convert(varchar, getdate(), 112),5,8)
                                                            + REPLICATE('0', 2 - LEN((select count(m051_export_no) + 1 from flm051 where substring(m051_export_no, 1, 8) = convert(varchar, getdate(), 112)))) +RTRIM(CAST((select count(m051_export_no) + 1 from flm051 where substring(m051_export_no, 2, 8) = convert(varchar, getdate(), 112)) AS CHAR)) )
                                                            from flm051
                                                            where (1 = 1)" + ls_addsql + get_addsql() + " group by m051_export_no";
        ExcelExporter exporter = new ExcelExporter();
        exporter.AddColumn("製卡批號", "m051_export_no", new SqlRetriever(ls_sql_m051_export_no));
        exporter.AddColumn("安全時數證明編號", "m051_m05licno", new DataSetRetriever(ds));
        exporter.AddColumn("中文姓名", "m05_cname", new DataSetRetriever(ds));
        exporter.AddColumn("身分證／證件號碼", "m05_idno", new DataSetRetriever(ds));
        exporter.AddColumn("QRcode資訊", "m051_qrcode", new QrcodeRetriever(ds));


        exporter.AddColumn("開班編號", "m05_remark", new ClassNoRetriever(ds));
        exporter.AddColumn("宣導機構", "m02_m01code", new PromotionalOrganizationRetriever(ds));
        exporter.AddColumn("開班年度", "m02_yyy", new OpeningYearRetriever(ds));
        exporter.AddColumn("宣導日期", "m02_sdate", new PromotionDateRetriever(ds));
        exporter.AddColumn("場次名稱", "m02_name", new SessionNameRetreiver(ds));
        
        exporter.AddColumn("時數", "m02_hrs", new HoursRetriever(ds));
        exporter.AddColumn("宣導活動項目", "m02_content1", new Content1Retriever(ds));
        exporter.AddColumn("體驗活動項目", "m02_content2", new Content2Retriever(ds));


        return exporter.GenerateWorkBook(ds);
    }

    private string get_addsql()
    {
        return "A series of SQL commands need to be appended.";
    }

}


//--------------------------------------------------------------------------------------

public interface IDataRetriever
{
    string DataRetrieve(string columnID, int dataRowIndex);
}

public class DataSetRetriever: IDataRetriever
{
    DataSet _ds;

    public DataSetRetriever(DataSet dataset)
    {
        _ds = dataset;
    }
    public string DataRetrieve(string columnID, int rowIndex)
    {
        return  _ds.Tables[0].Rows[rowIndex][columnID].ToString();
    }
}

public class SqlRetriever : IDataRetriever
{
    string _SqlQuery;
    public SqlRetriever(string sqlQuery)
    {
        _SqlQuery = sqlQuery;
    }
    public string DataRetrieve(string columnID, int rowIndex )
    {
        string result = "";
        DbMethods.uf_ExecSQL(_SqlQuery, ref result);
        Console.WriteLine(result);
        return result;

    }
}
public class QrcodeRetriever : IDataRetriever
{
    DataSet _ds;
    public QrcodeRetriever(DataSet dataset)
    {
        _ds = dataset;
    }
    public string DataRetrieve(string columnID, int rowIndex)
    {
        string ls_QRcode = "";
        string ls_sql = "SELECT dbo.uf_get_qrcode('" + Security.uf_SQL(_ds.Tables[0].Rows[rowIndex]["m051_m05licno"].ToString()) 
                                    + "', " + Security.uf_SQL(_ds.Tables[0].Rows[rowIndex]["m051_seq"].ToString()) 
                                    + ", '" + Security.uf_SQL(_ds.Tables[0].Rows[rowIndex]["m05_idno"].ToString()) + "')";
        DbMethods.uf_ExecSQL(ls_sql, ref ls_QRcode);
        return ls_QRcode;
    }
}

public class ClassNoRetriever : IDataRetriever
{
    DataSet _ds;
    public ClassNoRetriever(DataSet dataset)
    {
        _ds = dataset;
    }
    public string DataRetrieve(string columnID, int rowIndex)
    {
        string ls_classno = "";
        string ls_sql = @"select m051_remark from flm051 where m051_m05licno = '" 
                        + Security.uf_SQL(_ds.Tables[0].Rows[rowIndex]["m051_m05licno"])
                        + "' and m051_seq = '" + Security.uf_SQL(_ds.Tables[0].Rows[rowIndex]["m051_seq"]) + "'";
        DbMethods.uf_ExecSQL(ls_sql, ref ls_classno);
        string test_input = "開班編號:0111101學員編號:010007";

        return GetClassNo(test_input);
    }

    private string GetClassNo(string input)
    {
        string pattern = @"開班編號:(\d+)";
        string result = "";

        Match match = Regex.Match(input, pattern);
        if (match.Success)
        {
            // 獲取表達式中的捕獲組
            result = match.Groups[1].Value;
            Console.WriteLine("開班編號: " + result);
        }
        else
        {
            Console.WriteLine("找不到開班編號");
        }
        return result;
    }
}

public class PromotionalOrganizationRetriever : IDataRetriever
{
    DataSet _ds;
    IDataRetriever _classnoRetriever;
    public PromotionalOrganizationRetriever(DataSet dataset)
    {
        _ds = dataset;
        _classnoRetriever = new ClassNoRetriever(_ds);
    }

    public string DataRetrieve(string columnID, int rowIndex)
    {
        string ls_promotionalOrganization = "";
        string ls_sql_m02_yyy = @"SELECT d.m01_cname FROM flm051 AS a JOIN flm02 AS b ON 
               SUBSTRING(a.m051_remark, 
               CHARINDEX('開班編號:', a.m051_remark) + 5,CHARINDEX('學員編號:', a.m051_remark) - CHARINDEX('開班編號:', a.m051_remark) - 5) = b.m02_classno
                JOIN flm01 AS d ON b.m02_m01code = d.m01_code
               WHERE
                    a.m051_m05licno = '" + _ds.Tables[0].Rows[rowIndex]["m051_m05licno"] + "'AND a.m051_seq = '" + _ds.Tables[0].Rows[rowIndex]["m051_seq"] + "'; ";

        DbMethods.uf_ExecSQL(ls_sql_m02_yyy, ref ls_promotionalOrganization);
        return ls_promotionalOrganization;
    }
}

public class OpeningYearRetriever : IDataRetriever
{
    DataSet _ds;
    IDataRetriever _classnoRetriever;
    public OpeningYearRetriever(DataSet dataset)
    {
        _ds = dataset;
        _classnoRetriever = new ClassNoRetriever(_ds);
    }

    public string DataRetrieve(string columnID, int rowIndex)
    {
        string ls_openingYear = "";
        //  先取得開班編號
        string ls_classno = _classnoRetriever.DataRetrieve(columnID, rowIndex);
        string ls_sql_m02_yyy = @"select m02_yyy + '年' from flm02 where m02_classno = '" + ls_classno + "' ";
        DbMethods.uf_ExecSQL(ls_sql_m02_yyy, ref ls_openingYear);
        return ls_openingYear;
    }
}

public class PromotionDateRetriever : IDataRetriever
{
    DataSet _ds;
    IDataRetriever _classnoRetriever;
    public PromotionDateRetriever(DataSet dataset)
    {
        _ds = dataset;
        _classnoRetriever = new ClassNoRetriever(_ds);
    }

    public string DataRetrieve(string columnID, int rowIndex)
    {
        string ls_promotionDate = "";
        //  先取得開班編號
        string ls_classno = _classnoRetriever.DataRetrieve(columnID, rowIndex);
        
        string ls_sql_m02_sdate = @"select FORMAT(m02_sdate, 'yyyy-MM-dd') + '~' + FORMAT(m02_edate, 'yyyy-MM-dd')from flm02 where m02_classno = '" + ls_classno + "' ";
        DbMethods.uf_ExecSQL(ls_sql_m02_sdate, ref ls_promotionDate);
        return ls_promotionDate;
    }
}

public class SessionNameRetreiver :IDataRetriever
{
    DataSet _ds;
    IDataRetriever _classnoRetriever;
    public SessionNameRetreiver(DataSet dataset)
    {
        _ds = dataset;
        _classnoRetriever = new ClassNoRetriever(_ds);
    }

    public string DataRetrieve(string columnID, int rowIndex)
    {
        string ls_sessionName = "";
        //  先取得開班編號
        string ls_classno = _classnoRetriever.DataRetrieve(columnID, rowIndex);
        
        string ls_sql_m02_name = @"select m02_name from flm02 where m02_classno = '" + ls_classno + "' ";
        DbMethods.uf_ExecSQL(ls_sql_m02_name, ref ls_sessionName);
        return ls_sessionName;
    }
}

public class HoursRetriever : IDataRetriever
{
    DataSet _ds;
    IDataRetriever _classnoRetriever;
    public HoursRetriever(DataSet dataset)
    {
        _ds = dataset;
        _classnoRetriever = new ClassNoRetriever(_ds);
    }

    public string DataRetrieve(string columnID, int rowIndex)
    {
        string ls_hours = "";
        //  先取得開班編號
        string ls_classno = _classnoRetriever.DataRetrieve(columnID, rowIndex);
        
        string ls_sql_m02_hrs = @"select CONVERT(VARCHAR(50), m02_hrs) + '小時' from flm02 where m02_classno = '" + ls_classno + "' ";

        DbMethods.uf_ExecSQL(ls_sql_m02_hrs, ref ls_hours);
        return ls_hours;
    }
}

public class Content1Retriever : IDataRetriever
{
    DataSet _ds;
    IDataRetriever _classnoRetriever;
    public Content1Retriever(DataSet dataset)
    {
        _ds = dataset;
        _classnoRetriever = new ClassNoRetriever(_ds);
    }

    public string DataRetrieve(string columnID, int rowIndex)
    {
        string ls_contents1 = "";
        //  先取得開班編號
        string ls_classno = _classnoRetriever.DataRetrieve(columnID, rowIndex);
        
        string ls_sql_m02_content1 = @"select m02_content1 from flm02 where m02_classno = '" + ls_classno + "' ";
        DbMethods.uf_ExecSQL(ls_sql_m02_content1, ref ls_contents1);
        return ls_contents1;
    }
}
public class Content2Retriever : IDataRetriever
{
    DataSet _ds;
    IDataRetriever _classnoRetriever;
    public Content2Retriever(DataSet dataset)
    {
        _ds = dataset;
        _classnoRetriever = new ClassNoRetriever(_ds);
    }

    public string DataRetrieve(string columnID, int rowIndex)
    {
        string ls_contents2 = "";
        //  先取得開班編號
        string ls_classno = _classnoRetriever.DataRetrieve(columnID, rowIndex);
        
        string ls_sql_m02_content2 = @"select m02_content2 from flm02 where m02_classno = '" + ls_classno + "' ";
        DbMethods.uf_ExecSQL(ls_sql_m02_content2, ref ls_contents2);
        return ls_contents2;
    }
}
//  建立欄位名稱資訊
public class Column
{
    public string HeaderName { get; set; }
    public string ID { get; set; }
    public IDataRetriever DataRetriever { get; set; }

    public Column(string name, string id, IDataRetriever dataretriever)
    {
        HeaderName = name;
        ID = id;
        DataRetriever = dataretriever;
    }

    public string GetData(int rowIndex)
    {
        return DataRetriever.DataRetrieve(ID, rowIndex);
    }
}




//  建立 Excel 匯出類別 (在 Add column 的時候，也需要給column retriever)
class ExcelExporter
{
    //  欄位名稱
    List<Column> _columns;
    XSSFWorkbook _workbook;

    public ExcelExporter()
    {
        _workbook = new XSSFWorkbook();
        _columns = new List<Column>();
    }

    //  建立匯出的欄位，和取值的方式
    public void AddColumn(string headerName, string id,IDataRetriever dataRetriever)
    {
        _columns.Add(new Column(headerName, id,dataRetriever));
    }

    //  產生欄位
    public XSSFWorkbook  GenerateWorkBook(DataSet ds)
    {
        ISheet sheet = _workbook.CreateSheet("匯出卡片編號");
        CreateHeaderRow(sheet);


        for (int rowIndex = 0; rowIndex < ds.Tables[0].Rows.Count; rowIndex++)
        {
            IRow dataRow;
            dataRow = sheet.CreateRow(rowIndex + 1);
            Debug.Write($"\n第{rowIndex +1} 行\n");
            for (int li_cr = 0; li_cr < _columns.Count; li_cr++)
            {
                Debug.WriteLine($"{_columns[li_cr].HeaderName} : {_columns[li_cr].GetData(rowIndex)}");
                ICell cell = dataRow.CreateCell(li_cr);
            }
        }

        return _workbook;
    }
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