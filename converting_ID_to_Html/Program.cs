using OfficeOpenXml;
using OfficeOpenXml.Packaging.Ionic.Zlib;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Xml;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace converting_ID_to_Html
{
    internal class Program
    {
        private static List<int> ID_For_The_Database = new List<int>();
        static void Main(string[] args)
        {
            int id = 210996;
            if (id != null)
            {
 
                byte[] HtmlDataByteS = GetHtmlData(id);
                byte[] HtmlByteS = DeCompressDoc(HtmlDataByteS);
                File.WriteAllBytes(@"C:\Users\j.shepelev\Desktop\test\123.pdf", HtmlByteS);

            }
            //GetAllIDFromSource(); 
            bool Ver9711 = false;

            if (Ver9711 == true)
            {
                for (int i = 0; i < ID_For_The_Database.Count; i++)
                {
                    string path = @"C:\Users\j.shepelev\Desktop\test\" + ID_For_The_Database[i] + ".html";
                    byte[] HtmlDataByteS = GetHtmlData(ID_For_The_Database[i]);
                    byte[] HtmlByteS = DeCompressDoc(HtmlDataByteS);
                    File.WriteAllBytes(path, HtmlByteS);
                }
            }
            else
            {
                for (int i = 0; i < ID_For_The_Database.Count; i++)
                {
                    string path = @"C:\Users\j.shepelev\Desktop\test\" + ID_For_The_Database[i] + ".html";
                    var HtmlInfo = GetHtmlInfo(ID_For_The_Database[i]);
                    File.WriteAllText(path, HtmlInfo);
                }

            }

        }

        public static string GetHtmlInfo(int idOrder)
        {
            using (SqlConnection connection = GetDebtConnection_6091())
            using (SqlCommand cmd = new SqlCommand(string.Empty, connection))
            {
                cmd.CommandText = @"SELECT 
                                           [HtmlData]
                                    FROM [dbo].[collectorder_xml]
                                    WHERE [ID_CollectOrder] = @ID_CollectOrder";

                cmd.Parameters.AddWithValue("@ID_CollectOrder", idOrder);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        byte[] htmlBytes = reader.GetData<byte[]>(0);
                        var HtmlText = htmlBytes != null ? GZipStream.UncompressString(htmlBytes) : null;
                        return HtmlText;
                    }
                }

            }
            throw new InvalidOperationException("Не удалось сжать документ");
        }
        public static byte[] GetHtmlData(int ID)
        {
            using (SqlConnection connection = GetDebtConnection_9711())
            using (SqlCommand cmd = new SqlCommand(string.Empty, connection))
            {
                cmd.CommandText = @"SELECT PdfData
                                        FROM [DebtPipeline].[dbo].[Xml]
                                        where ID_Pipeline = @ID;";

                cmd.Parameters.AddWithValue("@ID", ID);
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        byte[] HtmlDataByteS = reader.GetData<byte[]>(0);
                        return HtmlDataByteS;

                    }
                }
                throw new InvalidOperationException("Не удалось сжать документ");
            }
        }
        public static byte[] DeCompressDoc(byte[] HtmlDataByteS)
        {


            using (SqlConnection connection = GetDebtConnection_9711())
            using (SqlCommand cmd = new SqlCommand(string.Empty, connection))
            {
                cmd.CommandText = @"SELECT dbo.decompress(@strongbytes)";

                cmd.Parameters.AddWithValue(@"strongbytes", HtmlDataByteS);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        byte[] strongCompress = reader.GetData<byte[]>(0);
                        return strongCompress;
                    }
                }
            }
            throw new InvalidOperationException("Не удалось сжать документ");
        }
        public static SqlConnection GetDebtConnection_9711(bool needOpen = true)
        {
            string str = "Data Source=RVDK-SVR-9711;Initial Catalog=DebtPipeline;MultipleActiveResultSets=true";
            SqlConnection connection = new SqlConnection(str);

            if (needOpen)
                connection.Open();

            return connection;
        }
        public static SqlConnection GetDebtConnection_6091(bool needOpen = true)
        {
            string str = "Data Source=RVDK-SVR-6091,1500;Initial Catalog=Debt;Integrated Security=True";
            SqlConnection connection = new SqlConnection(str);

            if (needOpen)
                connection.Open();

            return connection;
        }
        private static void GetAllIDFromSource()
        {
            //string ExcelPath = Environment.CurrentDirectory + "\\Address.xlsx";

            string ExcelPath = Environment.CurrentDirectory + "\\Test_Fias.xlsx";
            FileInfo info = new FileInfo(ExcelPath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excel = new ExcelPackage(info))
            {
                var worksheet = excel.Workbook.Worksheets["Лист1"];
                try
                {
                    for (int i = 0; !string.IsNullOrEmpty(worksheet.Cells[i + 2, 1].Text); i++)
                    {
                        string res = worksheet.Cells[i, 1].Text;
                        int result = int.Parse(res);
                        ID_For_The_Database.Add(result);
                    }
                }
                catch
                {
                    ;
                }
            }
        }
    }
}