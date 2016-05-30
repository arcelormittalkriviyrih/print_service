using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace UploadTemplate
{
    class Program
    {
        static void Main(string[] args)
        {
            //UploadLabelTemplate.exe Files.ID D:\template.xls
            SqlConnection sqConn = null;
            SqlCommand sqComm = null;
            FileStream fs = null;
            try
            {
                sqConn = new SqlConnection(@"data source=MSSQL2014SRV\MSSQLSERVER2012; initial catalog = B2MML-BatchML; persist security info = True; user id = vimas; password = mercury; MultipleActiveResultSets = true; Pooling = true;"); // строка соединения
                sqConn.Open();
                sqComm = new SqlCommand("UPDATE Files SET Data = @Data WHERE ID = @ID", sqConn);

                fs = new FileStream(args[1], FileMode.Open);  // открываем файл
                byte[] fileBuffer = new byte[fs.Length];
                fs.Read(fileBuffer, 0, (int)fs.Length);                                               // читаем в бинарный буфер
                fs.Close();

                sqComm.Parameters.AddWithValue("@Data", null); //System.Data.DbType.Binary 
                sqComm.Parameters.AddWithValue("@ID", args[0]);
                sqComm.Parameters["@Data"].Value = fileBuffer;   // записываем бинарный буфер в значение параметра
                sqComm.ExecuteNonQuery();                                                       // добавляем запись в базу
                sqConn.Close();
                Console.WriteLine("Ok update");
            }
            catch (Exception ex)
            {
                if (fs != null) fs.Dispose();
                if (sqComm != null) (sqComm as IDisposable).Dispose();
                if (sqConn != null) (sqConn as IDisposable).Dispose();
                Console.WriteLine(ex);
            }
            Console.ReadKey();
        }
    }
}
