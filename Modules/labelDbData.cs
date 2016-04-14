using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;

namespace PrintWindowsService
{
    public class labelDbData
    {
        private SqlConnection dbConnection;
        private SqlCommand selectCommandProdResponse;
        private SqlCommand selectLabelProperty;
        private SqlCommand CommandUpdateStatus;
        private SqlCommand selectCommandFiles;

        public labelDbData(string connectionString)
        {
            dbConnection = new SqlConnection(connectionString);
            selectCommandProdResponse = new SqlCommand("SELECT ID, ResponseState, ProductionRequestID, EquipmentID, EquipmentClassID, ProductSegmentID, ProcessSegmentID\n" +
                  "FROM v_ProductionResponse\n" +
                  "WHERE (ResponseState = @State)\n" +
                  "  AND (EquipmentClassID = @EquipmentClassID)", dbConnection);
            selectCommandProdResponse.Parameters.AddWithValue("@State", "ToPrint");
            selectCommandProdResponse.Parameters.AddWithValue("@EquipmentClassID", "/2/");

            selectLabelProperty = new SqlCommand("SELECT TypeProperty, ClassPropertyID, ValueProperty\n" +
                  "FROM v_PrintProperties\n" +
                  "WHERE (ProductionResponse = @ProductionResponse)", dbConnection);
            selectLabelProperty.Parameters.AddWithValue("@ProductionResponse", null);

            CommandUpdateStatus = new SqlCommand("BEGIN TRANSACTION T1; UPDATE ProductionResponse SET ResponseState = @State WHERE ID = @ProductionResponseID; COMMIT TRANSACTION T1;", dbConnection);
            CommandUpdateStatus.Parameters.AddWithValue("@State", null);
            CommandUpdateStatus.Parameters.AddWithValue("@ProductionResponseID", null);

            selectCommandFiles = new SqlCommand("SELECT pf.Data\n" +
                  "FROM v_ProductionParameter_Files pf\n" +
                  "WHERE pf.ProductSegmentID = @ProductSegmentID\n" +
                  "  AND pf.ProcessSegmentID = @ProcessSegmentID\n" +
                  "  AND pf.PropertyType = @PropertyType"
                  //"  AND pf.FileType = @FileType" ???
                  , dbConnection);
            selectCommandFiles.Parameters.AddWithValue("@ProductSegmentID", null);
            selectCommandFiles.Parameters.AddWithValue("@ProcessSegmentID", null);
            selectCommandFiles.Parameters.AddWithValue("@PropertyType", 1);
        }

        ~ labelDbData()
        {
            selectCommandFiles.Dispose();
            CommandUpdateStatus.Dispose();
            selectLabelProperty.Dispose();
            selectCommandProdResponse.Dispose();
            dbConnection.Dispose();
        }

        public void fillJobData(ref List<jobProps> resultData)
        {
            //List<jobProps> resultData = new List<jobProps>();

            try
            {
                dbConnection.Open();

                using (SqlDataReader dbReaderProdResponse = selectCommandProdResponse.ExecuteReader())
                {
                    while (dbReaderProdResponse.Read())
                    {
                        //чтение параметров для шаблона и печати
                        selectLabelProperty.Parameters["@ProductionResponse"].Value = dbReaderProdResponse["ID"];
                        DataTable tableLabelProperty = new DataTable();
                        using (SqlDataAdapter adapterLabelProp = new SqlDataAdapter(selectLabelProperty))
                        {
                            adapterLabelProp.Fill(tableLabelProperty);
                        }

                        //чтение шаблона для печати этикетки
                        selectCommandFiles.Parameters["@ProductSegmentID"].Value = dbReaderProdResponse["ProductSegmentID"];
                        selectCommandFiles.Parameters["@ProcessSegmentID"].Value = dbReaderProdResponse["ProcessSegmentID"];
                        byte[] XlFile;
                        using (SqlDataReader dbReaderFiles = selectCommandFiles.ExecuteReader())
                        {
                            dbReaderFiles.Read();
                            XlFile = (byte[])dbReaderFiles["Data"];
                            dbReaderFiles.Close();
                        }

                        resultData.Add(new jobProps((int)dbReaderProdResponse["ID"], XlFile, tableLabelProperty));
                    }
                    dbReaderProdResponse.Close();
                }
            }
            finally
            {
                dbConnection.Close();
            }

            //return resultData;
        }

        public void updateJobStatus(int aProductionResponseID, string aPrintState)
        {
            try
            {
                dbConnection.Open();
                CommandUpdateStatus.Parameters["@ProductionResponseID"].Value = aProductionResponseID;
                CommandUpdateStatus.Parameters["@State"].Value = aPrintState;
                CommandUpdateStatus.ExecuteNonQuery();
            }
            finally
            {
                dbConnection.Close();
            }
        }
    }
}