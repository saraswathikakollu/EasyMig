using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using EasyMig.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing;
using System.Drawing;
using System.IO;
using System.Configuration;
using System.Data.SqlClient;

namespace EasyMig.Utility
{
    public class TestData
    {
        public void GenerateTestData(Int32 noOfRecords)
        {
            EasyMig.Utility.SampleCustomerRepository repository = new SampleCustomerRepository();           
            // to get sample data
            var customers = repository.GetCustomers(noOfRecords);
          
            //Create datatable
            DataTable dt = CreateDataTable(customers);
          
            //export data to excel file 
            string fileName = ExportToExcel(dt);
           
            SaveFile(fileName);
        }

        public void GenerateMaterialMasterTestData(int Records)
        {
            EasyMig.Utility.SampleCustomerRepository repository = new SampleCustomerRepository();
            // to get sample data
            var materialmasters = repository.GetMaterialMaster(Records);

            //Create datatable
            DataTable dt2 = CreateDataTableMaterialMaster(materialmasters);

            //export data to excel file 
            string fileName = ExportToExcel(dt2);

            SaveFile(fileName);
        }

      //  public void GenerateMaterialPlantTestData(int records)
        //{
          //  EasyMig.Utility.SampleCustomerRepository repository = new SampleCustomerRepository();
            //// to get sample data
            //var materialplant = repository.GetMaterialPlant(records);

            ////Create datatable
            //DataTable dt3 = CreateDataTableMaterialPlant(materialplant);

            ////export data to excel file 
            //string fileName = ExportToExcel(dt3);

            //SaveFile(fileName);
        //}
        /// <summary>
        /// This method used to get the test data which already created
        /// </summary>
        /// <returns></returns>
        public List<EasyMig.Models.TestDataFile> GetTestData()
        {
            List<EasyMig.Models.TestDataFile> testDataFiles = GetSampleDataFiles();
            return testDataFiles;
        }

        /// <summary>
        /// This method is to get test data files from database
        /// </summary>
        /// <returns></returns>
        public List<Models.TestDataFile> GetSampleDataFiles()
        {
            string strCon = ConfigurationManager.ConnectionStrings["easyMigCon"].ConnectionString.ToString();
            string strPath = ConfigurationManager.AppSettings["filePath"].ToString();
            SqlConnection con = new SqlConnection(strCon);
            SqlDataAdapter da = new SqlDataAdapter();
            DataSet ds = new DataSet();
            List<Models.TestDataFile> objFile = new List<TestDataFile>();
            try
            {
                SqlCommand cmd = new SqlCommand();
                cmd.CommandText = "SELECT * FROM TestDataFile";
                cmd.Connection = con;
                da.SelectCommand = cmd;
                da.Fill(ds);

                if (ds.Tables[0].Rows.Count > 0)
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        TestDataFile obj = new TestDataFile();
                        obj.Id = Convert.ToInt32(ds.Tables[0].Rows[i]["Id"]);
                        obj.Name = ds.Tables[0].Rows[i]["Name"].ToString();
                        obj.CreatedOn = Convert.ToDateTime(ds.Tables[0].Rows[i]["CreatedOn"]);
                        obj.DownloadPath = strPath + ds.Tables[0].Rows[i]["Name"].ToString();
                        objFile.Add(obj);
                    }
                }
                return objFile;

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        public void SaveFile(string fileName)
        {
            string strCon = ConfigurationManager.ConnectionStrings["easyMigCon"].ConnectionString.ToString();
            SqlConnection con = new SqlConnection(strCon);
            try
            {
                SqlCommand cmd = new SqlCommand();
                cmd.CommandText = "INSERT INTO TestDataFile (Name,CreatedOn) VALUES ('" + fileName + "','" + DateTime.Now + "')";
                cmd.Connection = con;
                con.Open();
                cmd.ExecuteNonQuery();
            }
            catch(Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        public DataTable CreateDataTable(IEnumerable<Customer> customers)
        {
            DataTable Dt = new DataTable("SampleData");

            Dt.Columns.Add(new DataColumn("Id", typeof(string)));
            Dt.Columns.Add(new DataColumn("FirstName", typeof(string)));
            Dt.Columns.Add(new DataColumn("LastName", typeof(string)));
            Dt.Columns.Add(new DataColumn("City", typeof(string)));
            Dt.Columns.Add(new DataColumn("Phone", typeof(string)));
            Dt.Columns.Add(new DataColumn("ZipCode", typeof(string)));
            DataRow dr;

            Customer c1;

            foreach (object obj in customers)
            {
                c1 = (Customer)obj;
                dr = Dt.NewRow();
                //Add rows  
                dr["Id"] = c1.Id;
                dr["FirstName"] = c1.FirstName;
                dr["LastName"] = c1.LastName;
                dr["City"] = c1.City;
                dr["Phone"] = c1.Phone;
                dr["ZipCode"] = c1.ZipCode;



                Dt.Rows.Add(dr);
            }
            return Dt;
        }

        public DataTable CreateDataTableMaterialMaster(IEnumerable<MaterialMaster> materialmaster)
        {
            DataTable Dt2 = new DataTable("MaterialMasterBasicDataStructure");

            Dt2.Columns.Add(new DataColumn("Material_No", typeof(Int64)));
            Dt2.Columns.Add(new DataColumn("Industry_Sector", typeof(string)));
            Dt2.Columns.Add(new DataColumn("Material_Type", typeof(string)));
            Dt2.Columns.Add(new DataColumn("Material_Group", typeof(string)));
            Dt2.Columns.Add(new DataColumn("Old_Material", typeof(string)));
            Dt2.Columns.Add(new DataColumn("Base_Unit", typeof(string)));
            Dt2.Columns.Add(new DataColumn("Gross_Weight", typeof(Double)));
            Dt2.Columns.Add(new DataColumn("Weight_Unit", typeof(string)));
            Dt2.Columns.Add(new DataColumn("Material_Description", typeof(string)));
            DataRow dr2;

            MaterialMaster c2;
            Int64 matNo = 1100013;
            Boolean isData = false;
            Int32 i = 1;

            Int64 oldmat = 74556;
            Boolean isData1 = false;
            Int32 j = 1;

            Double Grosswt = 0.5;
            Boolean isData2 = false;
            Double k = 1.0;

         //   string[] strArray =  { "Hawa", "ROH", "HIBE", "AEM", "NLAG", "VERP", "FERT", "IBAU", "HALB", "BLG", "ERSA", "KMAT", "DIEN", "ZMDG", "FRIP", "VOLL", "FOOD" };
           


            foreach (object obj in materialmaster)
            {

               // Random rand = new Random();
               // int index = rand.Next(strArray.Length);


                c2 = (MaterialMaster)obj;
                dr2 = Dt2.NewRow();
                //Add rows  
                dr2["Material_No"] = matNo; 
                dr2["Industry_Sector"] = c2.Industry_Sector;
                dr2["Material_Type"] = isData ? "HAWA" : "ROH";
                dr2["Material_Group"] = "01";
                dr2["Old_Material"] = oldmat;
                dr2["Base_Unit"] = isData ? "EA" : "KG";
                dr2["Gross_Weight"] = Grosswt;
                dr2["Weight_Unit"] = isData ? "EA" : "KG";
                dr2["Material_Description"] = "MATERIAL " + i;

                isData = isData ? false : true;
                matNo++;
                i++;

                isData1 = isData1 ? false : true;
                oldmat++;
                j++;

                isData2 = isData2 ? false : true;
                Grosswt++;
                k++;


                Dt2.Rows.Add(dr2);
            }
            return Dt2;
        }

      // public DataTable CreateDataTableMaterialPlant(IEnumerable<MaterialPlant> materialplant)
      //  {
        //    DataTable Dt3 = new DataTable("MaterialPlantBasicDataStructure");

          //  Dt3.Columns.Add(new DataColumn("Material_No", typeof(Int64)));
            //Dt3.Columns.Add(new DataColumn("Plant", typeof(Int64)));
        //}
        
            public void ReadStructure()
        {
            //TODO:
        }

        public string ExportToExcel(DataTable dt)
        {
            try
            {
                //string strPath = @"D:\Projects\EasyMIG\EasyMig\EasyMig\ExcelExports\";
                string strPath = ConfigurationManager.AppSettings["filePath"].ToString();
                string fName = string.Empty;

                using (ExcelPackage objPackage = new ExcelPackage())
                {
                    ExcelWorksheet objSheet = objPackage.Workbook.Worksheets.Add(dt.TableName);
                    objSheet.Cells["A1"].LoadFromDataTable(dt,true);
                    objSheet.Cells.Style.Font.SetFromFont(new Font("Calibri", 10));
                    objSheet.Cells.AutoFitColumns();
                    //Format the header    
                    using (ExcelRange objRange = objSheet.Cells["A1:XFD1"])
                    {
                        objRange.Style.Font.Bold = true;
                        objRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        objRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }
                    Random random = new Random();
                    fName = "SampleData" + DateTime.Now.ToShortDateString() + random.Next().ToString() + ".xlsx";
                    strPath = strPath + fName;

                    //Write it back to the client    
                    if (File.Exists(strPath))
                        File.Delete(strPath);

                    //Create excel file on physical disk    
                    FileStream objFileStrm = File.Create(strPath);
                    objFileStrm.Close();

                    //Write content to excel file    
                    File.WriteAllBytes(strPath, objPackage.GetAsByteArray());
                }
                return fName;
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }


        public List<Mandate> ValidateExcelFile(string file)
        {
            try
            {
                string strPath = ConfigurationManager.AppSettings["filePath"].ToString();
                DataTable dt = new DataTable();
                bool hasHeader = true;

                using (ExcelPackage exlPackage = new ExcelPackage())
                {
                    using (var fstream = File.OpenRead(file))
                    {
                        exlPackage.Load(fstream);
                    }

                    var ws = exlPackage.Workbook.Worksheets[1];
                    foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                    {
                        dt.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                    }

                    var startRow = hasHeader ? 2 : 1;
                    for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                    {
                        var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                        DataRow row = dt.Rows.Add();
                        foreach (var cell in wsRow)
                        {
                            row[cell.Start.Column - 1] = cell.Text;
                        }
                    }
                }

                var mandate = MandatoryCheck(dt);

                return mandate;
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }
        

        public List<Mandate> MandatoryCheck(DataTable dt)
        {            
            try
            {
                string[] mandateColumns = ConfigurationManager.AppSettings["mandateColumns"].ToString().Split(',');
                List<Mandate> lstMandate = new List<Mandate>();
                Mandate mandate = null;
                DataRow dRow = null;
                DataTable dtResult = dt.Clone();
                int colIndex = 0;

                for(int i=0; i< mandateColumns.Length;i++)
                {
                    //dRow = dt.Rows[i];
                    colIndex = dt.Columns[mandateColumns[i]].Ordinal;
                    if(dt.Columns["First Name"].ToString() == mandateColumns[i].ToString())
                    {
                        for(int j=0; j < dt.Rows.Count; j++)
                        {
                            if (string.IsNullOrWhiteSpace(dt.Rows[j][colIndex].ToString()))
                            {
                                mandate = new Mandate();
                                mandate.rowId = Convert.ToInt32(dt.Rows[j][0]);
                                mandate.FirstName = dt.Rows[j][colIndex].ToString();
                                mandate.LastName = dt.Rows[j]["Last Name"].ToString();
                                lstMandate.Add(mandate);
                            }
                        }
                    }
                }
                return lstMandate;
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }


    }
}