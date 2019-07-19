using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EasyMig.Models
{
    public class TestDataFile
    {
        public Int32 Id { get; set; }
        public string Name { get; set; }
        public DateTime CreatedOn { get; set; }
        public string DownloadPath { get; set; }
    }

    public class Mandate
    {
        public Int32 rowId { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string City { get; set; }
    }

 
    public class Mandate1
    {
        public Int32 rowId { get; set; }
        public int Material_No { get; set; }
        public string Industry_Sector { get; set; }
        public string Material_Type { get; set; }
        public string Material_Group { get; set; }
        public string Old_Material { get; set; }
        public int Base_Unit { get; set; }
        public float Gross_Weight { get; set; }
        public int Weight_Unit { get; set; }
        public string Material_Description { get; set; }
    }
 //   public class Mandate2
   // {
     //   public Int64 Material_No { get; set; }
       // public Int64 Plant { get; set; }
    //}
}