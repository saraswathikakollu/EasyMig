using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EasyMig.Models
{
    public class Customer
    {
        public Guid Id { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string City { get; set; }
        public string Phone { get; set; }
        public string ZipCode { get; set; }
    }

    public class MaterialMaster
    {
        public Int64 Material_No { get; set; }
        public string Industry_Sector { get; set; }
        public string Material_Type { get; set; }
        public string Material_Group { get; set; }
        public string Old_Material { get; set; }
        public string Base_Unit { get; set; }
        public string Gross_Weight { get; set; }
        public string Weight_Unit { get; set; }
        public string Material_Description { get; set; }
    }


  //  public class MaterialPlant
    //{
      //  public Int64 Material_No { get; set; }
        //public Int64 Plant { get; set; }
    //}
}







