using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using EasyMig.Models;
using StackExchange.Redis;
using Bogus;

namespace EasyMig.Utility
{
    public class SampleCustomerRepository
    {
        public IEnumerable<Customer> GetCustomers(Int32 noOfRecords)
        {
            Randomizer.Seed = new Random(123456);
            var customerGenerator = new Faker<Customer>()
                .RuleFor(c => c.Id, Guid.NewGuid())
                .RuleFor(c => c.FirstName, f => f.Person.FirstName)
                .RuleFor(c => c.LastName, f => f.Person.LastName)
                .RuleFor(c => c.City, f => f.Person.Address.City)
                .RuleFor(c => c.ZipCode, f => f.Person.Address.ZipCode)
                .RuleFor(c => c.Phone, f => f.Phone.PhoneNumber());

            return customerGenerator.Generate(noOfRecords);

        }

        public IEnumerable<MaterialMaster> GetMaterialMaster(Int32 noOfRecords)
        {
            Randomizer.Seed = new Random(123456789);
            var materialGenerator = new Faker<MaterialMaster>()
                .RuleFor(c => c.Material_No, 1)
                .RuleFor(c => c.Industry_Sector, f => "1")
                .RuleFor(c => c.Material_Type, f => f.Vehicle.Type().ToString())
                .RuleFor(c => c.Material_Group, f => f.Person.UserName)
                .RuleFor(c => c.Old_Material, f => f.Vehicle.Model())
                .RuleFor(c => c.Base_Unit, f => "1")
                .RuleFor(c => c.Gross_Weight, f => f.Finance.RoutingNumber())
                .RuleFor(c => c.Weight_Unit, f => "1")
                .RuleFor(c => c.Material_Description, f => f.Company.CompanyName());
            return materialGenerator.Generate(noOfRecords);


        }

//        public IEnumerable<MaterialPlant> GetMaterialPlant(Int32 noOfRecords)
  //      {
    //        Randomizer.Seed = new Random(123456789);
      //      var materialplantGenerator = new Faker<MaterialPlant>()
        //        .RuleFor(c => c.Material_No, 1)
          //      .RuleFor(c => c.Plant, f => f.UniqueIndex);
        //}   
    }
}