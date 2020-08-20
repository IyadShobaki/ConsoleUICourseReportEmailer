using ConsoleUI.Models;
using ConsoleUI.Repository;
using ConsoleUI.Workers;
using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;

namespace ConsoleUI
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                EnrollmentDetailReportCommand command = new EnrollmentDetailReportCommand(@"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=CourseReport;Integrated Security=True");
                IList<EnrollmentDetailReportModel> models = command.GetList();

                var reportFileName = "EnrollmentDetailsReport.xlsx";

                EnrollmentDetailReportSpreadSheetCreator enrollmentSheetCreator = new EnrollmentDetailReportSpreadSheetCreator();
                enrollmentSheetCreator.Create(reportFileName, models);

                EnrollmentDetailReportEmailSender emailer = new EnrollmentDetailReportEmailSender();

                emailer.Send(reportFileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Something went wrong: {0}", ex.Message);
            }


            Console.ReadLine();
        }
    }
}



//using ConsoleUI.Models;
//using Newtonsoft.Json;
//using System;
//using System.Data;

//namespace ConsoleUI
//{
//    class Program
//    {
//        static void Main(string[] args)
//        {
//            EnrollmentDetailReportModel model = new EnrollmentDetailReportModel()
//            {
//                EnrollmentId = 1,
//                FirstName = "Mark",
//                LastName = "Hue",
//                CourseCode = "CA",
//                Description = "description"
//            };

//            var json = JsonConvert.SerializeObject(model); // change the object from C# to json 

//            // change the object from json to c#
//            EnrollmentDetailReportModel objectFromJson = (EnrollmentDetailReportModel)JsonConvert.DeserializeObject(json, typeof(EnrollmentDetailReportModel));


//            DataTable table = new DataTable();
//            DataColumn column1 = new DataColumn("EnrollmentId", typeof(int));
//            DataColumn column2 = new DataColumn("FirstName", typeof(string));
//            DataColumn column3 = new DataColumn("LastName", typeof(string));
//            DataColumn column4 = new DataColumn("CourseCode", typeof(string));
//            DataColumn column5 = new DataColumn("Description", typeof(string));

//            table.Columns.Add(column1);
//            table.Columns.Add(column2);
//            table.Columns.Add(column3);
//            table.Columns.Add(column4);
//            table.Columns.Add(column5);

//            table.Rows.Add(1, "Mark", "Hue", "CA", "description");

//            foreach (DataRow row in table.Rows)
//            {
//                foreach (DataColumn column in table.Columns)
//                {
//                    Console.WriteLine(row[column]);
//                }
//            }

//            Console.ReadLine();
//        }
//    }
//}
