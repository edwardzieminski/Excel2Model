using Excel = Microsoft.Office.Interop.Excel;
using Excel2Model;
using Xunit;
using System;

namespace Excel2ModelUnitTests
{
    public class GeneralTest
    {
        [Fact]
        public void Is_it_working()
        {
            var concreteMapper = new ConcreteMapper();
            var listOfModels = concreteMapper.GetDataFromExcel();
            Console.WriteLine();
        }
    }

    class SomeModel
    {
        public string KolumnaA { get; set; }
        public string KolumnaB { get; set; }
    }

    class ConcreteMapper : Mapper<SomeModel>
    {
        public ConcreteMapper()
        {
            var excelApp = new Excel.Application();
            var workbook = excelApp.Workbooks.Open(@"C:\Users\eziemins\Desktop\Makra i inne\2020-11 listopad\Excel2Model\test.xlsx");
            var worksheet = (Excel.Worksheet)workbook.Worksheets[0];
            Worksheet = worksheet;
            DataStartRow = 1;
            AddColumn("A", someModel => someModel.KolumnaA);
            AddColumn("B", someModel => someModel.KolumnaB);
        }
    }
}
