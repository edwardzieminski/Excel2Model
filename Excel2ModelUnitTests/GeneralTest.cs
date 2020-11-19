using Excel = Microsoft.Office.Interop.Excel;
using Excel2Model;
using Xunit;
using System;

namespace Excel2ModelUnitTests
{
    public class GeneralTest
    {
    }

    class SomeModel
    {
        public string SomeProperty { get; set; }
    }

    class ConcreteMapper : Mapper<SomeModel>
    {
        public ConcreteMapper()
        {
            var worksheet = new Excel.Worksheet();
            Worksheet = worksheet;
            HeaderRow = 1;
            DataStartRow = 2;
            AddColumn("A", someModel => someModel.SomeProperty);
        }
    }
}
