namespace Excel2Model.Models
{
    public record WorksheetModel
    {
        public string WorkbookPath { get; init; }
        public string WorksheetName { get; init; }
        public int WorksheetIndex { get; init; }
        public int HeaderRow { get; init; }
        public int DataStartRow { get; init; }
    }
}