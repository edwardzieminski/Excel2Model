namespace Excel2Model.Models
{
    public record WorksheetModel (string WorkbookPath, string WorksheetName, int WorksheetIndex, int HeaderRow, int DataStartRow)
    {
        
    }
}