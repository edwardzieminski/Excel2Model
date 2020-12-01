using Excel2Model.Models;
using Excel2Model.Validation;
using Optional;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel2Model.Utilities
{
    public static class ExcelInteropUtilities
    {
        public static Option<Excel.Worksheet, ValidationError> TryGetWorksheet(WorksheetModel worksheetModel)
        {
            if (FilesUtilities.FileExists(worksheetModel.WorkbookPath) == false)
                return Option.None<Excel.Worksheet, ValidationError>(new ValidationError("Provided file does not exist"));

            if (IsWorksheetIndexOrNameProvided(worksheetModel))
                return Option.None<Excel.Worksheet, ValidationError>(new ValidationError("Provided worksheet is incorrect."));

            if (AreBothWorksheetIndexAndNameProvided(worksheetModel))
                return Option.None<Excel.Worksheet, ValidationError>(new ValidationError("Please provide either worksheet name or index. Not both in the same time."));

            Excel.Application excelApp;
            Excel.Workbook excelWkb;
            Excel.Worksheet excelWks = new Excel.Worksheet();

            try
            {
                excelApp = new Excel.Application();
            }
            catch
            {
                return Option.None<Excel.Worksheet, ValidationError>(new ValidationError("Could not open Excel application."));
            }

            try
            {
                excelWkb = excelApp.Workbooks.Open(worksheetModel.WorkbookPath, ReadOnly: true);
            }
            catch
            {
                return Option.None<Excel.Worksheet, ValidationError>(new ValidationError("Provided file was not recognized as a workbook by Excel interop."));
            }

            if (IsWorksheetNameProvided(worksheetModel))
            {
                try
                {
                    excelWks = (Excel.Worksheet)excelWkb.Worksheets[worksheetModel.WorksheetName];
                }
                catch
                {
                    return Option.None<Excel.Worksheet, ValidationError>(new ValidationError("Provided worksheet name is not correct."));
                }
            }

            if (IsWorksheetIndexCorrect(worksheetModel))
            {
                try
                {
                    excelWks = (Excel.Worksheet)excelWkb.Worksheets[worksheetModel.WorksheetIndex];
                }
                catch
                {
                    return Option.None<Excel.Worksheet, ValidationError>(new ValidationError("Provided worksheet index is not correct."));
                }
            }

            return Option.Some<Excel.Worksheet, ValidationError>(excelWks);
        }

        public static bool IsWorksheetIndexCorrect(WorksheetModel worksheetModel) =>
            worksheetModel.WorksheetIndex > 0;

        public static bool IsWorksheetNameProvided(WorksheetModel worksheetModel) =>
            string.IsNullOrWhiteSpace(worksheetModel.WorksheetName) == false;

        public static bool IsWorksheetIndexOrNameProvided(WorksheetModel worksheetModel) =>
            (IsWorksheetIndexCorrect(worksheetModel) || IsWorksheetNameProvided(worksheetModel));

        public static bool AreBothWorksheetIndexAndNameProvided(WorksheetModel worksheetModel) =>
            (IsWorksheetIndexCorrect(worksheetModel) && IsWorksheetNameProvided(worksheetModel));
    }
}
