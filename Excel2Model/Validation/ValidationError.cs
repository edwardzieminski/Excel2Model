namespace Excel2Model.Validation
{
    public class ValidationError
    {
        public ValidationError(string errorMessage)
        {
            ErrorMessage = errorMessage;
        }

        public string ErrorMessage { get; private init; }
    }
}