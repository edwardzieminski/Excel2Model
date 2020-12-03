using Excel2Model.Models;
using FluentValidation;


namespace Excel2Model.Validation
{
    public class ColumnMapValidator : AbstractValidator<ColumnMapModel>
    {
        public ColumnMapValidator()
        {
            When
            (
                predicate:  x => x.ColumnIndex != default,
                action:     () =>
                {
                    RuleFor(x => x.ColumnIndex)
                        .Cascade(CascadeMode.Stop)
                        .InclusiveBetween(1, 702)
                        .WithMessage(x => $"Column index should inclusively fall between 1 and 702. Please correct column: \n{ x }");
                }
            );
            

            RuleFor(x => x.ColumnName)
                .Cascade(CascadeMode.Stop)
                .NotEmpty()
                .WithMessage(x => $"Column not found. Please correct column: \n{ x }")
                .Matches("^[A-Z]{1,2}$")
                .WithMessage(x => $"Column name should fall between A and ZZ. Please correct column: \n{ x }");
        }
    }
}
