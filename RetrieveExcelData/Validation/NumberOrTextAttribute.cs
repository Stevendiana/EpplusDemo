using RetrieveExcelData.Models;
using System.ComponentModel.DataAnnotations;

namespace RetrieveExcelData.Validation
{
    public class NumberOrTextAttribute : ValidationAttribute
    {
        public NumberOrTextAttribute(string? dataType, string? celladdress)
        {
            DataType = dataType;
            Celladdress = celladdress;
        }

        public string? DataType { get; }
        public string? Celladdress { get; }

        public string GetDataTypeErrorMessage() =>
            $"Please enter a number as cell value and not a text.";

        public string GetAddressErrorMessage() =>
            $"Please enter a cell address.";

        protected override ValidationResult IsValid(object value,
            ValidationContext validationContext)
        {
            
            var datapropType = validationContext.ObjectType.GetProperty(DataType);
            var addresspropType = validationContext.ObjectType.GetProperty(Celladdress);

            var dataValue = datapropType.GetValue(validationContext.ObjectInstance, null);
            var addressValue = addresspropType.GetValue(validationContext.ObjectInstance, null);

            var dataprop = (dataValue != null) ? dataValue.ToString() : null;
            var addressprop = (addressValue != null) ? addressValue.ToString() : null;
            double number;

            if (value != null && dataprop != null)
            {
                var cellValue = value.ToString();
                
                if (addressprop!=null)
                {
                    if (dataprop.ToLower() == "number" && double.TryParse(cellValue, out number))
                    {
                        return ValidationResult.Success;
                    }
                    if (dataprop.ToLower() == "text")
                    {
                        return ValidationResult.Success;
                    }
                    return new ValidationResult(ErrorMessage ?? GetDataTypeErrorMessage());
                }
               return new ValidationResult(ErrorMessage ?? GetAddressErrorMessage());

            }
            return ValidationResult.Success;
        }
    }
}
