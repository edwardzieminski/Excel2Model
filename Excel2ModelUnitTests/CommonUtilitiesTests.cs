using Xunit;
using Excel2Model.Utilities;

namespace Excel2ModelUnitTests
{
    public class CommonUtilitiesTests
    {
        [Fact]
        public void GetPropertyFromExpression_provides_correct_PropertyInfo_object()
        {
            // ARRANGE
            var expectedPropertyName = "TestPropertyInt";

            // ACT
            var actualPropertyInfo = CommonUtilities.GetPropertyFromExpression<TestClass>(x => x.TestPropertyInt);
            var actualPropertyName = actualPropertyInfo.Name;

            // ASSERT
            Assert.Equal(expectedPropertyName, actualPropertyName);
        }
    }

    public class TestClass
    {
        public int TestPropertyInt { get; set; }
        public string TestPropertyString { get; set; }
    }
}
