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
            var actualPropertyInfo = CommonUtilities.TryGetPropertyFromExpression<TestClass>(x => x.TestPropertyInt);
            var actualPropertyNameIsEqualToExpectedPropertyName = actualPropertyInfo.Exists(x => x.Name == expectedPropertyName);

            // ASSERT
            Assert.True(actualPropertyNameIsEqualToExpectedPropertyName);
        }
    }

    public class TestClass
    {
        public int TestPropertyInt { get; set; }
        public string TestPropertyString { get; set; }
    }
}
