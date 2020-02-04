using CoreUtils.OpenXmlService;
using FluentAssertions;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Test
{
    public class ExcelDocument_should_
    {
        [Test]
        public void build_a_sheet_from_a_list()
        {
            var document = new ExcelDocument();
            var expected =new List<List<string>> { 
                new List<string>{ "a1", "b1", "c1" },
                new List<string> { "1", "2", "3" },
                new List<string> { DateTime.Now.ToString(CultureInfo.InvariantCulture), "False", "True" }
            };

            document.AddList(expected, "Sheet1",CultureInfo.InvariantCulture);
            var actual = document.GetRange("Sheet1", "A1", "C3",CultureInfo.InvariantCulture);
            actual.Should().BeEquivalentTo(expected);
        }
    }
}
