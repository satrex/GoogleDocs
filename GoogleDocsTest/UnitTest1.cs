using System;
using Xunit;

namespace GoogleDocsTest
{
    public class UnitTest1
    {
        [Fact]
        public void Test1()
        {
            const string fileId =  @"1TYgZoEtfvBibpdNuWKZV6varVDASEdTTtTC1AEtolsY";
            var title = Satrex.GoogleDocs.GoogleDocsInternal.GetTitle(fileId);
            Assert.NotEmpty(title);
            Assert.Equal("366_半夏厚朴湯", title );
        }
    }
}
