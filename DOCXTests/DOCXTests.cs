using System.IO.Compression;
using Xunit;

namespace DOCXTests
{
    public class DocxTests
    {
        [Fact]
        public void EnableTrackedChangesTest()
        {
            const string testFile = @"testFiles/notTracked.docx";

            using (var zip = ZipFile.Open(testFile, ZipArchiveMode.Update))
            {
                using (var test = new DOCX.Docx(zip))
                {
                    var result = test.EnableTrackedChanges();
                    Assert.True(result.status);
                }
            }
        }

        [Fact]
        public void AnonymizeAndDeanonymizeCommentsTest()
        {
            const string testFile = @"testFiles/testComments.docx";

            using (var test = new DOCX.Docx(testFile))
            {
                var result = test.AnonymizeComments();
                Assert.True(result.status);
            }
                       
            using (var test = new DOCX.Docx(testFile))
            {
                var result = test.DeanonymizeComments();
                Assert.True(result.status);
            }
        }
    }
}
