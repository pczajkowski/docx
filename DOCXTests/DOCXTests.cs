using System.IO;
using Xunit;

namespace DOCXTests
{
    public class DocxTests
    {
        private static bool FileCompare(string file1, string file2)
        {
            // Borrowed from https://stackoverflow.com/questions/7931304/comparing-two-files-in-c-sharp#7931353

            int file1Byte;
            int file2Byte;

            // Determine if the same file was referenced two times.
            if (file1 == file2)
            {
                // Return true to indicate that the files are the same.
                return true;
            }

            // Open the two files.
            var fs1 = new FileStream(file1, FileMode.Open, FileAccess.Read);
            var fs2 = new FileStream(file2, FileMode.Open, FileAccess.Read);

            // Check the file sizes. If they are not the same, the files 
            // are not the same.
            if (fs1.Length != fs2.Length)
            {
                // Close the file
                fs1.Close();
                fs2.Close();

                // Return false to indicate files are different
                return false;
            }

            // Read and compare a byte from each file until either a
            // non-matching set of bytes is found or until the end of
            // file1 is reached.
            do
            {
                // Read one byte from each file.
                file1Byte = fs1.ReadByte();
                file2Byte = fs2.ReadByte();
            }
            while (file1Byte == file2Byte && file1Byte != -1);

            // Close the files.
            fs1.Close();
            fs2.Close();

            // Return the success of the comparison. "file1byte" is 
            // equal to "file2byte" at this point only if the files are 
            // the same.
            return file1Byte - file2Byte == 0;
        }

        [Fact]
        public void EnableTrackedChangesTest()
        {
            const string testFile = @"testFiles/notTracked.docx";
            const string expectedFile = @"testFiles/tracked.docx";

            using (var test = new DOCX.Docx(testFile))
            {
                var result = test.EnableTrackedChanges();
                Assert.True(result.status);
            }

            Assert.True(FileCompare(expectedFile, testFile));

            // Test for no duplication
            using (var test = new DOCX.Docx(expectedFile))
            {
                var result = test.EnableTrackedChanges();
                Assert.True(result.status);
            }

            Assert.True(FileCompare(expectedFile, testFile));
        }

        [Fact]
        public void AnonymizeAuthorsTest()
        {
            const string testFile = @"testFiles/testComments.docx";
            const string expectedFile = @"testFiles/testCommentsAnonymized.docx";

            using (var test = new DOCX.Docx(testFile))
            {
                var result = test.AnonymizeComments();
                Assert.True(result.status);
            }
            
            Assert.True(FileCompare(expectedFile, testFile));
        }
    }
}
