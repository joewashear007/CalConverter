using CalConverter.Lib.Parsers;

namespace CalConverter.Lib.Tests;

[TestClass]
public class ParserTests
{
    [TestMethod]
    public void ParserSampleDocument()
    {
        string file = "Preceptor Schedule.xlsx";
        Parser parser = new Parser();

        var items = parser.ProcessFile(file, "Sheet1");

        foreach ( var schedule in items)
        {
            Console.WriteLine(schedule);
        }

        Assert.IsTrue(items.Any());
    }
}