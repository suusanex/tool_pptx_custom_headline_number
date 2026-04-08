namespace PptxHeadlineNumbering.Tests;

public class HeadingCounterTests
{
    [Test]
    public void Increment_ResetsDependentCounters()
    {
        var counter = new HeadingCounter(CreateLevels());

        counter.Increment("H1");
        counter.Increment("H2");
        counter.Increment("H3");
        counter.Increment("H1");

        Assert.Multiple(() =>
        {
            Assert.That(counter.GetCount("H1"), Is.EqualTo(2));
            Assert.That(counter.GetCount("H2"), Is.EqualTo(0));
            Assert.That(counter.GetCount("H3"), Is.EqualTo(0));
        });

        counter.Increment("H2");
        counter.Increment("H3");
        Assert.That(counter.Format("{H1}.{H2}-{H3}"), Is.EqualTo("2.1-1"));
    }

    [Test]
    public void Format_ThrowsWhenUnknownCounterExists()
    {
        var counter = new HeadingCounter(CreateLevels());
        counter.Increment("H1");

        var exception = Assert.Throws<KeyNotFoundException>(() => counter.Format("{H1}.{Unknown}"));
        Assert.That(exception!.Message, Does.Contain("Unknown"));
    }

    [Test]
    public void UT_IT_030__H2IncrementResetsH3()
    {
        var counter = new HeadingCounter(CreateLevels());

        counter.Increment("H1"); // H1=1
        counter.Increment("H2"); // H2=1
        counter.Increment("H3"); // H3=1
        counter.Increment("H3"); // H3=2
        counter.Increment("H2"); // H2=2, H3 は 0 にリセット

        Assert.Multiple(() =>
        {
            Assert.That(counter.GetCount("H1"), Is.EqualTo(1));
            Assert.That(counter.GetCount("H2"), Is.EqualTo(2));
            Assert.That(counter.GetCount("H3"), Is.EqualTo(0));
        });

        counter.Increment("H3");
        Assert.That(counter.Format("{H1}.{H2}-{H3}"), Is.EqualTo("1.2-1"));
    }

    [Test]
    public void UT_IT_030__H1OnlyConsecutiveMonotonicallyIncreases()
    {
        var counter = new HeadingCounter(CreateLevels());

        counter.Increment("H1");
        counter.Increment("H1");
        counter.Increment("H1");

        Assert.Multiple(() =>
        {
            Assert.That(counter.GetCount("H1"), Is.EqualTo(3));
            Assert.That(counter.GetCount("H2"), Is.EqualTo(0));
            Assert.That(counter.GetCount("H3"), Is.EqualTo(0));
        });
    }

    [Test]
    public void UT_IT_030__H3OnlyConsecutiveMonotonicallyIncreases()
    {
        var counter = new HeadingCounter(CreateLevels());

        counter.Increment("H3");
        counter.Increment("H3");
        counter.Increment("H3");

        // H3 だけインクリメントしても他のカウンタはリセットされない
        Assert.Multiple(() =>
        {
            Assert.That(counter.GetCount("H3"), Is.EqualTo(3));
            Assert.That(counter.GetCount("H1"), Is.EqualTo(0));
            Assert.That(counter.GetCount("H2"), Is.EqualTo(0));
        });
    }

    [Test]
    public void UT_IT_030__ConstructorThrowsWhenLevelsEmpty()
    {
        Assert.Throws<ArgumentException>(() => new HeadingCounter(Array.Empty<NumberingLevelRule>()));
    }

    private static IReadOnlyList<NumberingLevelRule> CreateLevels()
    {
        return
        [
            new NumberingLevelRule
            {
                Name = "H1",
                Match = new NumberingMatchRule { PlaceholderTypes = ["title"] },
                Format = "{H1}.",
                ResetsOnNewLevel = []
            },
            new NumberingLevelRule
            {
                Name = "H2",
                Match = new NumberingMatchRule { PlaceholderTypes = ["body"], ParagraphLevel = 0 },
                Format = "{H1}.{H2}",
                ResetsOnNewLevel = ["H1"]
            },
            new NumberingLevelRule
            {
                Name = "H3",
                Match = new NumberingMatchRule { PlaceholderTypes = ["body"], ParagraphLevel = 1 },
                Format = "{H3})",
                ResetsOnNewLevel = ["H1", "H2"]
            },
        ];
    }
}
