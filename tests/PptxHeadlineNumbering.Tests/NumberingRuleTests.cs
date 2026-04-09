using System.IO;
using System.Text.Json;
using DocumentFormat.OpenXml.Presentation;

namespace PptxHeadlineNumbering.Tests;

public class NumberingRuleTests
{
    [Test]
    public void LoadFromFile_LoadsValidJson()
    {
        var path = Path.Combine(TestContext.CurrentContext.WorkDirectory, $"{Guid.NewGuid():N}.json");
        File.WriteAllText(
            path,
            """
            {
              "prefixRegex":"^(?:\\d+(?:\\.\\d+)*[.)]?)(?:[\\s\\u3000]+)?",
              "separator":" ",
              "insertWhenPrefixMissing":true,
              "excludedSlideRanges":[{"startSlideNumber":1,"endSlideNumber":2},{"startSlideNumber":10,"endSlideNumber":10}],
              "levels":[
                {"name":"H1","match":{"placeholderTypes":["title","ctrTitle"]},"format":"{H1}.","resetsOnNewLevel":[]},
                {"name":"H2","matches":[{"placeholderTypes":["body","obj"],"paragraphLevel":0},{"shapeNames":["Content Placeholder 2"],"paragraphLevel":0}],"format":"{H1}.{H2}","resetsOnNewLevel":["H1"]}
              ]
            }
            """);

        try
        {
            var rule = NumberingRule.LoadFromFile(path);

            Assert.Multiple(() =>
            {
                Assert.That(rule.Levels, Has.Count.EqualTo(2));
                Assert.That(rule.ExcludedSlideRanges, Has.Count.EqualTo(2));
                Assert.That(rule.Separator, Is.EqualTo(" "));
                Assert.That(rule.InsertWhenPrefixMissing, Is.True);
                Assert.That(rule.BuildPrefixRegex().ToString(), Is.EqualTo("^(?:\\d+(?:\\.\\d+)*[.)]?)(?:[\\s\\u3000]+)?"));
            });
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public void LoadFromFile_ThrowsForBrokenJson()
    {
        var path = Path.Combine(TestContext.CurrentContext.WorkDirectory, $"{Guid.NewGuid():N}.json");
        File.WriteAllText(path, "{");

        try
        {
            Assert.Throws<JsonException>(() => NumberingRule.LoadFromFile(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public void UT_IT_080__LoadFromFile_ThrowsForEmptyFile()
    {
        var path = Path.Combine(TestContext.CurrentContext.WorkDirectory, $"{Guid.NewGuid():N}.json");
        File.WriteAllText(path, string.Empty);

        try
        {
            Assert.Throws<JsonException>(() => NumberingRule.LoadFromFile(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public void UT_IT_080__LoadFromFile_ThrowsForEmptyLevels()
    {
        var path = Path.Combine(TestContext.CurrentContext.WorkDirectory, $"{Guid.NewGuid():N}.json");
        File.WriteAllText(path, """{"levels":[]}""");

        try
        {
            Assert.Throws<InvalidDataException>(() => NumberingRule.LoadFromFile(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public void UT_IT_080__LoadFromFile_ThrowsForInvalidPrefixRegex()
    {
        var path = Path.Combine(TestContext.CurrentContext.WorkDirectory, $"{Guid.NewGuid():N}.json");
        File.WriteAllText(
            path,
            """
            {
              "prefixRegex":"[invalid",
              "levels":[{"name":"H1","match":{"placeholderTypes":["title"]},"format":"{H1}.","resetsOnNewLevel":[]}]
            }
            """);

        try
        {
            // RegexParseException は ArgumentException のサブクラスのため Assert.Catch を使用
            Assert.Catch<ArgumentException>(() => NumberingRule.LoadFromFile(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public void UT_IT_080__LoadFromFile_ThrowsWhenFormatMissing()
    {
        var path = Path.Combine(TestContext.CurrentContext.WorkDirectory, $"{Guid.NewGuid():N}.json");
        // "format" フィールドが省略されたとき、Validate が InvalidDataException をスローする
        File.WriteAllText(
            path,
            """
            {
              "levels":[{"name":"H1","match":{"placeholderTypes":["title"]},"resetsOnNewLevel":[]}]
            }
            """);

        try
        {
            Assert.Throws<InvalidDataException>(() => NumberingRule.LoadFromFile(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public void UT_IT_080__LoadFromFile_AllowsExplicitEmptyFormat()
    {
        // format:"" は prefix 削除の明示的指定として有効であり、Validate を通過する（TP-080）
        var path = Path.Combine(TestContext.CurrentContext.WorkDirectory, $"{Guid.NewGuid():N}.json");
        File.WriteAllText(
            path,
            """
            {
              "prefixRegex":"^(?:\\d+(?:\\.\\d+)*[.)]?)(?:[\\s\\u3000]+)?",
              "levels":[{"name":"H1","match":{"placeholderTypes":["title"]},"format":"","resetsOnNewLevel":[]}]
            }
            """);

        try
        {
            var rule = NumberingRule.LoadFromFile(path);
            Assert.That(rule.Levels[0].Format, Is.EqualTo(string.Empty));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestCase(" ")]
    [TestCase("\u3000")]
    public void UT_IT_080__LoadFromFile_ThrowsForWhitespaceOnlyFormat(string format)
    {
        // format に空白のみを指定した場合は曖昧な設定として拒否される（TP-080）
        var path = Path.Combine(TestContext.CurrentContext.WorkDirectory, $"{Guid.NewGuid():N}.json");
        File.WriteAllText(
            path,
            $$"""
            {
              "prefixRegex":"^(?:\\d+(?:\\.\\d+)*[.)]?)(?:[\\s\\u3000]+)?",
              "levels":[{"name":"H1","match":{"placeholderTypes":["title"]},"format":"{{format}}","resetsOnNewLevel":[]}]
            }
            """);

        try
        {
            Assert.Throws<InvalidDataException>(() => NumberingRule.LoadFromFile(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public void LoadFromFile_AllowsExplicitEmptyFormat()
    {
        var path = Path.Combine(TestContext.CurrentContext.WorkDirectory, $"{Guid.NewGuid():N}.json");
        File.WriteAllText(
            path,
            """
            {
              "prefixRegex":"^(?:\\d+(?:\\.\\d+)*[.)]?)(?:[\\s\\u3000]+)?",
              "levels":[{"name":"H1","match":{"placeholderTypes":["title"]},"format":"","resetsOnNewLevel":[]}]
            }
            """);

        try
        {
            var rule = NumberingRule.LoadFromFile(path);
            Assert.That(rule.Levels[0].Format, Is.EqualTo(string.Empty));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestCase(" ")]
    [TestCase("\u3000")]
    public void LoadFromFile_ThrowsWhenFormatIsWhitespaceOnly(string format)
    {
        var path = Path.Combine(TestContext.CurrentContext.WorkDirectory, $"{Guid.NewGuid():N}.json");
        File.WriteAllText(
            path,
            $$"""
            {
              "prefixRegex":"^(?:\\d+(?:\\.\\d+)*[.)]?)(?:[\\s\\u3000]+)?",
              "levels":[{"name":"H1","match":{"placeholderTypes":["title"]},"format":"{{format}}","resetsOnNewLevel":[]}]
            }
            """);

        try
        {
            Assert.Throws<InvalidDataException>(() => NumberingRule.LoadFromFile(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public void MatchLevel_MatchesShapeNameWhenPlaceholderTypeIsMissing()
    {
        var rule = new NumberingRule
        {
            Levels =
            [
                new NumberingLevelRule
                {
                    Name = "H2",
                    Match = new NumberingMatchRule
                    {
                        ShapeNames = ["コンテンツ プレースホルダー 2"],
                        ParagraphLevel = 0
                    },
                    Format = "{H1}.{H2}",
                    ResetsOnNewLevel = ["H1"]
                },
                new NumberingLevelRule
                {
                    Name = "H1",
                    Match = new NumberingMatchRule
                    {
                        PlaceholderTypes = ["title"]
                    },
                    Format = "{H1}.",
                    ResetsOnNewLevel = []
                }
            ]
        };

        var matched = rule.MatchLevel(null, "コンテンツ プレースホルダー 2", 0);

        Assert.That(matched?.Name, Is.EqualTo("H2"));
    }

    [Test]
    public void MatchLevel_SupportsMatchesArrayAsOrCondition()
    {
        var rule = new NumberingRule
        {
            Levels =
            [
                new NumberingLevelRule
                {
                    Name = "H2",
                    Matches =
                    [
                        new NumberingMatchRule
                        {
                            PlaceholderTypes = ["body"],
                            ParagraphLevel = 0
                        },
                        new NumberingMatchRule
                        {
                            ShapeNames = ["字幕 2"],
                            ParagraphLevel = 0
                        }
                    ],
                    Format = "{H1}.{H2}",
                    ResetsOnNewLevel = ["H1"]
                },
                new NumberingLevelRule
                {
                    Name = "H1",
                    Match = new NumberingMatchRule
                    {
                        PlaceholderTypes = ["title"]
                    },
                    Format = "{H1}.",
                    ResetsOnNewLevel = []
                }
            ]
        };

        Assert.Multiple(() =>
        {
            Assert.That(rule.MatchLevel(PlaceholderValues.Body, "Body 1", 0)?.Name, Is.EqualTo("H2"));
            Assert.That(rule.MatchLevel(null, "字幕 2", 0)?.Name, Is.EqualTo("H2"));
        });
    }

    [Test]
    public void IsExcludedSlide_UsesOneBasedInclusiveRanges()
    {
        var rule = new NumberingRule
        {
            Levels =
            [
                new NumberingLevelRule
                {
                    Name = "H1",
                    Match = new NumberingMatchRule
                    {
                        PlaceholderTypes = ["title"]
                    },
                    Format = "{H1}.",
                    ResetsOnNewLevel = []
                }
            ],
            ExcludedSlideRanges =
            [
                new ExcludedSlideRange { StartSlideNumber = 1, EndSlideNumber = 2 },
                new ExcludedSlideRange { StartSlideNumber = 5, EndSlideNumber = 5 }
            ]
        };

        Assert.Multiple(() =>
        {
            Assert.That(rule.IsExcludedSlide(0), Is.True);
            Assert.That(rule.IsExcludedSlide(1), Is.True);
            Assert.That(rule.IsExcludedSlide(2), Is.False);
            Assert.That(rule.IsExcludedSlide(4), Is.True);
        });
    }

    [Test]
    public void UT_IT_080__LoadFromFile_ThrowsForInvalidExcludedSlideRange()
    {
        var path = Path.Combine(TestContext.CurrentContext.WorkDirectory, $"{Guid.NewGuid():N}.json");
        File.WriteAllText(
            path,
            """
            {
              "excludedSlideRanges":[{"startSlideNumber":3,"endSlideNumber":2}],
              "levels":[{"name":"H1","match":{"placeholderTypes":["title"]},"format":"{H1}.","resetsOnNewLevel":[]}]
            }
            """);

        try
        {
            Assert.Throws<InvalidDataException>(() => NumberingRule.LoadFromFile(path));
        }
        finally
        {
            File.Delete(path);
        }
    }
}
