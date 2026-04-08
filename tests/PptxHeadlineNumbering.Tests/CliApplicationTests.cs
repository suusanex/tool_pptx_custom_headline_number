namespace PptxHeadlineNumbering.Tests;

public class CliApplicationTests
{
    [Test]
    public void Run_ReturnsErrorForMissingArguments()
    {
        var stdout = new StringWriter();
        var stderr = new StringWriter();
        var application = new CliApplication(stdout, stderr);

        var exitCode = application.Run([]);

        Assert.Multiple(() =>
        {
            Assert.That(exitCode, Is.EqualTo(1));
            Assert.That(stderr.ToString(), Does.Contain("Usage:"));
        });
    }

    [Test]
    public void Run_ReturnsErrorForMissingInputFile()
    {
        var stdout = new StringWriter();
        var stderr = new StringWriter();
        var application = new CliApplication(stdout, stderr);

        var exitCode = application.Run(["inspect", "not-found.pptx"]);

        Assert.Multiple(() =>
        {
            Assert.That(exitCode, Is.EqualTo(1));
            Assert.That(stderr.ToString(), Does.Contain("not-found.pptx"));
        });
    }

    [Test]
    public void Run_ReturnsErrorWhenRuleOptionMissing()
    {
        var stdout = new StringWriter();
        var stderr = new StringWriter();
        var application = new CliApplication(stdout, stderr);

        var exitCode = application.Run(["apply", "in.pptx", "out.pptx"]);

        Assert.Multiple(() =>
        {
            Assert.That(exitCode, Is.EqualTo(1));
            Assert.That(stderr.ToString(), Does.Contain("apply requires"));
        });
    }

    [Test]
    public void UT_IT_070__Run_ReturnsErrorForUnknownCommand()
    {
        // 不明なサブコマンドを指定した場合、終了コード 1 でエラーメッセージが出力される（TP-070）
        var stdout = new StringWriter();
        var stderr = new StringWriter();
        var application = new CliApplication(stdout, stderr);

        var exitCode = application.Run(["unknowncmd"]);

        Assert.Multiple(() =>
        {
            Assert.That(exitCode, Is.EqualTo(1));
            Assert.That(stderr.ToString(), Does.Contain("Unknown command"));
        });
    }

    [Test]
    public void UT_IT_070__Run_ReturnsErrorWhenNormalizedApplyPathsAreSame()
    {
        // "./input.pptx" と "input.pptx" は Path.GetFullPath で同じ絶対パスに正規化されるため
        // 入出力パス同一と判定されて ArgumentException → 終了コード 1 になる（TP-070）
        var stdout = new StringWriter();
        var stderr = new StringWriter();
        var application = new CliApplication(stdout, stderr);

        var exitCode = application.Run(["apply", "input.pptx", "./input.pptx", "--rule", "rule.json"]);

        Assert.Multiple(() =>
        {
            Assert.That(exitCode, Is.EqualTo(1));
            Assert.That(stderr.ToString(), Does.Contain("different"));
        });
    }

    [Test]
    public void UT_IT_060__Run_ReturnsErrorWhenApplyInputNotFound()
    {
        // apply で存在しない入力ファイルを指定した場合、終了コード 1 になる（TP-060）
        var tempDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var rulePath = Path.Combine(tempDir, "rule.json");
        File.WriteAllText(
            rulePath,
            """
            {
              "prefixRegex":"^[^\\s\\u3000]+(?:[\\s\\u3000]+)?",
              "separator":" ",
              "insertWhenPrefixMissing":true,
              "levels":[{"name":"H1","match":{"placeholderTypes":["title"]},"format":"{H1}.","resetsOnNewLevel":[]}]
            }
            """);

        var stdout = new StringWriter();
        var stderr = new StringWriter();
        var application = new CliApplication(stdout, stderr);

        try
        {
            var exitCode = application.Run(
                ["apply",
                    Path.Combine(tempDir, "nonexistent.pptx"),
                    Path.Combine(tempDir, "out.pptx"),
                    "--rule", rulePath]);

            Assert.Multiple(() =>
            {
                Assert.That(exitCode, Is.EqualTo(1));
                Assert.That(stderr.ToString(), Is.Not.Empty);
            });
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }

    [Test]
    public void UT_IT_060__Run_ReturnsErrorForCorruptPptxFile()
    {
        // 壊れた .pptx（非 ZIP）を inspect した場合、終了コード 1 になる（TP-060）
        var tempDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var filePath = Path.Combine(tempDir, "corrupt.pptx");
        File.WriteAllBytes(filePath, [0x00, 0x01, 0x02, 0x03]); // ゴミバイト列（ZIP でない）

        var stdout = new StringWriter();
        var stderr = new StringWriter();
        var application = new CliApplication(stdout, stderr);

        try
        {
            var exitCode = application.Run(["inspect", filePath]);
            Assert.That(exitCode, Is.EqualTo(1));
        }
        finally
        {
            Directory.Delete(tempDir, true);
        }
    }
}
