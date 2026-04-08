using System.Diagnostics;

namespace PptxHeadlineNumbering;

public sealed class CliApplication
{
    private readonly TextWriter _stdout;
    private readonly TextWriter _stderr;
    private readonly InspectCommand _inspectCommand;
    private readonly ApplyCommand _applyCommand;

    public CliApplication(
        TextWriter stdout,
        TextWriter stderr,
        InspectCommand? inspectCommand = null,
        ApplyCommand? applyCommand = null)
    {
        _stdout = stdout ?? throw new ArgumentNullException(nameof(stdout));
        _stderr = stderr ?? throw new ArgumentNullException(nameof(stderr));
        _inspectCommand = inspectCommand ?? new InspectCommand(new SlideWalker(), _stdout);
        _applyCommand = applyCommand ?? new ApplyCommand(new SlideWalker(), new PrefixReplacer());
    }

    public int Run(string[] args)
    {
        try
        {
            ArgumentNullException.ThrowIfNull(args);

            if (args.Length == 0)
            {
                WriteUsage();
                return 1;
            }

            var command = args[0];
            if (string.Equals(command, "inspect", StringComparison.OrdinalIgnoreCase))
            {
                return RunInspect(args);
            }

            if (string.Equals(command, "apply", StringComparison.OrdinalIgnoreCase))
            {
                return RunApply(args);
            }

            _stderr.WriteLine($"Unknown command: {command}");
            WriteUsage();
            return 1;
        }
        catch (Exception ex)
        {
            Trace.WriteLine(ex.ToString());
            _stderr.WriteLine(ex.Message);
            return 1;
        }
    }

    private int RunInspect(string[] args)
    {
        if (args.Length != 2)
        {
            throw new ArgumentException("inspect requires <input.pptx>.");
        }

        _inspectCommand.Execute(args[1]);
        return 0;
    }

    private int RunApply(string[] args)
    {
        if (args.Length < 5)
        {
            throw new ArgumentException("apply requires <input.pptx> <output.pptx> --rule <rule.json>.");
        }

        var inputPath = args[1];
        var outputPath = args[2];
        string? rulePath = null;
        for (var i = 3; i < args.Length - 1; i++)
        {
            if (string.Equals(args[i], "--rule", StringComparison.OrdinalIgnoreCase))
            {
                rulePath = args[i + 1];
                break;
            }
        }

        if (string.IsNullOrWhiteSpace(rulePath))
        {
            throw new ArgumentException("--rule option is required.");
        }

        _applyCommand.Execute(inputPath, outputPath, rulePath);
        return 0;
    }

    private void WriteUsage()
    {
        _stderr.WriteLine("Usage:");
        _stderr.WriteLine("  pptx-headline-numbering inspect <input.pptx>");
        _stderr.WriteLine("  pptx-headline-numbering apply <input.pptx> <output.pptx> --rule <rule.json>");
    }
}
