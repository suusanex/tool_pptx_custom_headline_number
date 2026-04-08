namespace PptxHeadlineNumbering;

public static class Program
{
    public static int Main(string[] args)
    {
        var application = new CliApplication(Console.Out, Console.Error);
        return application.Run(args);
    }
}
