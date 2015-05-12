using CommandLine;
using CommandLine.Text;

namespace ConfluenceWordDocImporter
{
    public enum Browser
    {
        Chrome,
        FireFox
    }

    public class Options
    {
        [Option('f', "file", Required = true, HelpText = "Path to the file or folder of files to import")]
        public string InputFileFolder { get; set; }

        [Option('u', "username", Required = true, HelpText = "Username to use when accessing Confluence")]
        public string Username { get; set; }

        [Option('w', "password", Required = true, HelpText = "Password to use when accessing Confluence")]
        public string Password { get; set; }

        [Option('p', "parentPage", Required = true, HelpText = "Url to the page that should act as parent page for the new page")]
        public string ParentPage { get; set; }

        [Option('b', "browser", DefaultValue = Browser.Chrome, HelpText = "Default Chrome, but can be set to FireFox if that is installed")]
        public Browser Browser { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            var help = new HelpText
            {
                Heading = new HeadingInfo("ConfluenceWordDocImporter", "0.1"),
                Copyright = new CopyrightInfo("Martin Jæger", 2015),
                AdditionalNewLineAfterOption = true,
                AddDashesToOption = true
            };
            help.AddOptions(this);
            return help;
        }
    }
}
