# Confluence Word Doc Importer #
**Why use?**

You LOVE Confluence, and you would like to transform your Word documents into Confluence pages. So you go looking for a bulk import feature.....but it's not there. So either you would have to import each and everyone of them, by hand, or simply give up. **But fear not!**

This nifty little .NET console application, will import your Word documents into Confluence as pages.

Here are some examples:

**Importing a single file**


**Importing all doc and docx files from a folder** (not recursive)


**How is it done?**

This tool uses Selenium WebDriver to drive either Chrome or FireFox, to do what a user would have done, when importing Word documents into Confluence.

**What you will need**

- A windows machine to run the program from
- A Confluence instance
- .NET Framework 4.5
- Chrome or FireFox installed

**This is what the --help switch prints out**

  -f, --file          Required. Path to the file or folder of files to import

  -u, --username      Required. Username to use when accessing Confluence

  -w, --password      Required. Password to use when accessing Confluence

  -p, --parentPage    Required. Url to the page that should act as parent page
                      for the new page

  -b, --browser       (Default: Chrome) Default Chrome, but can be set to
                      FireFox if that is installed

  --help              Display this help screen.
