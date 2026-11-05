// ooxml-validate — validate .pptx / .docx / .xlsx files with Microsoft's Open XML SDK,
// the same rule set Office applies. Exit code: 0 = valid, 1 = errors found, 2 = usage error.
//
//   ooxml-validate deck.pptx
//   ooxml-validate --json deck.pptx report.docx
//   ooxml-validate --format Office2016 deck.pptx

using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

var json = false;
var format = FileFormatVersions.Microsoft365;
var maxErrors = 1000;
var files = new List<string>();

for (var i = 0; i < args.Length; i++)
{
    switch (args[i])
    {
        case "--json": json = true; break;
        case "--format" when i + 1 < args.Length:
            if (!Enum.TryParse(args[++i], ignoreCase: true, out format))
                return Usage($"unknown format '{args[i]}'. Use one of: {string.Join(", ", Enum.GetNames<FileFormatVersions>())}");
            break;
        case "--max" when i + 1 < args.Length:
            if (!int.TryParse(args[++i], out maxErrors) || maxErrors < 1) return Usage("--max must be a positive number");
            break;
        case "-h" or "--help": return Usage(null);
        default:
            if (args[i].StartsWith('-')) return Usage($"unknown option '{args[i]}'");
            files.Add(args[i]);
            break;
    }
}
if (files.Count == 0) return Usage("no files given");

var results = new List<FileResult>();
foreach (var path in files)
    results.Add(ValidateFile(path, format, maxErrors));

if (json)
{
    Console.WriteLine(JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true }));
}
else
{
    foreach (var r in results)
    {
        if (r.Errors.Count == 0)
        {
            Console.WriteLine($"PASS  {r.File}");
            continue;
        }
        Console.WriteLine($"FAIL  {r.File}  ({r.Errors.Count} error(s))");
        foreach (var e in r.Errors)
        {
            Console.WriteLine($"  [{e.Kind}] {e.Part}");
            if (!string.IsNullOrEmpty(e.XPath)) Console.WriteLine($"      at {e.XPath}");
            if (!string.IsNullOrEmpty(e.Element)) Console.WriteLine($"      element {e.Element}");
            Console.WriteLine($"      {e.Id}: {e.Description}");
        }
    }
}

return results.Any(r => r.Errors.Count > 0) ? 1 : 0;

static FileResult ValidateFile(string path, FileFormatVersions format, int maxErrors)
{
    var result = new FileResult(path, new List<ErrorInfo>());
    if (!File.Exists(path))
    {
        result.Errors.Add(new ErrorInfo("Open", "", "", "", "FileNotFound", $"File not found: {path}"));
        return result;
    }

    try
    {
        using OpenXmlPackage doc = Path.GetExtension(path).ToLowerInvariant() switch
        {
            ".pptx" or ".potx" or ".pptm" => PresentationDocument.Open(path, false),
            ".docx" or ".dotx" or ".docm" => WordprocessingDocument.Open(path, false),
            ".xlsx" or ".xltx" or ".xlsm" => SpreadsheetDocument.Open(path, false),
            var ext => throw new NotSupportedException($"Unsupported file type '{ext}'"),
        };

        var validator = new OpenXmlValidator(format) { MaxNumberOfErrors = maxErrors };
        foreach (var e in validator.Validate(doc))
        {
            result.Errors.Add(new ErrorInfo(
                Kind: e.ErrorType.ToString(),
                Part: e.Part?.Uri.ToString() ?? "",
                XPath: e.Path?.XPath ?? "",
                Element: e.Node?.LocalName ?? "",
                Id: e.Id ?? "",
                Description: e.Description ?? ""));
        }
    }
    catch (Exception ex)
    {
        // A package the SDK cannot even open is as broken as it gets.
        result.Errors.Add(new ErrorInfo("Open", "", "", "", ex.GetType().Name, ex.Message));
    }
    return result;
}

static int Usage(string? problem)
{
    if (problem is not null) Console.Error.WriteLine($"error: {problem}");
    Console.Error.WriteLine("usage: ooxml-validate [--json] [--format Microsoft365|Office2019|Office2016|...] [--max N] <file> [file ...]");
    return 2;
}

record FileResult(string File, List<ErrorInfo> Errors);
record ErrorInfo(string Kind, string Part, string XPath, string Element, string Id, string Description);
