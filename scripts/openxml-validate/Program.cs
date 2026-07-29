// Validate Office documents with Microsoft's own Open XML SDK validator.
//
// This is a different question from the ECMA-376 schema check next door. That one validates
// individual XML parts against the published schemas, in isolation, and skips relationships
// and content types entirely. The SDK opens the package: it checks part structure and the
// relationships between parts as well as the markup, and it applies Office's own rules
// rather than only what the schema says.
//
// Prints one TAB-separated line per problem: file, part, description. The caller compares
// the output for a document against the output for the document it came from, so only a
// problem the input did not already have counts (real files draw complaints too).
//
// Usage: openxml-validate <file> [<file> ...]

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

if (args.Length == 0)
{
    Console.Error.WriteLine("usage: openxml-validate <file> [<file> ...]");
    return 2;
}

var validator = new OpenXmlValidator();
var exit = 0;

foreach (var path in args)
{
    OpenXmlPackage? package = null;
    try
    {
        package = Path.GetExtension(path).ToLowerInvariant() switch
        {
            ".docx" or ".docm" or ".dotx" => WordprocessingDocument.Open(path, false),
            ".xlsx" or ".xlsm" or ".xltx" => SpreadsheetDocument.Open(path, false),
            ".pptx" or ".pptm" or ".potx" => PresentationDocument.Open(path, false),
            _ => null,
        };
        if (package is null)
        {
            Console.Error.WriteLine($"skipping {path}: not an Open XML document");
            continue;
        }

        foreach (var error in validator.Validate(package))
        {
            // The node's own text is left out: it embeds document content, which would make
            // two files differ for reasons that have nothing to do with validity.
            var part = error.Part?.Uri?.ToString() ?? "(package)";
            var xpath = error.Path?.XPath ?? "";
            Console.WriteLine($"{path}\t{part}\t{error.ErrorType}\t{xpath}\t{error.Description}");
        }
    }
    catch (Exception ex)
    {
        // A document the SDK cannot even open is the worst outcome there is, so it is
        // reported rather than swallowed.
        Console.WriteLine($"{path}\t(package)\tOpenFailed\t\t{ex.GetType().Name}: {ex.Message}");
        exit = 1;
    }
    finally
    {
        package?.Dispose();
    }
}

return exit;
