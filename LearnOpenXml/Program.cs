using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using LearnOpenXml.Utils;

var pairs = new Dictionary<string, string>()
{
    { "test", "another text" }
};

await using var original = typeof(Replacer).Assembly.GetManifestResourceStream("LearnOpenXml.test.docx")!;
await using var modified = new FileStream("C:\\Projects\\Chacaroon\\LearnOpenXml\\LearnOpenXml\\test_modified.docx", FileMode.Create, FileAccess.ReadWrite);
await original.CopyToAsync(modified);

using var wordProcessingDocument = WordprocessingDocument.Open(modified, isEditable: true);

var body = wordProcessingDocument.MainDocumentPart!.Document.Body!;

Replacer.ReplaceKeys(wordProcessingDocument.MainDocumentPart!.Document.Body!, pairs);
Cleaner.RemoveEmptyTableRows(body);

var logoPlaceholder = body.Descendants<Paragraph>().FirstOrDefault(x => x.InnerText == "{logo}");

if (logoPlaceholder == null)
{
    throw new Exception("Can't insert new logo");
}

await using var newLogo = typeof(ImageReplacer).Assembly.GetManifestResourceStream("LearnOpenXml.new-logo.png")!;
ImageReplacer.InsertPicture(wordProcessingDocument, logoPlaceholder, newLogo);

Console.WriteLine("Done");