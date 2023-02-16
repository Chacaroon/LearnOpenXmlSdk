using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace LearnOpenXml.Utils;

class Replacer
{
    public static void ReplaceKeys(OpenXmlElement element, IReadOnlyDictionary<string, string> keyValues, bool escapeKey = true)
    {
        foreach (var paragraph in element.Descendants<Paragraph>())
        {
            foreach (var (key, value) in keyValues)
            {
                var escapedKey = EscapeKeyIfNeeded(key, escapeKey);
                
                while (paragraph.InnerText.IndexOf(escapedKey, StringComparison.Ordinal) is var index && index > -1)
                {
                    ReplaceParagraph(paragraph, index, escapedKey, value);
                }
            }
        }
    }

    public static void ReplaceKey(OpenXmlElement element, string key, string value, bool escapeKey = true)
    {
        var escapedKey = EscapeKeyIfNeeded(key, escapeKey);
        foreach (var paragraph in element.Descendants<Paragraph>())
        {
            while (paragraph.InnerText.IndexOf(escapedKey, StringComparison.Ordinal) is var index && index > -1)
            {
                ReplaceParagraph(paragraph, index, escapedKey, value);
            }
        }
    }

    private static void ReplaceParagraph(OpenXmlElement paragraph, int index, string key, string value)
    {
        var count = 0;
        var builder = new StringBuilder();
        var textNodes = paragraph.Descendants<Text>().ToArray();

        if (textNodes.Length == 1)
        {
            textNodes[0].Text = textNodes[0].Text.Replace(key, value);
            return;
        }
        
        Text? textElement = null;
        foreach (var item in textNodes)
        {
            if (item.Text.Length == 0)
            {
                continue;
            }
            
            // Text in a paragraph can be split by pieces of type Text
            // Trying to join pieces into one item and replace the key in it
            if (count <= index && index + key.Length < count + item.Text.Length     // <text>some text {key to replace} some text</text> 
                || Between(index, count, count + item.Text.Length)                  // <text>some text {key to</text>
                || Between(index + key.Length, count, count + item.Text.Length)     // <text> replace} some text</text>
                || Between(count, index, index + key.Length))                       // <<text>some text {key </text><text>to rep</text><<text>lace} some text</text>
            {
                builder.Append(item.Text);

                textElement?.Remove();
                textElement = item;
            }
            
            count += item.Text.Length;
        }
        
        textElement!.Text = builder.Replace(key, value).ToString();
    }

    private static string EscapeKeyIfNeeded(string key, bool escape) => escape ? $"{{{key}}}" : key;

    private static bool Between(int number, int from, int to) => from <= number && number < to;
}