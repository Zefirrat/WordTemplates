using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace Ugntu.WordTemplates.Core.Core.Engines;

class MarkdownToWordConverter
{
    public void ConvertMarkdownToWord(string markdownText, Application wordApp, Document wordDoc, string replacedWordKey)
    {
        try
        {
            DetectLanguageAndSetStyles(wordApp);
            
            // Ищем маркер "#{markdown_body}" в документе и заменяем его на содержимое Markdown
            var placeholder = $"#{{{replacedWordKey}}}";
            Range range = FindPlaceholder(wordDoc, placeholder);
            if (range != null)
            {
                InsertMarkdownToWordDocument(wordDoc, markdownText, range);
            }
            else
            {
                Console.WriteLine($"Маркер {placeholder} не найден в документе.");
                throw new Exception($"Маркер {placeholder} не найден в документе.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Ошибка: " + ex.Message);
            throw;
        }
    }

    private Range FindPlaceholder(Document wordDoc, string placeholder)
    {
        // Поиск маркера в документе
        Range range = wordDoc.Content;
        Find find = range.Find;
        find.Text = placeholder;

        if (find.Execute())
        {
            return range; // Возвращаем диапазон, если маркер найден
        }

        return null; // Возвращаем null, если маркер не найден
    }

    private void InsertMarkdownToWordDocument(Document wordDoc, string markdownText, Range range)
    {
        // Разделяем Markdown на строки
        string[] lines = markdownText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

        foreach (var line in lines)
        {
            // Проверяем на заголовки (например, ## Heading 2)
            if (Regex.IsMatch(line, @"^#{1,6}\s"))
            {
                int headingLevel = line.TakeWhile(c => c == '#').Count(); // Определяем уровень заголовка
                string headingText = line.Substring(headingLevel).Trim(); // Получаем текст заголовка

                InsertHeading(wordDoc, headingText, headingLevel, range); // Вставляем заголовок
            }
            // Проверяем на нумерованные и маркерованные списки
            else if (Regex.IsMatch(line, @"^\d+\.\s"))
            {
                string listItemText = line.Substring(line.IndexOf(' ') + 1);
                InsertListItem(wordDoc, listItemText, isNumbered: true, range);
            }
            else if (Regex.IsMatch(line, @"^[-*]\s"))
            {
                string listItemText = line.Substring(2).Trim();
                InsertListItem(wordDoc, listItemText, isNumbered: false, range);
            }
            // Проверяем на обычный текст или параграф
            else
            {
                InsertParagraph(wordDoc, line, range);
            }
        }
    }

    private void InsertHeading(Document wordDoc, string text, int level, Range range)
    {
        range.Text = text;

        // Применение стиля заголовка в зависимости от уровня
        switch (level)
        {
            case 1:
                range.set_Style("Заголовок 1");
                break;
            case 2:
                range.set_Style("Заголовок 2");
                break;
            case 3:
                range.set_Style("Заголовок 3");
                break;
            case 4:
                range.set_Style("Заголовок 4");
                break;
            case 5:
                range.set_Style("Заголовок 5");
                break;
            case 6:
                range.set_Style("Заголовок 6");
                break;
        }

        range.InsertParagraphAfter();
        range.Collapse(WdCollapseDirection.wdCollapseEnd);
    }

    private void InsertParagraph(Document wordDoc, string text, Range range)
    {
        range.Text = text;
        range.set_Style("Обычный"); // Применяем стиль обычного текста
        range.InsertParagraphAfter();
        range.Collapse(WdCollapseDirection.wdCollapseEnd);
    }

    private void InsertListItem(Document wordDoc, string text, bool isNumbered, Range range)
    {
        range.Text = text;

        if (isNumbered)
        {
            range.set_Style("Нумерованный список");
        }
        else
        {
            range.set_Style("Маркированный список");
        }

        range.InsertParagraphAfter();
        range.Collapse(WdCollapseDirection.wdCollapseEnd);
    }

    private string heading1Style;
    private string heading2Style;
    private string heading3Style;
    private string heading4Style;
    private string heading5Style;
    private string heading6Style;
    private string normalStyle;
    private string bulletedListStyle;
    private string numberedListStyle;

    private void DetectLanguageAndSetStyles(Application wordApp)
    {
        int lcid = (int)wordApp.Language; // Определяем текущую локаль Word

        if (lcid == 1049) // Русский язык (LCID 1049)
        {
            heading1Style = "Заголовок 1";
            heading2Style = "Заголовок 2";
            heading3Style = "Заголовок 3";
            heading4Style = "Заголовок 4";
            heading5Style = "Заголовок 5";
            heading6Style = "Заголовок 6";
            normalStyle = "Обычный";
            bulletedListStyle = "Маркированный список";
            numberedListStyle = "Нумерованный список";
        }
        else // Английская локаль или другая (по умолчанию английские стили)
        {
            heading1Style = "Heading 1";
            heading2Style = "Heading 2";
            heading3Style = "Heading 3";
            heading4Style = "Heading 4";
            heading5Style = "Heading 5";
            heading6Style = "Heading 6";
            normalStyle = "Normal";
            bulletedListStyle = "Bulleted List";
            numberedListStyle = "Numbered List";
        }
    }
}