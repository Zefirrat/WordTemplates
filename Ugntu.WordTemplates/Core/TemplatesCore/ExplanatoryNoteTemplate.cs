using Ugntu.WordTemplates.Core.Engines;

namespace Ugntu.WordTemplates.Core;

public class ExplanatoryNoteTemplate(IDocumentEngine documentEngine) : TemplateBase(
        "Пояснительная записка",
        "poyasnitelnaya-zapiska-v3.docx.template", documentEngine)
{
    public override IEnumerable<TemplateParameter> TemplateParameters => new List<TemplateParameter>
    {
            new()
            {
                    Key = "Министерство",
                    Name = "Полное название министерства",
                    ExampleValue = "Министерство науки и высшего образования Российской Федерации"
            },
            new()
            {
                    Key = "Учреждение",
                    Name = "Полное название учебного учреждения",
                    ExampleValue =
                            "Федеральное государственное бюджетное образовательное учреждение высшего образования «Уфимский государственный нефтяной технический университет»"
            },
            new()
            {
                    Key = "Институт",
                    Name = "Название института",
                    ExampleValue = "Институт экосистем бизнеса и креативных индустрий"
            },
            new()
            {
                    Key = "Кафедра",
                    Name = "Название кафедры",
                    ExampleValue = "Кафедра «Проектный менеджмент и экономика предпринимательства»"
            },
            new()
            {
                    Key = "ЗаведующийКафедрой",
                    Name = "ФИО заведующего кафедрой",
                    ExampleValue = "Александров Р.Д."
            },
            new()
            {
                    Key = "ТемаРаботы",
                    Name = "Тема выпускной работы",
                    ExampleValue =
                            "ЦЕНОВАЯ ПОЛИТИКА ПРЕДПРИЯТИЯ В УСЛОВИЯХ НЕСОВЕРШЕННОЙ КОНКУРЕНЦИИ (НА ПРИМЕРЕ ООО «ДОБРЫЙ ПРОДУКТ»)"
            },
            new()
            {
                    Key = "ТипРаботы",
                    Name = "Тип квалификационной работы",
                    ExampleValue = "Выпускная квалификационная работа (бакалаврская работа)"
            },
            new()
            {
                    Key = "НаправлениеПодготовки",
                    Name = "Код и название направления подготовки",
                    ExampleValue = "38.03.01 Экономика"
            },
            new()
            {
                    Key = "Профиль",
                    Name = "Профиль подготовки",
                    ExampleValue = "Экономика предпринимательства и инноваций"
            },
            new()
            {
                    Key = "СтудентГруппа",
                    Name = "Группа студента",
                    ExampleValue = "БИЦсв-22-01"
            },
            new()
            {
                    Key = "ФИОСтудента",
                    Name = "ФИО студента",
                    ExampleValue = "Селиванов Р.М."
            },
            new()
            {
                    Key = "ДолжностьРуководитель",
                    Name = "Должность руководителя",
                    ExampleValue = "доц., канд. экон. наук"
            },
            new()
            {
                    Key = "Руководитель",
                    Name = "ФИО руководителя",
                    ExampleValue = "Лебедев Д.А."
            },
            new()
            {
                    Key = "ДолжностьНормоконтролер",
                    Name = "Должность нормоконтролёра",
                    ExampleValue = "доц., канд. экон. наук"
            },
            new()
            {
                    Key = "Нормоконтролер",
                    Name = "ФИО нормоконтролёра",
                    ExampleValue = "Кузьмина А.Д."
            },
            new()
            {
                    Key = "Город",
                    Name = "Город защиты",
                    ExampleValue = "Уфа"
            },
            new()
            {
                    Key = "Год",
                    Name = "Год защиты",
                    ExampleValue = "2024"
            },
            new()
            {
                    Key = "markdown_body",
                    Name = "Основной текст работы (markdown)",
                    ExampleValue = "Введение (markdown)"
            },
            new()
            {
                    Key = "markdown_body_2",
                    Name = "Дополнительный текст (markdown)",
                    ExampleValue = "Основной текст (markdown)"
            },
            new()
            {
                    Key = "markdown_summary",
                    Name = "Содержание работы (markdown)",
                    ExampleValue = "Содержание (markdown)"
            },
            new()
            {
                    Key = "markdown_summary_literature",
                    Name = "Список литературы (markdown)",
                    ExampleValue = "Список литературы (markdown)"
            }
    };
}