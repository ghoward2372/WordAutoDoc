using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;

public class NumberingHelper
{



    // Helper method to try to get a bullet numbering id from the style definitions.
    public int GetBulletNumIdFromStyle(MainDocumentPart mainPart, string styleId = "List Bullet")
    {
        if (mainPart.StyleDefinitionsPart == null)
            return -1;

        var styles = mainPart.StyleDefinitionsPart.Styles;
        var bulletStyle = styles.Elements<Style>().FirstOrDefault(s => s.StyleId == styleId);
        if (bulletStyle != null && bulletStyle.StyleParagraphProperties != null)
        {
            var numPr = bulletStyle.StyleParagraphProperties.NumberingProperties;
            if (numPr != null && numPr.NumberingId != null && numPr.NumberingId.Val != null)
            {
                return (int)numPr.NumberingId.Val.Value;
            }
        }
        return -1;
    }


    public int GetOrCreateBulletNumberingInstance(MainDocumentPart mainPart)
    {
        // Ensure the numbering part exists.
        var numberingPart = mainPart.NumberingDefinitionsPart;
        if (numberingPart == null)
        {
            numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new Numbering();
            numberingPart.Numbering.Save();
        }

        // Look for an existing AbstractNum that defines bullet formatting.
        var bulletAbstract = numberingPart.Numbering
            .Descendants<AbstractNum>()
            .FirstOrDefault(a => a.Descendants<Level>()
                .Any(l => l.NumberingFormat?.Val == NumberFormatValues.Bullet));

        if (bulletAbstract != null)
        {
            // Look for a numbering instance (<w:num>) referencing that abstract.
            var bulletInstance = numberingPart.Numbering
                .Descendants<NumberingInstance>()
                .FirstOrDefault(n => n.AbstractNumId.Val == bulletAbstract.AbstractNumberId.Value);
            if (bulletInstance != null)
                return bulletInstance.NumberID.Value;
        }

        // If no bullet numbering exists, create a new bullet abstract.
        int newAbstractNumId = numberingPart.Numbering
            .Descendants<AbstractNum>()
            .Select(a => (int)a.AbstractNumberId.Value)
            .DefaultIfEmpty(0)
            .Max() + 1;
        var bulletAbstractNew = new AbstractNum(
            new Level(
                new NumberingFormat() { Val = NumberFormatValues.Bullet },
                new LevelText() { Val = "•" },
                new ParagraphProperties(new Indentation() { Left = "720" })
            )
            { LevelIndex = 0 }
        )
        {
            AbstractNumberId = newAbstractNumId
        };
        numberingPart.Numbering.Append(bulletAbstractNew);

        // Create a new numbering instance (<w:num>) for this abstract.
        int newNumId = numberingPart.Numbering
            .Descendants<NumberingInstance>()
            .Select(n => (int)n.NumberID.Value)
            .DefaultIfEmpty(0)
            .Max() + 1;
        var bulletInstanceNew = new NumberingInstance(new AbstractNumId() { Val = bulletAbstractNew.AbstractNumberId.Value })
        {
            NumberID = newNumId
        };
        numberingPart.Numbering.Append(bulletInstanceNew);
        numberingPart.Numbering.Save();

        return bulletInstanceNew.NumberID.Value;
    }




    /// <summary>
    /// Inserts a paragraph with explicit bullet numbering properties.
    /// </summary>
    public void InsertBulletParagraph(MainDocumentPart mainPart, string bulletText)
    {
        int bulletNumId = GetOrCreateBulletNumberingInstance(mainPart);

        // Create a new paragraph and add numbering properties that reference the bullet numbering instance.
        var p = new Paragraph(
            new ParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference() { Val = 0 },
                    new NumberingId() { Val = bulletNumId }
                )
            ),
            new Run(new Text(bulletText))
        );

        // Append the paragraph to the document body.
        mainPart.Document.Body.Append(p);
        mainPart.Document.Save();
    }
}
