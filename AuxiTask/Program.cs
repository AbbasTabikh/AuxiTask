using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using AuxiTask;

internal class Program
{
    private const string FilePath = @"C:\Users\abbas\source\repos\AuxiTask\AuxiTask\auxi_C#_Interview.pptx";

    private static void Main(string[] args)
    {
        using PresentationDocument presentationDocument = PresentationDocument.Open(FilePath, true);
        PresentationPart? presentationPart = presentationDocument.PresentationPart;

        if(presentationPart is null)
            throw new ArgumentNullException(nameof(presentationPart));
        
        //access the first slide
        SlidePart firstSlide = presentationPart!.SlideParts.ElementAt(1);


        /* Title part :*/
        var titleShape = GetTitleShape(firstSlide);

        if (titleShape != null)
            UpdateTitle(titleShape, "Output Slide");


        var chevrons = firstSlide.Slide.Descendants<P.Shape>().Where(x => string.IsNullOrEmpty(x.InnerText)).ToArray();
        var textboxes = firstSlide.Slide.Descendants<P.Shape>().Where(x => !string.IsNullOrEmpty(x.InnerText)).ToArray();

        /* Remove texboxes and replace the text in the chevrons */
        ReplaceTextBoxesText(chevrons, textboxes);


        /* Align and resize the chevrons */
        var firstchev = chevrons.First();
        Int64 lineLevel = firstchev.ShapeProperties.Transform2D.Offset.Y.Value;
        Int64 width =(Int64)(firstchev.ShapeProperties.Transform2D.Extents.Cx.Value * 1.35);
        Int64 height = firstchev.ShapeProperties.Transform2D.Extents.Cy.Value ;

        ResizeChevrons(chevrons, width, height);
        AlignChevrons(chevrons, lineLevel);


        /* bulleted textboxes */
        var listedTextBoxes = firstSlide.Slide.Descendants<P.Shape>().Where(IsList).ToArray();
        var listHeight = listedTextBoxes.First().ShapeProperties.Transform2D.Offset.Y;
        AlignLists(listedTextBoxes,chevrons, listHeight);
        RemoveUnderlineAndBoldText(listedTextBoxes[^1]);
        ChangeListsFont(listedTextBoxes, "Beirut");
        //SetDottedBullets(listedTextBoxes[1], listedTextBoxes[2]);
   
    }

    // Text boxes and chevron
    private static void AlignChevrons(P.Shape[] chevrons, Int64 lineLevel)
    {

        for (int i = 1; i < chevrons.Length; i++)
        {
            var prevChev = chevrons[i - 1];

            Console.WriteLine($"Current Chevron : {chevrons[i].InnerText}  - Previous Chevron : {prevChev.InnerText}\n");
            chevrons[i].ShapeProperties!.Transform2D!.Offset!.Y = lineLevel;
            Int64 newX = prevChev!.ShapeProperties!.Transform2D!.Offset!.X! + ((prevChev.ShapeProperties.Transform2D.Extents.Cx.Value - chevrons[i].ShapeProperties.Transform2D.Extents.Cx) + (long)(chevrons[i].ShapeProperties.Transform2D.Extents.Cy * 1.45));
            chevrons[i].ShapeProperties.Transform2D.Offset.X = newX;
        }
    }
    private static void ResizeChevrons(IEnumerable<P.Shape> chevrons, Int64 width, Int64 height)
    {
        foreach (var chevron in chevrons)
        {
            chevron.ShapeProperties.Transform2D.Extents.Cx = width;
            chevron.ShapeProperties.Transform2D.Extents.Cy = height;
        }
    }
    private static bool IsInside(P.Shape outer, P.Shape inner)
    {
        // Get the position and dimensions of Shape outer
        var outerShapeProperties = outer.ShapeProperties;
        // Get the position and dimensions of Shape inner
        var innerShapeProperties = inner.ShapeProperties;

        if(innerShapeProperties is null || outerShapeProperties is null)
            return false;

        var outerShapePosition = outerShapeProperties.Transform2D?.Offset;
        var outerShapeDimensions = outerShapeProperties.Transform2D?.Extents;

        var innerShapePosition = innerShapeProperties.Transform2D?.Offset;
        var innerShapeDimensions = innerShapeProperties.Transform2D?.Extents;

        return innerShapePosition?.X?.Value >= outerShapePosition?.X?.Value
            && innerShapePosition?.Y?.Value >= outerShapePosition?.Y?.Value
                && (innerShapePosition.X.Value + innerShapeDimensions?.Cx?.Value) <= (outerShapePosition?.X?.Value + outerShapeDimensions?.Cx?.Value)
                    && (innerShapePosition?.Y.Value + innerShapeDimensions?.Cy?.Value) <= (outerShapePosition?.Y.Value + outerShapeDimensions?.Cy?.Value);
    }
    private static void ReplaceTextBoxesText(IEnumerable<P.Shape> chevrons, IEnumerable<P.Shape> textboxes)
    {
        var first = chevrons.First();
        foreach (var chevron in chevrons)
        {
            foreach (var textbox in textboxes)
            {
                if (IsInside(chevron, textbox))
                {
                    Console.WriteLine($"{textbox.InnerText} is found inside chevron {chevron.NonVisualShapeProperties.NonVisualDrawingProperties.Id}");

                    D.Paragraph paragraph = new()
                    {
                        ParagraphProperties = new D.ParagraphProperties
                        {
                            Alignment = TextAlignmentTypeValues.Center,
                        }
                    };

                    D.Run run;

                    if(chevron == first)
                        run = new D.Run(new D.Text(textbox.InnerText + "\n"));
                    else
                        run = new D.Run(new D.Text(textbox.InnerText + "\n\n"));

                    run.RunProperties = new D.RunProperties();
                    run.RunProperties.FontSize = 1800;
                    run.RunProperties.Bold = false;
                    paragraph.Append(run);

                    chevron.TextBody.Append(paragraph);
                    textbox.Remove();

                }
            }
        }
    }
    

    //Title
    private static P.Shape? GetTitleShape(SlidePart slidePart)
    {
        // Get the collection of shapes in the corresponding slide
        var shapesInSlide = slidePart.Slide.Descendants<P.Shape>();

        foreach (var shape in shapesInSlide)
        {
           // Retrieve the first child element of ApplicationNonVisualDrawingProperties that is of type PlaceholderShape
           var placeholderShape = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>();

            if (placeholderShape is null)
                continue;

            if (placeholderShape.Type is not null && placeholderShape.Type == P.PlaceholderValues.Title)
                return shape;
        }

        return null;
    }
    private static void UpdateTitle(P.Shape shape, string text)
    {
        var paragraph = shape.TextBody?.Elements<D.Paragraph>().FirstOrDefault();

        if (paragraph is null)
            return;

        paragraph.RemoveAllChildren();

        // Apply the ParagraphProperties to the paragraph
        paragraph.ParagraphProperties = new D.ParagraphProperties { Alignment = TextAlignmentTypeValues.Center };

        //create a new run for the new text
        D.Run run = new(new D.Text (text));

        RunFonts runFonts = new()
        {
            Ascii = "Beirut"
        };

        var runProp = run.RunProperties;

        if (runProp is null)
            run.RunProperties = new D.RunProperties();

        run.RunProperties.Bold = true;
        run.RunProperties.FontSize = 5000;
        run.RunProperties.AppendChild(runFonts);

        //append to the paragraph
        paragraph.Append(run);
    }


    //Bulleted textboxes
    private static void ChangeListsFont(IEnumerable<P.Shape> bulletedLists, string  fontName)
    {
        foreach (var bulletedList in bulletedLists)
        {
            foreach (var paragraph in bulletedList.Descendants<D.Paragraph>())
            {
                foreach (var run in paragraph.Descendants<D.Run>())
                {
                    if (run.RunProperties != null)
                    {
                        run.SetFont(fontName);
                    }
                }
            }
        }
    }
    private static bool IsList(P.Shape textBox)
    {
        var paragraphs = textBox.TextBody.Elements<D.Paragraph>();

        if(paragraphs is null || !paragraphs.Any() ) 
            return false;

        var paragraph = paragraphs.First();
        if(paragraph.ParagraphProperties is not null && paragraph.ParagraphProperties.GetFirstChild<BulletFont>() != null)
            return true;

        return false;
    }
    private static void AlignLists(P.Shape[] bulletedLists, P.Shape[] chevrons, Int64 level)
    {
        for (int i = 1; i < bulletedLists.Length; i++)
        {
            bulletedLists[i].ShapeProperties.Transform2D.Offset.Y = level;
            bulletedLists[i].ShapeProperties.Transform2D.Offset.X = chevrons[i].ShapeProperties.Transform2D.Offset.X; 
        }
    }
    private static void RemoveUnderlineAndBoldText(P.Shape underlinedBulletList)
    {

        var paragraphs = underlinedBulletList.Descendants<D.Paragraph>();

        foreach (var paragraph in paragraphs)
        {
            D.RunProperties runProperties = paragraph.GetFirstChild<D.RunProperties>();
            foreach (var run in paragraph.Descendants<D.Run>())
            {
                if (run.RunProperties != null)
                {
                    // Remove the underline
                    run.RunProperties.Underline = null;
                    run.SetBold(false);
                }
            }
        }
    }
    private static void SetDottedBullets(P.Shape numberedList, P.Shape characterbasedList)
    {
        var paragraphsInNumberedList = numberedList.Descendants<D.Paragraph>();
        var paragraphsIncharacterbasedList = numberedList.Descendants<D.Paragraph>();

        foreach (var paragraph in paragraphsInNumberedList)
        {

        }

        foreach (var paragraph in paragraphsIncharacterbasedList)
        {
            var props = paragraph.ParagraphProperties.GetFirstChild<CharacterBullet>();
            //props.Char = "\u2022";
        }
    }


}