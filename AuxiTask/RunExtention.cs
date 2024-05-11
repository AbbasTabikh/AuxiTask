using DocumentFormat.OpenXml.Wordprocessing;
using D = DocumentFormat.OpenXml.Drawing;

namespace AuxiTask
{
    internal static class RunExtention
    {
        public static void SetFont(this D.Run run, string fontName)
        {
            RunFonts runFont = new()
            {
                Ascii = fontName
            };

            var runProp = run.RunProperties;
            if(runProp is null)
                runProp = new D.RunProperties();

            runProp.AppendChild(runFont);
        }

        public static void SetBold(this D.Run run, bool value)
        {
            var runProp = run.RunProperties;

            if (runProp is null)
                runProp = new D.RunProperties();

            runProp.Bold = value;
        }

        public static void SetFontSize(this D.Run run, Int32 size)
        {
            var runProp = run.RunProperties;

            if (runProp is null)
                runProp = new D.RunProperties();

            runProp.FontSize = size;
        }
    }
}
