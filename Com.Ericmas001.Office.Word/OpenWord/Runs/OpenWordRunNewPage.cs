using DocumentFormat.OpenXml.Wordprocessing;

namespace Com.Ericmas001.Office.Word.OpenWord.Runs
{
    public class OpenWordRunNewPage : OpenWordRunText
    {
        public OpenWordRunNewPage()
            : base(new Break { Type = BreakValues.Page })
        {

        }
    }
}
