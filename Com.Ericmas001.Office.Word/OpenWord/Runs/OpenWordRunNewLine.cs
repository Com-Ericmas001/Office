using DocumentFormat.OpenXml.Wordprocessing;

namespace Com.Ericmas001.Office.Word.OpenWord.Runs
{

    public class OpenWordRunNewLine : OpenWordRunText
    {
        public OpenWordRunNewLine()
            : base(new Break())
        {

        }
    }
}
