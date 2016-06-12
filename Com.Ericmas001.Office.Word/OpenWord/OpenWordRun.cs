using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Com.Ericmas001.Office.Word.OpenWord
{
    public abstract class OpenWordRun
    {
        public abstract Run ObtainRun(WordprocessingDocument package);
    }
}
