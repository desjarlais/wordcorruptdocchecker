using System.Collections.Generic;

namespace DocCorruptionChecker
{
    class ValidTags
    {
        // Valid tag replacements
        public const string strValidMcChoice1 = "</mc:Choice><mc:Choice>";
        public const string strValidMcChoice2 = "</mc:Choice><mc:Fallback>";
        public const string strValidMcChoice3 = "</mc:Choice></mc:AlternateContent>";
        public const string strValidMcChoice4 = "</mc:Choice></mc:AlternateContent></w:r>";
        public const string strOmitFallback = "</mc:AlternateContent></w:r>";

        // Example: valid xml tag string values
        // <m:oMath><mc:AlternateContent><mc:Choice Requires="wps">
        // Example: Escape character value for RegEx matches
        // <m:oMath><mc:AlternateContent><mc:Choice Requires=\"wps\">
        public const string strValidomathwps = "<m:oMath><mc:AlternateContent><mc:Choice Requires=\"wps\">";
        public const string strValidomathwpg = "<m:oMath><mc:AlternateContent><mc:Choice Requires=\"wpg\">";
        public const string strValidomathwpi = "<m:oMath><mc:AlternateContent><mc:Choice Requires=\"wpi\">";
        public const string strValidomathwpc = "<m:oMath><mc:AlternateContent><mc:Choice Requires=\"wpc\">";
        public const string strValidVshape = "</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent>";

        public IEnumerable<string> validXmlTags()
        {
            yield return strValidMcChoice1;
            yield return strValidMcChoice2;
            yield return strValidMcChoice3;
            yield return strValidMcChoice4;
            yield return strValidomathwpc;
            yield return strValidomathwpg;
            yield return strValidomathwpi;
            yield return strValidomathwps;
            yield return strOmitFallback;
            yield return strValidVshape;
        }
    }
}
