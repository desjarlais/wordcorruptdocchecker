using System;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.IO.Packaging;
using System.Text.RegularExpressions;
using System.Xml;
using System.Collections.Generic;

namespace DocCorruptionChecker
{
    public partial class Form1 : Form
    {
        static List<string> nodes = new List<string>();
        static StringBuilder sbNodeBuffer = new StringBuilder();
        static StringBuilder sbChildNodeBuffer = new StringBuilder();

        const string txtFallbackStart = "<mc:Fallback>";
        const string txtFallbackEnd = "</mc:Fallback>";
        public static string fixedFallback = string.Empty;

        public Form1()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                tbxFileName.Text = openFileDialog1.FileName;
            }
        }

        private void btnScanDocument_Click(object sender, EventArgs e)
        {
            string strOrigFileName = tbxFileName.Text;
            string strDestPath = Path.GetDirectoryName(strOrigFileName) + "\\";
            string strExtension = Path.GetExtension(strOrigFileName);
            string strDestFileName = strDestPath + Path.GetFileNameWithoutExtension(strOrigFileName) + "(Fixed)" + strExtension;
            listBox1.Items.Clear();

            try
            {
                if ((File.GetAttributes(strOrigFileName) & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                {
                    listBox1.Items.Add("Document is Read-Only, unable to make changes.");
                    return;
                }

                File.Copy(strOrigFileName, strDestFileName);
            }
            // Catch exception if the file was already copied.
            catch (IOException copyError)
            {
                listBox1.Items.Add(copyError.Message);
                try
                {
                    File.Delete(strDestFileName);
                }
                catch (DirectoryNotFoundException dnf)
                {
                    listBox1.Items.Add("ERROR:" + dnf.Message);
                    return;
                }
                catch (IOException ioe)
                {
                    listBox1.Items.Add("ERROR: " + ioe.Message);
                    return;
                }
            }
            catch (UnauthorizedAccessException uae)
            {
                listBox1.Items.Add("Invalid file name. " + uae.Message);
            }

            try
            {
                using (Package package = Package.Open(strDestFileName, FileMode.Open, FileAccess.ReadWrite))
                {
                    foreach (PackagePart part in package.GetParts())
                    {
                        if (part.Uri.ToString() == "/word/document.xml")
                        {
                            XmlDocument xdoc = new XmlDocument();
                            try
                            {
                                xdoc.Load(part.GetStream(FileMode.Open, FileAccess.Read));
                                listBox1.Items.Add("This document does not contain invalid xml.");

                                // if the file doesn't contain invalid xml, delete the copied file outside the using block
                                File.Delete(strDestFileName);
                            }
                            catch (XmlException) // check for invalid xml
                            {
                                MemoryStream ms = new MemoryStream();
                                var valid = new ValidTags();
                                var invalid = new InvalidTags();

                                using (TextWriter tw = new StreamWriter(ms))
                                {
                                    string strDocText = string.Empty;
                                    using (TextReader tr = new StreamReader(part.GetStream(FileMode.Open, FileAccess.Read)))
                                    {
                                        try
                                        {
                                            strDocText = tr.ReadToEnd();

                                            foreach (var el in invalid.invalidXmlTags())
                                            {
                                                foreach (Match m in Regex.Matches(strDocText, el))
                                                {
                                                    switch (m.Value)
                                                    {
                                                        case ValidTags.strValidMcChoice1:
                                                            break;
                                                        case ValidTags.strValidMcChoice2:
                                                            break;
                                                        case ValidTags.strValidMcChoice3:
                                                            break;
                                                        case InvalidTags.strInvalidVshape:
                                                            strDocText = strDocText.Replace(m.Value, ValidTags.strValidVshape);
                                                            listBox1.Items.Add("Invalid Tag: " + m.Value);
                                                            listBox1.Items.Add("Replaced With: " + ValidTags.strValidVshape);
                                                            break;
                                                        case InvalidTags.strInvalidOmathWps:
                                                            strDocText = strDocText.Replace(m.Value, ValidTags.strValidomathwps);
                                                            listBox1.Items.Add("Invalid Tag: " + m.Value);
                                                            listBox1.Items.Add("Replaced With: " + ValidTags.strValidomathwps);
                                                            break;
                                                        case InvalidTags.strInvalidOmathWpg:
                                                            strDocText = strDocText.Replace(m.Value, ValidTags.strValidomathwpg);
                                                            listBox1.Items.Add("Invalid Tag: " + m.Value);
                                                            listBox1.Items.Add("Replaced With: " + ValidTags.strValidomathwpg);
                                                            break;
                                                        case InvalidTags.strInvalidOmathWpc:
                                                            strDocText = strDocText.Replace(m.Value, ValidTags.strValidomathwpc);
                                                            listBox1.Items.Add("Invalid Tag: " + m.Value);
                                                            listBox1.Items.Add("Replaced With: " + ValidTags.strValidomathwpc);
                                                            break;
                                                        case InvalidTags.strInvalidOmathWpi:
                                                            strDocText = strDocText.Replace(m.Value, ValidTags.strValidomathwpi);
                                                            listBox1.Items.Add("Invalid Tag: " + m.Value);
                                                            listBox1.Items.Add("Replaced With: " + ValidTags.strValidomathwpi);
                                                            break;
                                                        default:
                                                            // default catch for "strInvalidmcChoiceRegEx" and "strInvalidFallbackRegEx"
                                                            // since the exact string will never be the same and always has different trailing tags
                                                            // we need to conditionally check for specific patterns
                                                            // the first if </mc:Choice> is to catch and replace the invalid mc:Choice tags
                                                            if (m.Value.Contains("</mc:Choice>"))
                                                            {
                                                                if (m.Value.Contains("<mc:Fallback id="))
                                                                {
                                                                    // secondary check for a fallback that has an attribute.
                                                                    // we don't allow attributes in a fallback
                                                                    strDocText = strDocText.Replace(m.Value, ValidTags.strValidMcChoice4);
                                                                    listBox1.Items.Add("Invalid Tag: " + m.Value);
                                                                    listBox1.Items.Add("Replaced With: " + ValidTags.strValidMcChoice4);
                                                                    break;
                                                                }
                                                                else
                                                                {
                                                                    // replace mc:choice and hold onto the tag that follows
                                                                    strDocText = strDocText.Replace(m.Value, ValidTags.strValidMcChoice3 + m.Groups[2].Value);
                                                                    listBox1.Items.Add("Invalid Tag: " + m.Value);
                                                                    listBox1.Items.Add("Replaced With: " + ValidTags.strValidMcChoice3 + m.Groups[2].Value);
                                                                    break;
                                                                }
                                                            }
                                                            // the second if <w:pict/> is to catch and replace the invalid mc:Fallback tags
                                                            else if (m.Value.Contains("<w:pict/>"))
                                                            {
                                                                if (m.Value.Contains("</mc:Fallback>"))
                                                                {
                                                                    // if the match contains the closing fallback we just need to remove the entire fallback
                                                                    // this will leave the closing AC and Run tags, which should be correct
                                                                    strDocText = strDocText.Replace(m.Value, "");
                                                                    listBox1.Items.Add("Invalid Tag: " + m.Value);
                                                                    listBox1.Items.Add("Replaced With: " + "Fallback tag deleted.");
                                                                    break;
                                                                }
                                                                else
                                                                {
                                                                    // if there is no closing fallback tag, we can replace the match with the omitFallback valid tags
                                                                    // then we need to also add the trailing tag, since it's always different but needs to stay in the file
                                                                    strDocText = strDocText.Replace(m.Value, ValidTags.strOmitFallback + m.Groups[2].Value);
                                                                    listBox1.Items.Add("Invalid Tag: " + m.Value);
                                                                    listBox1.Items.Add("Replaced With: " + ValidTags.strOmitFallback + m.Groups[2].Value);
                                                                    break;
                                                                }
                                                            }
                                                            else
                                                            {
                                                                // leaving this open for future checks
                                                                break;
                                                            }
                                                    }
                                                }
                                            }

                                            // remove all fallback tags
                                            // start by getting a list of all nodes/values
                                            if (chkRemoveAllFallbackTags.Checked == true)
                                            {
                                                CharEnumerator charEnum = strDocText.GetEnumerator();
                                                while (charEnum.MoveNext())
                                                {
                                                    // opening tag
                                                    if (charEnum.Current == '<')
                                                    {
                                                        if (sbNodeBuffer.Length > 0)
                                                        {
                                                            nodes.Add(sbNodeBuffer.ToString());
                                                            sbNodeBuffer.Clear();
                                                        }
                                                        Node(charEnum.Current);
                                                    }
                                                    // close tag
                                                    else if (charEnum.Current == '>')
                                                    {
                                                        Node(charEnum.Current);
                                                        nodes.Add(sbNodeBuffer.ToString());
                                                        sbNodeBuffer.Clear();
                                                    }
                                                    // node text
                                                    else
                                                    {
                                                        Node(charEnum.Current);
                                                    }
                                                }

                                                listBox1.Items.Add("...removing all fallback tags");
                                                GetAllNodes(strDocText);
                                                strDocText = fixedFallback;
                                            }
                                        }
                                        catch (FileFormatException ffe)
                                        {
                                            listBox1.Items.Add("ERROR: " + ffe.Message);
                                            File.Delete(strDestFileName);
                                        }

                                        tw.Write(strDocText);
                                        tw.Flush();

                                        // rewrite the part
                                        ms.Position = 0;
                                        Stream partStream = part.GetStream(FileMode.Open, FileAccess.Write);
                                        partStream.SetLength(0);
                                        ms.WriteTo(partStream);

                                        listBox1.Items.Add("-------------------------------------------------------------");
                                        listBox1.Items.Add("Fixed Document Location: " + strDestFileName);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (IOException ioe)
            {
                listBox1.Items.Add("ERROR: " + ioe.Message);
                File.Delete(strDestFileName);
            }
            catch (FileFormatException ffe)
            {
                listBox1.Items.Add("ERROR: File may be password protected OR " + ffe.Message);
                File.Delete(strDestFileName);
            }
            catch (Exception ex)
            {
                listBox1.Items.Add("ERROR: " + ex.Message);
                File.Delete(strDestFileName);
            }
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            StringBuilder buffer = new StringBuilder();

            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                buffer.Append(listBox1.Items[i].ToString());
                buffer.Append('\n');
            }

            Clipboard.SetText(buffer.ToString());
        }

        public static void Node(char input)
        {
            sbNodeBuffer.Append(input);
        }

        public static void TagText(char input)
        {
            sbChildNodeBuffer.Append(input);
        }

        public static void GetAllNodes(string originalText)
        {
            // check each node and add fallback tags only to the list
            bool isFallback = false;
            List<string> fallback = new List<string>();

            foreach (object o in nodes)
            {
                if (o.ToString() == txtFallbackStart)
                {
                    isFallback = true;
                }

                if (isFallback)
                {
                    fallback.Add(o.ToString());
                }

                if (o.ToString() == txtFallbackEnd)
                {
                    isFallback = false;
                }
            }

            ParseOutFallbackTags(fallback, originalText);
        }

        public static void ParseOutFallbackTags(List<string> input, string originalText)
        {
            // we should only have a list of fallback start tags, end tags and each tag in between
            // the idea is to combine these start/middle/end tags into a long string
            // then they can be replaced with an empty string
            List<string> FallbackTagsAppended = new List<string>();
            StringBuilder sbFallback = new StringBuilder();

            foreach (object o in input)
            {
                if (o.ToString() == txtFallbackStart)
                {
                    sbFallback.Append(o);
                    continue;
                }
                else if (o.ToString() == txtFallbackEnd)
                {
                    sbFallback.Append(o);
                    FallbackTagsAppended.Add(sbFallback.ToString());
                    sbFallback.Clear();
                    continue;
                }
                else
                {
                    sbFallback.Append(o);
                    continue;
                }
            }

            sbFallback.Clear();

            foreach (object o in FallbackTagsAppended)
            {
                originalText = originalText.Replace(o.ToString(), "");
            }

            // each set of fallback tags should now be removed from the text
            // set it to the global variable so we can add it back into document.xml
            fixedFallback = originalText;
        }
    }
}