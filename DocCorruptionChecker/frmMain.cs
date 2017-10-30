using System;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.IO.Packaging;
using System.Text.RegularExpressions;
using System.Xml;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using System.Diagnostics;

namespace DocCorruptionChecker
{
    public partial class FrmMain : Form
    {
        private static List<string> _nodes = new List<string>();
        private static StringBuilder _sbNodeBuffer = new StringBuilder();

        private const string TxtFallbackStart = "<mc:Fallback>";
        private const string TxtFallbackEnd = "</mc:Fallback>";
        public static char PrevChar = '<';
        public bool IsRegularXmlTag;
        public bool IsFixed;
        public static string FixedFallback = string.Empty;
        public static string StrOrigFileName = string.Empty;
        public static string StrDestPath = string.Empty;
        public static string StrExtension = string.Empty;
        public static string StrDestFileName = string.Empty;

        public FrmMain()
        {
            InitializeComponent();
        }

        /// <summary>
        /// display the browse file dialog
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                tbxFileName.Text = openFileDialog1.FileName;
            }
        }
        
        /// <summary>
        /// get a copy of the corrupt file instead of working with the original
        /// pull out the document.xml file and start searching for known corrupt tag sequences
        /// if nothing is found, add the brute force style approach of deleting all mc:fallback tags
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnFixDocument_Click(object sender, EventArgs e)
        {
            try
            {
                StrOrigFileName = tbxFileName.Text;
                StrDestPath = Path.GetDirectoryName(StrOrigFileName) + "\\";
                StrExtension = Path.GetExtension(StrOrigFileName);
                StrDestFileName = StrDestPath + Path.GetFileNameWithoutExtension(StrOrigFileName) + "(Fixed)" + StrExtension;

                // check if file we are about to copy exists and append a number so its unique
                if (File.Exists(StrDestFileName))
                {
                    Random rNumber = new Random();
                    StrDestFileName = StrDestPath + Path.GetFileNameWithoutExtension(StrOrigFileName) + "(Fixed)" + rNumber.Next(1, 100) + StrExtension;
                }

                lstOutput.Items.Clear();
                
                if (StrExtension == ".docx")
                {
                    if ((File.GetAttributes(StrOrigFileName) & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                    {
                        lstOutput.Items.Add("ERROR: File is Read-Only.");
                        return;
                    }
                    else
                    {
                        File.Copy(StrOrigFileName, StrDestFileName);
                    }
                }

                using (Package package = Package.Open(StrDestFileName, FileMode.Open, FileAccess.ReadWrite))
                {
                    foreach (PackagePart part in package.GetParts())
                    {
                        if (part.Uri.ToString() == "/word/document.xml")
                        {
                            XmlDocument xdoc = new XmlDocument();
                            try
                            {
                                xdoc.Load(part.GetStream(FileMode.Open, FileAccess.Read));
                            }
                            catch (XmlException) // invalid xml found, try to fix the contents
                            {
                                MemoryStream ms = new MemoryStream();
                                InvalidTags invalid = new InvalidTags();

                                using (TextWriter tw = new StreamWriter(ms))
                                {
                                    using (TextReader tr = new StreamReader(part.GetStream(FileMode.Open, FileAccess.Read)))
                                    {
                                        string strDocText = tr.ReadToEnd();

                                        foreach (string el in invalid.InvalidXmlTags())
                                        {
                                            foreach (Match m in Regex.Matches(strDocText, el))
                                            {
                                                switch (m.Value)
                                                {
                                                    case ValidTags.StrValidMcChoice1:
                                                        break;
                                                    case ValidTags.StrValidMcChoice2:
                                                        break;
                                                    case ValidTags.StrValidMcChoice3:
                                                        break;
                                                    case InvalidTags.StrInvalidVshape:
                                                        strDocText = strDocText.Replace(m.Value, ValidTags.StrValidVshape);
                                                        lstOutput.Items.Add("Invalid Tag: " + m.Value);
                                                        lstOutput.Items.Add("Replaced With: " + ValidTags.StrValidVshape);
                                                        break;
                                                    case InvalidTags.StrInvalidOmathWps:
                                                        strDocText = strDocText.Replace(m.Value, ValidTags.StrValidomathwps);
                                                        lstOutput.Items.Add("Invalid Tag: " + m.Value);
                                                        lstOutput.Items.Add("Replaced With: " + ValidTags.StrValidomathwps);
                                                        break;
                                                    case InvalidTags.StrInvalidOmathWpg:
                                                        strDocText = strDocText.Replace(m.Value, ValidTags.StrValidomathwpg);
                                                        lstOutput.Items.Add("Invalid Tag: " + m.Value);
                                                        lstOutput.Items.Add("Replaced With: " + ValidTags.StrValidomathwpg);
                                                        break;
                                                    case InvalidTags.StrInvalidOmathWpc:
                                                        strDocText = strDocText.Replace(m.Value, ValidTags.StrValidomathwpc);
                                                        lstOutput.Items.Add("Invalid Tag: " + m.Value);
                                                        lstOutput.Items.Add("Replaced With: " + ValidTags.StrValidomathwpc);
                                                        break;
                                                    case InvalidTags.StrInvalidOmathWpi:
                                                        strDocText = strDocText.Replace(m.Value, ValidTags.StrValidomathwpi);
                                                        lstOutput.Items.Add("Invalid Tag: " + m.Value);
                                                        lstOutput.Items.Add("Replaced With: " + ValidTags.StrValidomathwpi);
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
                                                                strDocText = strDocText.Replace(m.Value, ValidTags.StrValidMcChoice4);
                                                                lstOutput.Items.Add("Invalid Tag: " + m.Value);
                                                                lstOutput.Items.Add("Replaced With: " + ValidTags.StrValidMcChoice4);
                                                                break;
                                                            }
                                                            
                                                            // replace mc:choice and hold onto the tag that follows
                                                            strDocText = strDocText.Replace(m.Value, ValidTags.StrValidMcChoice3 + m.Groups[2].Value);
                                                            lstOutput.Items.Add("Invalid Tag: " + m.Value);
                                                            lstOutput.Items.Add("Replaced With: " + ValidTags.StrValidMcChoice3 + m.Groups[2].Value);
                                                            break;
                                                        }
                                                        // the second if <w:pict/> is to catch and replace the invalid mc:Fallback tags
                                                        else if (m.Value.Contains("<w:pict/>"))
                                                        {
                                                            if (m.Value.Contains("</mc:Fallback>"))
                                                            {
                                                                // if the match contains the closing fallback we just need to remove the entire fallback
                                                                // this will leave the closing AC and Run tags, which should be correct
                                                                strDocText = strDocText.Replace(m.Value, "");
                                                                lstOutput.Items.Add("Invalid Tag: " + m.Value);
                                                                lstOutput.Items.Add("Replaced With: " + "Fallback tag deleted.");
                                                                break;
                                                            }
                                                            
                                                            // if there is no closing fallback tag, we can replace the match with the omitFallback valid tags
                                                            // then we need to also add the trailing tag, since it's always different but needs to stay in the file
                                                            strDocText = strDocText.Replace(m.Value, ValidTags.StrOmitFallback + m.Groups[2].Value);
                                                            lstOutput.Items.Add("Invalid Tag: " + m.Value);
                                                            lstOutput.Items.Add("Replaced With: " + ValidTags.StrOmitFallback + m.Groups[2].Value);
                                                            break;
                                                        }
                                                        else
                                                        {
                                                            // leaving this open for future checks
                                                            break;
                                                        }
                                                }
                                            }
                                        }

                                        // remove all fallback tags is a 3 step process
                                        // Step 1. start by getting a list of all nodes/values in the document.xml file
                                        if (Properties.Settings.Default.RemoveFallback == "true")
                                        {
                                            CharEnumerator charEnum = strDocText.GetEnumerator();
                                            while (charEnum.MoveNext())
                                            {
                                                // keep track of previous char
                                                PrevChar = charEnum.Current;

                                                // opening tag
                                                switch (charEnum.Current)
                                                {
                                                    case '<':
                                                        // if we haven't hit a close, but hit another '<' char
                                                        // we are not a true open tag so add it like a regular char
                                                        if (_sbNodeBuffer.Length > 0)
                                                        {
                                                            _nodes.Add(_sbNodeBuffer.ToString());
                                                            _sbNodeBuffer.Clear();
                                                        }
                                                        Node(charEnum.Current);
                                                        break;
                                                    case '>':
                                                        // there are 2 ways to close out a tag
                                                        // 1. self contained tag like <w:sz w:val="28"/>
                                                        // 2. standard xml <w:t>test</w:t>
                                                        // if previous char is '/', then we are an end tag
                                                        if (PrevChar == '/' || IsRegularXmlTag)
                                                        {
                                                            Node(charEnum.Current);
                                                            IsRegularXmlTag = false;
                                                        }
                                                        Node(charEnum.Current);
                                                        _nodes.Add(_sbNodeBuffer.ToString());
                                                        _sbNodeBuffer.Clear();
                                                        break;
                                                    default:
                                                        // this is the second xml closing style, keep track of char
                                                        if (PrevChar == '<' && charEnum.Current == '/')
                                                        {
                                                            IsRegularXmlTag = true;
                                                        }
                                                        Node(charEnum.Current);
                                                        break;
                                                }
                                            }

                                            lstOutput.Items.Add("...removing all fallback tags");
                                            GetAllNodes(strDocText);
                                            strDocText = FixedFallback;
                                        }

                                        tw.Write(strDocText);
                                        tw.Flush();

                                        // rewrite the part
                                        ms.Position = 0;
                                        Stream partStream = part.GetStream(FileMode.Open, FileAccess.Write);
                                        partStream.SetLength(0);
                                        ms.WriteTo(partStream);

                                        lstOutput.Items.Add("-------------------------------------------------------------");
                                        lstOutput.Items.Add("Fixed Document Location: " + StrDestFileName);
                                        IsFixed = true;

                                        // open the file in Word
                                        if (Properties.Settings.Default.OpenInWord == "true")
                                        {
                                            Process.Start("winword", StrDestFileName);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if (IsFixed == false)
                    {
                        lstOutput.Items.Add("This document does not contain invalid xml.");
                    }
                }
            }
            catch (IOException)
            {
                lstOutput.Items.Add("ERROR: Unable to fix document." );
            }
            catch (FileFormatException ffe)
            {
                // list out the possible reasons for this type of exception
                lstOutput.Items.Add("ERROR: Unable to fix document.");
                lstOutput.Items.Add("   Possible Causes:");
                lstOutput.Items.Add("      - File may be password protected");
                lstOutput.Items.Add("      - File was renamed to the .docx extension, but is not an actual .docx file");
                lstOutput.Items.Add("      - " + ffe.Message);
            }
            catch (Exception ex)
            {
                lstOutput.Items.Add("ERROR: Unable to fix document. " + ex.Message);
            }
            finally
            {
                // only delete destination file when there is an error
                // need to make sure the file stays when it is fixed
                if (IsFixed == false)
                {
                    // delete the copied file if it exists
                    if (File.Exists(StrDestFileName))
                    {
                        File.Delete(StrDestFileName);
                    }
                }
                else
                {
                    // since we were able to attempt the fixes
                    // check if we can open in the sdk and confirm it was indeed fixed
                    lstOutput.Items.Add("");
                    OpenWithSdk(StrDestFileName);
                }

                // need to reset the globals 
                IsFixed = false;
                IsRegularXmlTag = false;
                FixedFallback = string.Empty;
                StrOrigFileName = string.Empty;
                StrDestPath = string.Empty;
                StrExtension = string.Empty;
                StrDestFileName = string.Empty;
                PrevChar = '<';
            }
        }
        
        /// <summary>
        /// copy everything in the listbox and add to clipboard
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnCopy_Click(object sender, EventArgs e)
        {
            try
            {
                if (lstOutput.Items.Count <= 0) return;

                StringBuilder buffer = new StringBuilder();

                foreach (object t in lstOutput.Items)
                {
                    buffer.Append(t);
                    buffer.Append('\n');
                }

                Clipboard.SetText(buffer.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception" + ex.Message);
            }
        }

        // add current character to the node buffer
        public static void Node(char input)
        {
            _sbNodeBuffer.Append(input);
        }
        
        /// <summary>
        /// Step 2 of remove fallback tags
        /// this function loops through all nodes parsed out from Step 1
        /// check each node and add fallback tags only to the list
        /// </summary>
        /// <param name="originalText"></param>
        public static void GetAllNodes(string originalText)
        {
            bool isFallback = false;
            var fallback = new List<string>();

            foreach (string o in _nodes)
            {
                if (o == TxtFallbackStart)
                {
                    isFallback = true;
                }

                if (isFallback)
                {
                    fallback.Add(o);
                }

                if (o == TxtFallbackEnd)
                {
                    isFallback = false;
                }
            }

            ParseOutFallbackTags(fallback, originalText);
        }

        /// <summary>
        /// Step 3 of remove fallback tags
        /// we should only have a list of fallback start tags, end tags and each tag in between
        /// the idea is to combine these start/middle/end tags into a long string
        /// then they can be replaced with an empty string
        /// </summary>
        /// <param name="input"></param>
        /// <param name="originalText"></param>
        public static void ParseOutFallbackTags(List<string> input, string originalText)
        {
            var fallbackTagsAppended = new List<string>();
            StringBuilder sbFallback = new StringBuilder();

            foreach (object o in input)
            {
                switch (o.ToString())
                {
                    case TxtFallbackStart:
                        sbFallback.Append(o);
                        continue;
                    case TxtFallbackEnd:
                        sbFallback.Append(o);
                        fallbackTagsAppended.Add(sbFallback.ToString());
                        sbFallback.Clear();
                        continue;
                    default:
                        sbFallback.Append(o);
                        continue;
                }
            }

            sbFallback.Clear();

            // loop each item in the list and remove it from the document
            originalText = fallbackTagsAppended.Aggregate(originalText, (current, o) => current.Replace(o.ToString(), ""));

            // each set of fallback tags should now be removed from the text
            // set it to the global variable so we can add it back into document.xml
            FixedFallback = originalText;
        }

        /// <summary>
        /// function to open the fixed file in the SDK
        /// if the SDK fails to open the file, it still contains corruption
        /// warn the user to try remove all fallback tags
        /// </summary>
        /// <param name="file">the path to the initial fix attempt</param>
        public void OpenWithSdk(string file)
        {
            try
            {
                using (WordprocessingDocument document = WordprocessingDocument.Open(file, true))
                {
                    // need to try pulling the document.xml part from the zip file
                    // this will confirm if the file is still corrupt
                    // if it is, the exception catches and we can notify the user
                    MainDocumentPart main = document.MainDocumentPart;
                    string body = main.Document.Body.ToString();

                    // file opened so the file is successfully fixed.
                    lstOutput.Items.Add("Secondary check succeeded, the file is fixed correctly.");
                }
            }
            catch (Exception)
            {
                // if the file failed to open, it still contains errors
                lstOutput.Items.Add("Secondary check failed, not all corrupt tags were fixed.");
                lstOutput.Items.Add("Try using the Remove Fallback option.");
            }
        }

        private void BtnSettings_Click(object sender, EventArgs e)
        {
            FrmSettings form = new FrmSettings();
            form.Show();
        }
    }
}