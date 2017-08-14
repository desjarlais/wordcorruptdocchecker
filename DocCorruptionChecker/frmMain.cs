using System;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.IO.Packaging;
using System.Text.RegularExpressions;
using System.Xml;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace DocCorruptionChecker
{
    public partial class frmMain : Form
    {
        static List<string> nodes = new List<string>();
        static StringBuilder sbNodeBuffer = new StringBuilder();

        const string txtFallbackStart = "<mc:Fallback>";
        const string txtFallbackEnd = "</mc:Fallback>";
        public static char prevChar = '<';
        public bool isRegularXmlTag = false;
        public bool isFixed = false;
        public static string fixedFallback = string.Empty;
        public static string strOrigFileName = string.Empty;
        public static string strDestPath = string.Empty;
        public static string strExtension = string.Empty;
        public static string strDestFileName = string.Empty;

        public frmMain()
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
        
        private void btnFixDocument_Click(object sender, EventArgs e)
        {
            try
            {
                strOrigFileName = tbxFileName.Text;
                strDestPath = Path.GetDirectoryName(strOrigFileName) + "\\";
                strExtension = Path.GetExtension(strOrigFileName);
                strDestFileName = strDestPath + Path.GetFileNameWithoutExtension(strOrigFileName) + "(Fixed)" + strExtension;

                // check if file we are about to copy exists and append a number so its unique
                if (File.Exists(strDestFileName))
                {
                    Random rNumber = new Random();
                    strDestFileName = strDestPath + Path.GetFileNameWithoutExtension(strOrigFileName) + "(Fixed)" + rNumber.Next(1, 100) + strExtension;
                }

                listBox1.Items.Clear();
                
                if (strExtension == ".docx")
                {
                    if ((File.GetAttributes(strOrigFileName) & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                    {
                        listBox1.Items.Add("ERROR: File is Read-Only.");
                        return;
                    }
                    else
                    {
                        File.Copy(strOrigFileName, strDestFileName);
                    }
                }

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
                            }
                            catch (XmlException) // invalid xml found, try to fix the contents
                            {
                                MemoryStream ms = new MemoryStream();
                                var valid = new ValidTags();
                                var invalid = new InvalidTags();

                                using (TextWriter tw = new StreamWriter(ms))
                                {
                                    string strDocText = string.Empty;
                                    using (TextReader tr = new StreamReader(part.GetStream(FileMode.Open, FileAccess.Read)))
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

                                        // remove all fallback tags is a 3 step process
                                        // Step 1. start by getting a list of all nodes/values in the document.xml file
                                        if (chkRemoveAllFallbackTags.Checked == true)
                                        {
                                            CharEnumerator charEnum = strDocText.GetEnumerator();
                                            while (charEnum.MoveNext())
                                            {
                                                // keep track of previous char
                                                prevChar = charEnum.Current;

                                                // opening tag
                                                if (charEnum.Current == '<')
                                                {
                                                    // if we haven't hit a close, but hit another '<' char
                                                    // we are not a true open tag so add it like a regular char
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
                                                    // there are 2 ways to close out a tag
                                                    // 1. self contained tag like <w:sz w:val="28"/>
                                                    // 2. standard xml <w:t>test</w:t>
                                                    // if previous char is '/', then we are an end tag
                                                    if (prevChar == '/' || isRegularXmlTag == true)
                                                    {
                                                        Node(charEnum.Current);
                                                        isRegularXmlTag = false;
                                                    }
                                                    Node(charEnum.Current);
                                                    nodes.Add(sbNodeBuffer.ToString());
                                                    sbNodeBuffer.Clear();
                                                }
                                                // node text
                                                else
                                                {
                                                    // this is the second xml closing style, keep track of char
                                                    if (prevChar == '<' && charEnum.Current == '/')
                                                    {
                                                        isRegularXmlTag = true;
                                                    }
                                                    Node(charEnum.Current);
                                                }
                                            }

                                            listBox1.Items.Add("...removing all fallback tags");
                                            GetAllNodes(strDocText);
                                            strDocText = fixedFallback;
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
                                        isFixed = true;
                                    }
                                }
                            }
                        }
                    }
                    if (isFixed == false)
                    {
                        listBox1.Items.Add("This document does not contain invalid xml.");
                    }
                }
            }
            catch (IOException)
            {
                listBox1.Items.Add("ERROR: Unable to fix document." );
            }
            catch (FileFormatException ffe)
            {
                // list out the possible reasons for this type of exception
                listBox1.Items.Add("ERROR: Unable to fix document.");
                listBox1.Items.Add("   Possible Causes:");
                listBox1.Items.Add("      - File may be password protected");
                listBox1.Items.Add("      - File was renamed to the .docx extension, but is not an actual .docx file");
                listBox1.Items.Add("      - " + ffe.Message);
            }
            catch (Exception ex)
            {
                listBox1.Items.Add("ERROR: Unable to fix document. " + ex.Message);
            }
            finally
            {
                // only delete destination file when there is an error
                // need to make sure the file stays when it is fixed
                if (isFixed == false)
                {
                    // delete the copied file if it exists
                    if (File.Exists(strDestFileName))
                    {
                        File.Delete(strDestFileName);
                    }
                }
                else
                {
                    // since we were able to attempt the fixes
                    // check if we can open in the sdk and confirm it was indeed fixed
                    listBox1.Items.Add("");
                    OpenWithSDK(strDestFileName);
                }

                // need to reset the globals 
                isFixed = false;
                isRegularXmlTag = false;
                fixedFallback = string.Empty;
                strOrigFileName = string.Empty;
                strDestPath = string.Empty;
                strExtension = string.Empty;
                strDestFileName = string.Empty;
                prevChar = '<';
            }
        }
        
        private void btnCopy_Click(object sender, EventArgs e)
        {
            try
            {
                if (listBox1.Items.Count > 0)
                {
                    StringBuilder buffer = new StringBuilder();

                    for (int i = 0; i < listBox1.Items.Count; i++)
                    {
                        buffer.Append(listBox1.Items[i].ToString());
                        buffer.Append('\n');
                    }

                    Clipboard.SetText(buffer.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception" + ex.Message);
            }
        }

        public static void Node(char input)
        {
            sbNodeBuffer.Append(input);
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

            // loop each item in the list and remove it from the document
            foreach (object o in FallbackTagsAppended)
            {
                originalText = originalText.Replace(o.ToString(), "");
            }

            // each set of fallback tags should now be removed from the text
            // set it to the global variable so we can add it back into document.xml
            fixedFallback = originalText;
        }

        /// <summary>
        /// function to open the fixed file in the SDK
        /// if the SDK fails to open the file, it still contains corruption
        /// warn the user to try remove all fallback tags
        /// </summary>
        /// <param name="file">the path to the initial fix attempt</param>
        public void OpenWithSDK(string file)
        {
            try
            {
                using (WordprocessingDocument document = WordprocessingDocument.Open(file, true))
                {
                    // need to try pulling the main document.xml part from the zip file
                    // this will confirm if the file is still corrupt
                    // if not, the exception catches and we can notify the user
                    MainDocumentPart main = document.MainDocumentPart;
                    string body = main.Document.Body.ToString();

                    // file opened so the file is successfully fixed.
                    listBox1.Items.Add("Secondary check succeeded, the file is fixed correctly.");
                }
            }
            catch (Exception)
            {
                // if the file failed to open, it still contains errors
                listBox1.Items.Add("Secondary check failed, not all corrupt tags were fixed.");
                listBox1.Items.Add("Try using the Remove Fallback option.");
            }
        }
    }
}