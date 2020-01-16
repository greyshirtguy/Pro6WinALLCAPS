using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;
using System.Runtime.Serialization;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using System.Windows.Markup;

namespace Pro6WinALLCAPS
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // Path to Pro6 library (defaults to [Documents]\ProPresenter6 at startup)
        string strCurrentLibraryPath;

        public MainWindow()
        {
            InitializeComponent();
            strCurrentLibraryPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\ProPresenter6";
        }

        private bool ConvertAllTextWithinPro6DocumentToUpperCase(string pathToPro6Document)
        {
            // Load pro6 file into XMLDocument object
            XmlDocument rvPresentationDocumentXMLDoc = new XmlDocument();
            try
            {
                rvPresentationDocumentXMLDoc.Load(pathToPro6Document);
            }
            catch (Exception)
            {
                System.Windows.MessageBox.Show("Unable to open/load: " + System.IO.Path.GetFileName(pathToPro6Document) + Environment.NewLine + "Do you have open in Pro6 at the moment? - Close the document and try again. Otherwise try closing Pro6.", "Could not load Pro6 Document.", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }


            // Perform a quick check to see if this a .Pro6 document from a Mac that has never been opened in Pro6 on Windows (it will be missing WinFlowData and WinFontData).
            if (!rvPresentationDocumentXMLDoc.OuterXml.Contains("WinFlowData"))
            {
                System.Windows.MessageBox.Show(System.IO.Path.GetFileName(pathToPro6Document) + " has never been opened Pro6 on Windows" + Environment.NewLine + "Please open it with Pro6 on Windows first!. This will add the required Windows-specific data. Then come back and try again", "Missing Required Windows Specific Data", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }


            // All text on slides in Pro6 is contained within objects called "RVTextElement" (When you "Add Text To Slide" you get one of these).
            // Each RVTextElement (usually) has 4 child "NSString" XML nodes (identified by an rvXMLIvarName attribute) 
            // 1. rvXMLIvarName="PlainText"
            // 2. rvXMLIvarName="RTFData"
            // 3. rvXMLIvarName="WinFlowData"
            // 4. rvXMLIvarName="WinFontData"

            // Note: The text value of each the four NSString XML nodes is BASE64 encoded...So we have to deal with encoding and decoding when working with these.

            // "PlainText" is used for (I dunno - maybe stage display). Quite obviously it contains a plain text (unformmatted) copy of the formated text in the RVTextElement
            // "RTFData" contains the formated text as (a slightly customised) RichTextFormat and seems to be specifically used by Pro6 on MacOS
            // The next two types of NSString nodes seem to used specifically by Pro6 on Windows...
            // "WinFlowData" contains the formatted text as a .NET FlowDocument object serialized to XML. Given the namespace, it seems it is serialized by XAMLWriter rather than DataContractSerializer - I'm not sure but using XAMLReader to deserialize it seems to work fine.
            // "WinFontData" contains an RVFont object serialized to XML (seems to always have just a little bit of extra font info like Kerning, Spacing, and Outline Color & Width)

            // Okay, now that we understand a bit about the XML.....Let's write a function to convert all text to uppercase.
            // In this case, I will load the entire XML file and and find all the RVTextElements, and for each one found, convert it's text (to uppercase).
            // We only need to do this for three types: PlainText, RTFDAta and WinFlowData
            // 
            // Sounds easy right? For the PlainText and WinFlowData (FlowDocument) it is easy enough!
            // However I'm stuck on RTFData!!! I have personally found parsing RTF to find all the text and convert to uppercase, while keeping all other formatting, difficult!
            // My Google-Fu failed me - it seems there's not many working with RTF these days!
            // We could use built .NET functions to convert the FlowDocument to RTF (but the RichTextFormat in RTFData is slightly customised to be Pro6 specific).
            // 
            // Enter the magic of .NET!
            // ProPresenter 6 for Windows is written using .NET framework and it is quite wonderfully organised into multiple .NET assemblies (as ProPresenter.exe and quite a few .NET assemblies in .DLLs)
            // We can reference these ProPresenter assemblies DLLs in our own .NET projects (like this one) and use the Classes, Methods/Functions, Enums, Interfaces, etc present in them.
            // Of course this is not supported and there is no public documentation so it's all guessing games on what the assemblies are and what they do.
            // However, the devs have used good naming conventions that are nice and self-describing within the assemblies (same for the file format).
            // This makes the guessing game much easier and it's obvious there is a LOT (probably MOST) of the ProPresenter funcationality within these assemblies.
            // For this application I tried doing the obvious and creating an RVPresentation object and ismply usings it's .load function to load it from file.
            // This would have been nice and easy (loaded in two lines) and from there I could then enumerate the slides and the text objects, to convert the formatted text to upper case and then simply save it back to file
            // But alas, I can't get anything to load from the ProPresenter.BO.RenderingEngine.Entities assemblies as it fails to load a dependancy of the assembly Microsoft.DirectX.Direct3DX that it itself depends on
            // I could not figure out the cause/solution in a timely fashion, and eventually moved on to this different solution.
            // For this application we load the XML of the file as a single XMLDocument and then use standard XMl functions to find the WinFlowData and directly modify it's XML to convert all text values to uppercase.
            // But I do use two ProPresenter assemblies using "ProPresenter.BO.Common" and "ProPresenter.Common" - these contain RVSerializerHelper and RVFont classes respectivaly.
            // I found that RVSerializerHelper contains an override for. Serliaze that takes a FlowDocument and an RVFont and returns the encoded RTFData string (all with custom Pro6 RTF formatting)
            // This means I can manually update the FlowDocument XMl to convert it's text to uppercase using standard XML libraries...
            // Manually update the PlainText to capatilzed text...
            // Deserialize and create the RVFont object from WinFontData...
            // Create a new FlowDocument (using the new XML that has been capatilized)...
            // Call RVSerializerHelper.Serialize(newcapatilizedFlowDocument, rvFont) to get the capilized version of the Mac Pro6 specific RTF data (already encoded and ready to update the XML of RTFData node)
            // Update the XMl of PlainText, RTFData and WinFlowData with newly capalitzed versions and save the file.

            using (XmlNodeList xmlNodeListAllRVTextElements = rvPresentationDocumentXMLDoc.SelectNodes(".//RVTextElement"))
            {
                if (xmlNodeListAllRVTextElements != null)
                {
                    foreach (XmlNode xmlNodeRVText in xmlNodeListAllRVTextElements)
                    {
                        // Update (or add) the useAllCaps attribute to = "true" (This is the All-CAPS option supported only for Pro6 on MacOS)
                        // NB: This leaves the MacOS-specific text in RTFData unchanged (it is just displayed as uppercase)
                        XmlAttribute useAllCapsAttr = rvPresentationDocumentXMLDoc.CreateAttribute("useAllCaps"); //useAllCaps
                        useAllCapsAttr.Value = "true";
                        xmlNodeRVText.Attributes.Append(useAllCapsAttr);  // .Append will update if attribute already exists

                        using (XmlNodeList xmlNodeListNSStrings = xmlNodeRVText.SelectNodes(".//NSString"))
                        {
                            if (xmlNodeListNSStrings != null)
                            {
                                // This object will hold the formated text (FlowDocument) used by Pro6 on Windows.
                                FlowDocument winFlowData = null;

                                foreach (XmlNode xmlNodeNSString in xmlNodeListNSStrings)
                                {
                                    if (xmlNodeNSString.Attributes == null)
                                    {
                                        continue;
                                    }
                                    XmlNode rvXMLIvarName = xmlNodeNSString.Attributes.GetNamedItem("rvXMLIvarName");
                                    if (rvXMLIvarName.Value == "PlainText")
                                    {
                                        string plainText = Base64Decode(xmlNodeNSString.InnerText);
                                        if (plainText != "Double-click to edit")
                                        {
                                            string uppercasePlainText = plainText.ToUpper();
                                            xmlNodeNSString.InnerText = Base64Encode(uppercasePlainText);
                                        }
                                    }
                                    else if (rvXMLIvarName.Value == "RTFData")
                                    {
                                        // Nothing to do here - since we leave the text itself unchanged and simpy set useAllCaps="true" for the RVTextText Element
                                    }
                                    else if (rvXMLIvarName.Value == "WinFlowData")
                                    {
                                        // Get the FlowDocument XML as a string (it's stored in the file base64 encoded)
                                        string flowDocXMLString = Base64Decode(xmlNodeNSString.InnerText);

                                        // A list of possible XML special escapes that might be within the text but should not be changed when converting to uppercase:
                                        String[] xmlSpecials = { "&lt;", "&gt;", "&amp;", "&apos;", "&quot;" };

                                        // Create a temporary FlowDocument
                                        XmlDocument flowDocXMLDocument = new XmlDocument();

                                        // Load the temporary FlowDocument using the FlowDocument XML string and then search for all Text nodes and convert text to uppercase...
                                        flowDocXMLDocument.LoadXml(flowDocXMLString);
                                        XmlNodeList allTheFlowDocXMLNodes = flowDocXMLDocument.SelectNodes("//*");
                                        foreach (XmlNode flowDocXMLNode in allTheFlowDocXMLNodes)
                                        {
                                            if (flowDocXMLNode.HasChildNodes)
                                            {
                                                foreach (XmlNode childNode in flowDocXMLNode.ChildNodes)
                                                {
                                                    if (childNode.NodeType == XmlNodeType.Text)
                                                    {
                                                        // childNode found that contains plain text - let's convert the case of it's text (and revert XML specials back to lower case)
                                                        if (childNode.InnerText != "Double-click to edit")
                                                        {
                                                            childNode.InnerText = childNode.InnerText.ToUpper(); // Make text uppercase
                                                            foreach (string specialString in xmlSpecials)
                                                            {
                                                                childNode.InnerText = childNode.InnerText.Replace(specialString.ToUpper(), specialString); // Revert any XML special text back to lowercase
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        // Get the case-converted FlowDocument as a string
                                        string flowDocXMLStringCaseConverted = flowDocXMLDocument.OuterXml;
                                        // Base64 encode the case-converted FlowDocument XML string and update the text of the WinFlowData NSString element
                                        xmlNodeNSString.InnerText = Base64Encode(flowDocXMLStringCaseConverted);

                                    }
                                    else if (rvXMLIvarName.Value == "WinFontData")
                                    {
                                        // Nothing to do here - no need to change WinFontData when converting to uppercase
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Create a new file with -ALLCAPS suffix
            rvPresentationDocumentXMLDoc.Save(pathToPro6Document.Replace(".pro6", "-ALLCAPS.pro6"));

            return true;
        }

        private string Base64Encode(string plainText)
        {
            try
            {
                var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
                return System.Convert.ToBase64String(plainTextBytes);
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        private string Base64Decode(string base64EncodedData)
        {
            try 
            { 
                var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
                return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
            }
            catch (Exception)
            {
                return string.Empty;
            }
}

        private void UpdateLibraryList(string strLibraryPath)
        {
            lstLibrary.Items.Clear();
            if (Directory.Exists(strLibraryPath))
            {
                string[] files = Directory.GetFiles(strLibraryPath, "*.pro6");
                foreach (string file in files)
                {
                    string fileName = System.IO.Path.GetFileNameWithoutExtension(file);
                    lstLibrary.Items.Add(fileName);
                }
            }
        }

        private void BtnAllCapsCopy_Click(object sender, RoutedEventArgs e)
        {
            if (lstLibrary.SelectedItems.Count > 0)
            {
                object selectItem = lstLibrary.SelectedItem;
                
                // Do nothing if user tries to convert a document that is already -ALLCAPS
                if (selectItem.ToString().EndsWith("-ALLCAPS"))
                    return;

                if (ConvertAllTextWithinPro6DocumentToUpperCase(strCurrentLibraryPath + "\\" + selectItem.ToString() + ".pro6"))
                {
                    if (!lstLibrary.Items.Contains(selectItem.ToString() + "-ALLCAPS"))
                        lstLibrary.Items.Insert(lstLibrary.Items.IndexOf(selectItem), selectItem.ToString() + "-ALLCAPS");
                }
            }
        }

        private void LstLibrary_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            BtnAllCapsCopy_Click(null, null);
        }

        private void btnSelectLibraryPath_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();

            dlg.Description = "Select ProPresenter 6 Library Folder...";

            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                strCurrentLibraryPath = dlg.SelectedPath;
                UpdateLibraryList(strCurrentLibraryPath);

            }
        }

        private void Image_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            System.Windows.MessageBox.Show("Hello");
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            UpdateLibraryList(strCurrentLibraryPath);
        }
    }
}
