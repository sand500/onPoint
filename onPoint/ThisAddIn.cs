using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Diagnostics;
using System.Windows.Forms;
using System.Net.Http;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;
using Flurl.Http;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using Newtonsoft.Json;

namespace onPoint
{
    public partial class ThisAddIn
    {
        private UserControl1 myUserControl1;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        public Dictionary<int, SlideContents> slideDataList;
        private String key = "default";

        private static void xmlWrite(Microsoft.Office.Interop.PowerPoint.Presentation presentation, string s)
        {

            int xmlCount = presentation.CustomXMLParts.Count;
            // if (xmlCount > 0)
            // {

            //  }
            //  else
            //  {
            Debug.Print("Write: " + s);
            foreach (Microsoft.Office.Core.CustomXMLPart cXML in presentation.CustomXMLParts)
            {
                try
                {
                    cXML.Delete();
                }
                catch (Exception e)
                {

                }

            }
            presentation.CustomXMLParts.Add(s);
            // }
        }

        private static string xmlRead(Microsoft.Office.Interop.PowerPoint.Presentation presentation)
        {

            int xmlCount = presentation.CustomXMLParts.Count;
            if (xmlCount > 0)
            {
                string s = "";
                foreach (Microsoft.Office.Core.CustomXMLPart cXML in presentation.CustomXMLParts)
                {
                    s = cXML.XML;
                    Debug.Print("Read: " + s);
                    // return s;
                }
                return s;

            }
            return "";
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {



        }

        private void ThisAddIn_SlideChange(PowerPoint.SlideRange sld)
        {
            if (sld.Count == 1)
            {
                Debug.Print("Changed Slide" + sld.SlideID);
                //slideDataList.Add(sld.SlideID, new SlideContents());
                myUserControl1.changeSlide(sld.SlideID);
            }

        }

        private async void ThisAddIn_NextSlide(PowerPoint.SlideShowWindow Wn)
        {
            int current = Wn.View.Slide.SlideID;
            String s = await ("https://onpoint.firebaseio.com/monitor/" + key + ".json").PatchJsonAsync(new { currentSlide = current }).ReceiveString();
        }

        private async void ThisAddIn_StartSlideshow(PowerPoint.SlideShowWindow sw)
        {
            key = myUserControl1.getKey();
            String s = await ("https://onpoint.firebaseio.com/shows/" + key + ".json").PatchJsonAsync(slideDataList).ReceiveString();
            Debug.Print(s + "\n");
        }

        private void ThisAddIn_afterPresentationOpen(PowerPoint.Presentation p)
        {

            string s = xmlRead(p);
            if (!s.Equals(""))
            {
                String c = Deserialize<String>(s);
                classHolder ch = Newtonsoft.Json.JsonConvert.DeserializeObject<classHolder>(c);
                Dictionary<int, SlideContents> dd = ch.slideDataList;
                if (dd.Count > 0)
                {
                    slideDataList = dd;
                    myUserControl1.changeDict(dd);
                    //changeSlide(p.sl);
                }
                myUserControl1.changeKey(ch.key);
            }
        }

        private void ThisAddIn_beforePresentationSave(PowerPoint.Presentation p, ref bool b)
        {


            String c = Serialize<String>(Newtonsoft.Json.JsonConvert.SerializeObject(new classHolder( slideDataList,key )));
            xmlWrite(p, c);
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);




            slideDataList = new Dictionary<int, SlideContents>();

            this.Application.SlideShowNextSlide += new PowerPoint.EApplication_SlideShowNextSlideEventHandler(ThisAddIn_NextSlide);
            this.Application.PresentationNewSlide += new PowerPoint.EApplication_PresentationNewSlideEventHandler(Application_PresentationNewSlide);
            this.Application.SlideSelectionChanged += new PowerPoint.EApplication_SlideSelectionChangedEventHandler(ThisAddIn_SlideChange);
            this.Application.SlideShowBegin += new PowerPoint.EApplication_SlideShowBeginEventHandler(ThisAddIn_StartSlideshow);
            this.Application.AfterPresentationOpen += new PowerPoint.EApplication_AfterPresentationOpenEventHandler(ThisAddIn_afterPresentationOpen);
            this.Application.PresentationBeforeSave += new PowerPoint.EApplication_PresentationBeforeSaveEventHandler(ThisAddIn_beforePresentationSave);


            myUserControl1 = new UserControl1(slideDataList);
            myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl1, "My Task Pane");
            myCustomTaskPane.Visible = true;
        }

        void Application_PresentationNewSlide(PowerPoint.Slide Sld)
        {
            //PowerPoint.Shape textBox = Sld.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 50);
            //textBox.TextFrame.TextRange.InsertAfter("This text was added by using code.");
            Debug.Print("Made new slide: " + Sld.SlideID);
            slideDataList.Add(Sld.SlideID, new SlideContents());
        }







        public static string Serialize<T>(T value)
        {

            if (value == null)
            {
                return null;
            }

            XmlSerializer serializer = new XmlSerializer(typeof(T));

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = new UnicodeEncoding(false, false); // no BOM in a .NET string
            settings.Indent = false;
            settings.OmitXmlDeclaration = false;

            using (StringWriter textWriter = new StringWriter())
            {
                using (XmlWriter xmlWriter = XmlWriter.Create(textWriter, settings))
                {
                    serializer.Serialize(xmlWriter, value);
                }
                return textWriter.ToString();
            }
        }

        public static T Deserialize<T>(string xml)
        {

            if (string.IsNullOrEmpty(xml))
            {
                return default(T);
            }

            XmlSerializer serializer = new XmlSerializer(typeof(T));

            XmlReaderSettings settings = new XmlReaderSettings();
            // No settings need modifying here

            using (StringReader textReader = new StringReader(xml))
            {
                using (XmlReader xmlReader = XmlReader.Create(textReader, settings))
                {
                    return (T)serializer.Deserialize(xmlReader);
                }
            }


        }

        public class classHolder
        {
            public Dictionary<int, SlideContents> slideDataList;
            public String key = "default";
            public classHolder(Dictionary<int, SlideContents> sd, String k)
            {
                slideDataList = sd;
                key = k;
            }
        }
    }
    }
#endregion