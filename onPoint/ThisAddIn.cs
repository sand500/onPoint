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

namespace onPoint
{
    public partial class ThisAddIn
    {
        private UserControl1 myUserControl1;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        private Dictionary<int, SlideContents> slideDataList;

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
            String s = await("https://onpoint.firebaseio.com/monitor/ijpst.json").PatchJsonAsync(new { currentSlide = current }).ReceiveString();

        }
        
        private async void ThisAddIn_StartSlideshow(PowerPoint.SlideShowWindow sw )
        {
            String s = await("https://onpoint.firebaseio.com/shows/ijpst.json").PatchJsonAsync(slideDataList).ReceiveString();
            Debug.Print(s + "\n");
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
            myUserControl1 = new UserControl1(slideDataList);
            myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl1, "My Task Pane");
            myCustomTaskPane.Visible = true;

        

        }

        void Application_PresentationNewSlide(PowerPoint.Slide Sld)
        {
            PowerPoint.Shape textBox = Sld.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 50);
            textBox.TextFrame.TextRange.InsertAfter("This text was added by using code.");
            Debug.Print("Made new slide: " + Sld.SlideID);
            slideDataList.Add(Sld.SlideID, new SlideContents());
        }


    }
    #endregion
}
