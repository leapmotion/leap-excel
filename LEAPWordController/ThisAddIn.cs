using System;
using GestureLib;
using Leap;
using Microsoft.Office.Interop.Word;

namespace LEAPWordController
{
    public partial class ThisAddIn
    {
        private Controller controller;
        private GestureListener listener;
        private DateTime LastGesture;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            listener = new GestureListener(1500);
            listener.onGesture += listener_onGesture;

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (controller == null) return;
            StopLeap();
        }

        public void StartLeap()
        {
            LastGesture = DateTime.Now.AddSeconds(-1);
            controller = new Controller(listener);
        }

        public void StopLeap()
        {

            controller.RemoveListener(listener);
            controller.Dispose();
        }

        void listener_onGesture(GestureLib.Gesture gesture)
        {
            //string gestures = "";
            foreach (GestureLib.Gesture.Direction direction in gesture.directions)
            {
                if (DateTime.Now - LastGesture <= new TimeSpan(0, 0, 1)) return;  //Max once every second

                if (direction.ToString() == "Right")
                {
                    Application.ActiveWindow.PageScroll(1,0);
                        LastGesture = DateTime.Now;
                 }
                if (direction.ToString() == "Left")
                {
                    Application.ActiveWindow.PageScroll(0, 1);
                    LastGesture = DateTime.Now;
                }
            }
            //Console.WriteLine("gestured " + gestures + " with " + gesture.fingers + " fingers.");
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
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
        }
        
        #endregion
    }
}
