using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using GestureLib;
using Leap;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace LEAPExcelController
{
    public partial class ThisAddIn
    {
        private Controller controller;
        private GestureListener listener;
        private DateTime LastGesture;
        private Boolean bIsGrab = false;
        private Excel.Worksheet sheet;
        private Excel.Shape shape;
        //TODO: indikator for on/off evt title af window (latch til eventet når title ændrer sig)


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            listener = new GestureListener(1500);
            listener.onGesture += listener_onGesture;
            listener.onGrab += listener_onGrab;
            listener.onPalmVelocity += listener_onPalmVelocity;
            
            Console.WriteLine("Startup");

            var excelWorksheet = (Excel.Worksheet)Application.ActiveSheet;
            sheet = Application.ActiveSheet;
            int i = sheet.Shapes.Count;
            if (i > 0)
                shape = sheet.Shapes.Item(1);

            //if(Properties.Settings.Default.LeapEnabled)
               StartLeap();

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (controller == null) return;
            StopLeap();
        }

        public void StartLeap()
        {

            Application.Caption =
                Application.Caption + " - LEAP Activated ";
            LastGesture = DateTime.Now.AddSeconds(-1);
            controller = new Controller(listener);

        }

        public void StopLeap()
        {
           // Application.Caption = Application.Caption.Replace(" - LEAP Activated ", "");
            controller.RemoveListener(listener);
            controller.Dispose();
        }

        void listener_onPalmVelocity(Vector vector)
        {
            if (bIsGrab)
            {
                if (sheet != null)
                {
                    //sheet.Cells[1, 1] = vector.x;
                    //sheet.Cells[1, 2] = bIsGrab ? 1 : 0;
                    if (shape != null && vector.x != 0)
                    {
                        if (shape.HasChart == Office.MsoTriState.msoTrue && shape.Chart.ChartArea.Format.ThreeD != null)
                        {
                            //float z = shape.Chart.ChartArea.Format.ThreeD.RotationX;
                            Application.ActiveSheet.Shapes.Item(1).Chart.ChartArea.Format.ThreeD.RotationX = 180 - vector.x / 2;
                        }
                    }
                }
            }
        }

        void listener_onGrab(float strength)
        {
            if (strength > 0.7)
            {
                bIsGrab = true;
                /*
                var excelWorksheet = (Excel.Worksheet)Application.ActiveSheet;
                Excel.Worksheet sheet = Application.ActiveSheet;
                if (sheet != null)
                {
                    int i = sheet.Shapes.Count;
                    if (i > 0)
                    {
                        Excel.Shape s = sheet.Shapes.Item(1);
                        if (s.HasChart == Office.MsoTriState.msoTrue && s.Chart.ChartArea.Format.ThreeD != null)
                        {
                            float z = s.Chart.ChartArea.Format.ThreeD.RotationX;
                            Application.ActiveSheet.Shapes.Item(1).Chart.ChartArea.Format.ThreeD.RotationX = z + 10;
                            Application.ActiveSheet.Shapes.Item(1).Chart.ChartArea.Format.ThreeD.RotationX = z - 10;
                            Application.ActiveSheet.Shapes.Item(1).Chart.ChartArea.Format.ThreeD.RotationX = z;
                        }
                    }
                }*/
            }
            else
            {
                bIsGrab = false;
            }
            //sheet.Cells[1, 2] = strength;
        }

        void listener_onGesture(GestureLib.Gesture gesture)
        {
            string gestures = "";

            foreach (GestureLib.Gesture.Direction direction in gesture.directions)
            {
                     gestures += direction.ToString() + ", ";

                if ((DateTime.Now - LastGesture) <= new TimeSpan(0,0,0,1,0)) return;
                if (gesture.fingers >= 1)
                {
                    var excelWorksheet = (Excel.Worksheet)Application.ActiveSheet;
                    Excel.Worksheet newsheet = null;
                    switch (direction.ToString())
                    {
                        case "Right":
                        case "Left":
                            Excel.Worksheet sheet = Application.ActiveSheet;
                            if (sheet != null) {
                                int i = sheet.Shapes.Count;
                                if (i > 0) {
                                    Excel.Shape s = sheet.Shapes.Item(1);
                                    if (s.HasChart == Office.MsoTriState.msoTrue && s.Chart.ChartArea.Format.ThreeD != null) {
                                        float z = s.Chart.ChartArea.Format.ThreeD.RotationX;
                                        Application.ActiveSheet.Shapes.Item(1).Chart.ChartArea.Format.ThreeD.RotationX = z + 10;
                                    }
                                }
                            }
                            break;
                        case "Up":
                            if (Application.ActiveWindow != null)
                                Application.ActiveWindow.Zoom = Application.ActiveWindow.Zoom + 5;
                            break;
                        case "Down":
                            if (Application.ActiveWindow != null)
                                Application.ActiveWindow.Zoom = Application.ActiveWindow.Zoom - 5;
                            break;
                    }
                    if (newsheet != null)
                    {
                        LastGesture = DateTime.Now;
                        newsheet.Activate();
                    }
                    return;
                }
                if (gesture.fingers == 1 || gesture.fingers == 2)
                {
                    ColumnStep = 3;
                    RowStep = 5;
                    switch (direction.ToString())
                    {
                        case "Right": Application.ActiveWindow.SmallScroll(null, null, ColumnStep,null);
                            continue;
                        case "Left": Application.ActiveWindow.SmallScroll(null, null, null, ColumnStep);
                            continue;
                        case "Up": Application.ActiveWindow.SmallScroll(null, RowStep, null, null);
                            continue;
                        case "Down": Application.ActiveWindow.SmallScroll(RowStep, null, null, null);
                            continue;
                    }
                    LastGesture = DateTime.Now;
                }
                else if (gesture.fingers >= 3)
                {
                    switch (direction.ToString())
                    {
                        case "Right": Application.ActiveWindow.LargeScroll(0, 0, 1, 0);
                            continue;
                        case "Left": Application.ActiveWindow.LargeScroll(0, 0, 0, 1);
                            continue;
                        case "Up": Application.ActiveWindow.LargeScroll(0, 1, 0, 0);
                            continue;
                        case "Down": Application.ActiveWindow.LargeScroll(1, 0, 0, 0);
                            continue;
                    }
                    LastGesture = DateTime.Now;
                }

            }
            Console.WriteLine("gestured " + gestures + " with " + gesture.fingers + " fingers.");
        }

        private int ColumnStep { get; set; }
        private int RowStep { get; set; }



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
