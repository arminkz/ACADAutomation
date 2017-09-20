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

//Import AutoCAD Lib
using AutoCAD;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading;
using Microsoft.Win32;

namespace ACADAutomation
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        AutoCAD.AcadApplication app = default(AcadApplication);


        private void Test1(object sender, RoutedEventArgs e)
        {
            //Start AutoCAD
            Process acadProc = new Process();
            acadProc.StartInfo.FileName = "C:/Program Files/Autodesk/AutoCAD 2018/acad.exe";
            acadProc.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;
            acadProc.Start();

            if (!acadProc.WaitForInputIdle(300000))
                throw new ApplicationException("AutoCAD takes too much time to start.");

            this.Title = "[Waiting for AutoCAD ...]";

            while (true)
            {
                try
                {
                    // Getting running AutoCAD instance by Marshalling by passing Programmatic ID as a string, AutoCAD.Application is the Programmatic ID for AutoCAD.
                    app = (AcadApplication)Marshal.GetActiveObject("AutoCAD.Application");
                    break;
                }
                catch (COMException ex)
                {
                    const uint MK_E_UNAVAILABLE = 0x800401e3;
                    if ((uint)ex.ErrorCode != MK_E_UNAVAILABLE) throw;
                    Thread.Sleep(1000);
                }
            }

            //wait for AutoCAD
            while (true)
            {
                try
                {
                    AcadState state = app.GetAcadState();
                    if (state.IsQuiescent) break;
                    Thread.Sleep(1000);
                }
                catch
                {
                    Thread.Sleep(1000);
                }
            }

            //Create new Document
            app.Documents.Add();

            this.Title = "[Drawing ...]";

            //This is the main Drawing Part
            double[] CenterOfCircle = new double[3];
            CenterOfCircle[0] = 5;
            CenterOfCircle[1] = 5;
            CenterOfCircle[2] = 0;
            double RadiusOfCircle = 1;
            app.ActiveDocument.ModelSpace.AddCircle(CenterOfCircle, RadiusOfCircle);

            double[] p1 = new double[3] { 6, 5, 0 };
            double[] p2 = new double[3] { 7, 5, 0 };
            app.ActiveDocument.ModelSpace.AddLine(p1, p2);

            double[] c2 = new double[3] { 8, 5, 0 };
            app.ActiveDocument.ModelSpace.AddCircle(c2, 1);

            app.ActiveDocument.ModelSpace.AddLine(new double[] { 9, 5, 0 }, new double[] { 10, 5, 0 });

            app.ActiveDocument.ModelSpace.AddCircle(new double[] { 11, 5, 0 }, 1);

            this.Title = "[Saving Output ...]";

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "AutoCAD Drawing File (*.dwg)|*.dwg|All files (*.*)|*.*";
            if (sfd.ShowDialog() == true)
            {
                string filename = sfd.FileName;
                app.ActiveDocument.SaveAs(filename);
            }


            this.Title = "[Finalizing ...]";

            //Close AutoCAD
            app.ActiveDocument.Close();
            app.Quit();

            //Close Me
            Environment.Exit(0);
        }

        private void WinLoaded(object sender, RoutedEventArgs e)
        {
            
        }
    }
}
