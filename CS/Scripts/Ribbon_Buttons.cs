using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShapeExtract.Scripts
{
    class Ribbon_Buttons
    {

        public static void ExportShapeValues()
        {

        }

        public static void OpenFile()
        {
            string filePath = Properties.Settings.Default.App_FileExport; 
            try
            {
                if (filePath == string.Empty)
                    return;
                var attributes = File.GetAttributes(filePath);
                File.SetAttributes(filePath, attributes | FileAttributes.ReadOnly);
                System.Diagnostics.Process.Start(filePath);

            }
            catch (System.ComponentModel.Win32Exception)
            {
                MessageBox.Show("No application is assicated to this file type." + Environment.NewLine + Environment.NewLine + filePath, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }
        }

        /// <summary> 
        /// Opens the settings taskpane
        /// </summary>
        /// <remarks></remarks>
        public static void OpenSettings()
        {
            try
            {
                //open settings form

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        /// <summary> 
        /// Opens an as built help file
        /// </summary>
        /// <remarks></remarks>
        public static void OpenReadMe()
        {
            //ErrorHandler.CreateLogRecord();
            System.Diagnostics.Process.Start(Properties.Settings.Default.App_PathReadMe);

        }

        /// <summary> 
        /// Opens an as built help file
        /// </summary>
        /// <remarks></remarks>
        public static void OpenNewIssue()
        {
            //ErrorHandler.CreateLogRecord();
            System.Diagnostics.Process.Start(Properties.Settings.Default.App_PathReportIssue);

        }

    }
}
