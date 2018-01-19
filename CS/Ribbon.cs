using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using ShapeExtract.Scripts;

namespace ShapeExtract
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ShapeExtract.Ribbon.xml");
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            //AssemblyInfo.SetAddRemoveProgramsIcon("ExcelAddin.ico");
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Assigns the value to an application setting
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <returns>A method that returns true or false if the control is enabled </returns> 
        public void OnAction(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "btnExtractShapes":
                        ExportShapeValues();
                        break;
                    case "btnOpenFolder":
                        OpenFile();
                        break;
                    case "btnSettings":
                        OpenSettings();
                        break;
                    case "btnOpenReadMe":
                        OpenReadMe();
                        break;
                    case "btnOpenNewIssue":
                        OpenNewIssue();
                        break;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }

        }

        /// <summary> 
        /// Assigns text to a label on the ribbon from the xml file
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <returns>A method that returns a string for a label. </returns> 
        public string GetLabelText(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "tabShapeExtract":
                        if (Application.ProductVersion.Substring(0, 2) == "15") //for Excel 2013
                        {
                            return AssemblyInfo.Title.ToUpper();
                        }
                        else
                        {
                            return AssemblyInfo.Title;
                        }
                    case "txtCopyright":
                        return "© " + AssemblyInfo.Copyright;
                    case "txtDescription":
                        return AssemblyInfo.Title.Replace("&", "&&") + " " + AssemblyInfo.AssemblyVersion;
                    case "txtReleaseDate":
                        System.DateTime dteCreateDate = Properties.Settings.Default.App_ReleaseDate;
                        //return dteCreateDate.ToString("dd-MMM-yyyy hh:mm tt");
                        return string.Empty;
                    default:
                        return string.Empty;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return string.Empty;
            }
        }



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
