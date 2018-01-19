using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using log4net;
using log4net.Config;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]

// <summary> 
// This namespaces if for generic application classes
// </summary>
namespace ShapeExtract.Scripts
{
    /// <summary> 
    /// Used to handle exceptions
    /// </summary>
    public class ErrorHandler
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(ErrorHandler));

        /// <summary> 
        /// Used to produce an error message and create a log record
        /// <example>
        /// <code lang="C#">
        /// ErrorHandler.DisplayMessage(ex);
        /// </code>
        /// </example> 
        /// </summary>
        /// <param name="ex">Represents errors that occur during application execution.</param>
        /// <remarks></remarks>
        public static void DisplayMessage(Exception ex)
        {
            System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(1);
            System.Reflection.MethodBase caller = sf.GetMethod();
            string currentProcedure = (caller.Name).Trim();
            string errorMessageDescription = ex.ToString();
            errorMessageDescription = System.Text.RegularExpressions.Regex.Replace(errorMessageDescription, @"\r\n+", " "); //the carriage returns were messing up my log file
            string msg = "Contact your system administrator. A record has been created in the log file." + Environment.NewLine;
            msg += "Procedure: " + currentProcedure + Environment.NewLine;
            msg += "Description: " + ex.ToString() + Environment.NewLine;
            log.Error("[PROCEDURE]=|" + currentProcedure + "|[USER NAME]=|" + Environment.UserName + "|[MACHINE NAME]=|" + Environment.MachineName + "|[DESCRIPTION]=|" + errorMessageDescription);
            MessageBox.Show(msg, "Unexpected Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="text"></param>
        /// <param name="charValue"></param>
        /// <returns></returns>
        public static string RemoveString(string text, string charValue = ".")
        {
            try
            {
                if (text.Contains(charValue))
                {
                    return text.Substring(0, text.IndexOf("."));
                }
                else
                {
                    return text;
                }
            }
            catch (Exception)
            {
                return text;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static string ValidateString(string text)
        {
            try
            {
                if (string.IsNullOrEmpty(text))
                {
                    text = string.Empty;
                }
                else
                {
                    text = string.Empty;  //fix this
                    //text = text.Replace(vbCr, "").Replace(vbLf, "");
                    //text = text.Replace(",", ";");
                    //text = text.Trim;
                }

                return text;
            }
            catch (Exception)
            {
                return text;
            }

        }

    }
}