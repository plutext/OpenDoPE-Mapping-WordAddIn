//Copyright (c) Microsoft Corporation.  All rights reserved.
/*
 *  From http://xmlmapping.codeplex.com/license:

    Microsoft Platform and Application License

    This license governs use of the accompanying software. If you use the software, you accept this license. If you 
    do not accept the license, do not use the software.

    1. Definitions
    The terms “reproduce,” “reproduction,” “derivative works,” and “distribution” have the same meaning here as 
    under U.S. copyright law.
    A “contribution” is the original software, or any additions or changes to the software.
    A “contributor” is any person that distributes its contribution under this license.
    “Licensed patents” are a contributor’s patent claims that read directly on its contribution.

    2. Grant of Rights
    (A) Copyright Grant- Subject to the terms of this license, including the license conditions and limitations in 
    section 3, each contributor grants you a non-exclusive, worldwide, royalty-free copyright license to reproduce 
    its contribution, prepare derivative works of its contribution, and distribute its contribution or any derivative 
    works that you create.
    (B) Patent Grant- Subject to the terms of this license, including the license conditions and limitations in section 
    3, each contributor grants you a non-exclusive, worldwide, royalty-free license under its licensed patents to 
    make, have made, use, sell, offer for sale, import, and/or otherwise dispose of its contribution in the software 
    or derivative works of the contribution in the software.

    3. Conditions and Limitations
    (A) No Trademark License- This license does not grant you rights to use any contributors’ name, logo, or
    trademarks.
    (B) If you bring a patent claim against any contributor over patents that you claim are infringed by the
    software, your patent license from such contributor to the software ends automatically.
    (C) If you distribute any portion of the software, you must retain all copyright, patent, trademark, and
    attribution notices that are present in the software.
    (D) If you distribute any portion of the software in source code form, you may do so only under this license
    by including a complete copy of this license with your distribution. If you distribute any portion of the 
    software in compiled or object code form, you may only do so under a license that complies with this license.
    (E) The software is licensed “as-is.” You bear the risk of using it. The contributors give no express warranties, 
    guarantees or conditions. You may have additional consumer rights under your local laws which this license 
    cannot change. To the extent permitted under your local laws, the contributors exclude the implied warranties 
    of merchantability, fitness for a particular purpose and non-infringement.
    (F) Platform Limitation- The licenses granted in sections 2(A) & 2(B) extend only to the software or derivative
    works that you create that (1) run on a Microsoft Windows operating system product, and (2) operate with 
    Microsoft Word.
 */
using System.Collections.Generic;
using System.Globalization;
using Microsoft.Office.Tools;
using Microsoft.Win32;
using XmlMappingTaskPane.Controls;
using Word = Microsoft.Office.Interop.Word;
using NLog;
using NLog.Config;
using System;
using System.Text;

namespace XmlMappingTaskPane
{
    public partial class ThisAddIn
    {
        private ApplicationEvents m_appEvents; // hang on to a reference to all application-level events, so they can't go out of scope
        private IDictionary<Word.Window, CustomTaskPane> m_dicTaskPanes = new Dictionary<Word.Window, CustomTaskPane>(); //key = Word Window object; value = ctp (if any) for that window

        #region Startup/Shutdown

        static Logger log;

        static ThisAddIn()
        {
            NLog.Config.LoggingConfiguration config = new NLog.Config.LoggingConfiguration();
            NLog.Targets.Target t;
            //System.Diagnostics.Trace.WriteLine(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
            // eg file:///C:/Users/jharrop/Documents/Visual Studio 2010/Projects/com.plutext.search/com.plutext.search.main/bin/Debug/com.plutext.search.main.DLL
            if (System.Reflection.Assembly.GetExecutingAssembly().CodeBase.Contains("Debug"))
            {
                t = new NLog.Targets.DebuggerTarget();
                ((NLog.Targets.DebuggerTarget)t).Layout = "${callsite} ${message}";
            }
            else
            {
                t = new NLog.Targets.FileTarget();
                ((NLog.Targets.FileTarget)t).FileName = System.IO.Path.GetTempPath() + "plutext.txt";
                //// Win 7:  C:\Users\jharrop\AppData\Local\Temp\
                //System.Diagnostics.Trace.WriteLine("TEMP: " + System.IO.Path.GetTempPath());
                ((NLog.Targets.FileTarget)t).AutoFlush = true;
            }
            //ILayout layout = new NLog.Layout("${longdate} ${callsite} ${level} ${message}");
            //NLog.LayoutCollection lc = new NLog.LayoutCollection();
            //lc.Add(layout);
            ////t.GetLayouts().Add(layout);
            //t.PopulateLayouts(lc);

            config.AddTarget("ds", t);
            config.LoggingRules.Add(new NLog.Config.LoggingRule("*", LogLevel.Trace, t));
            LogManager.Configuration = config;
            log = LogManager.GetLogger("OpenDoPEFriendly");
            log.Info("Logging operational.");
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            log.Info("ThisAddIn_Startup");
            
            // start catching application-level events
            m_appEvents = new ApplicationEvents(
                //Globals.Ribbons.RibbonMapping, m_dicTaskPanes);
                Ribbon.ribbon, m_dicTaskPanes);

            // perform initial registry setup
            try
            {
                if (Registry.CurrentUser.OpenSubKey(System.Configuration.ConfigurationManager.AppSettings["Registry.CurrentUser.SubKey"]) == null)
                {
                    Registry.CurrentUser.CreateSubKey(System.Configuration.ConfigurationManager.AppSettings["Registry.CurrentUser.SubKey"], RegistryKeyPermissionCheck.ReadWriteSubTree);

                    using (RegistryKey rk = Registry.CurrentUser.OpenSubKey(System.Configuration.ConfigurationManager.AppSettings["Registry.CurrentUser.SubKey"], true))
                    {
                        string val = System.Configuration.ConfigurationManager.AppSettings["Ribbon.Button.XMLOptions.Value"];
                        if (String.IsNullOrWhiteSpace(val))
                        {
                            rk.SetValue("Options", ControlTreeView.cOptionsShowAttributes + ControlTreeView.cOptionsAutoSelectNode);
                        }
                        else
                        {
                            int intVal;
                            if (int.TryParse(val, out intVal))
                            {
                                rk.SetValue("Options", intVal);
                            }
                            else
                            {
                                rk.SetValue("Options", ControlTreeView.cOptionsShowAttributes + ControlTreeView.cOptionsAutoSelectNode);
                            }
                        }
                    }

                    //set up the schema library entries
                    int iLocale = int.Parse(Properties.Resources.Locale, CultureInfo.InvariantCulture);
                    SchemaLibrary.SetAlias("http://schemas.openxmlformats.org/package/2006/metadata/core-properties", Properties.Resources.CoreFilePropertiesName, iLocale);
                    SchemaLibrary.SetAlias("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", Properties.Resources.ExtendedFilePropertiesName, iLocale);
                    SchemaLibrary.SetAlias("http://schemas.microsoft.com/office/2006/coverPageProps", Properties.Resources.CoverPagePropertiesName, iLocale);
                }
            }
            catch (System.Security.SecurityException se)
            {
                log.Info(se);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #endregion

        /// <summary>
        /// Start handling the visibility events for a particular taskpane.
        /// </summary>
        /// <param name="ctp">A CustomTaskPane specifying the taskpane whose events we want to handle.</param>
        internal void ConnectTaskPaneEvents(CustomTaskPane ctp)
        {
            m_appEvents.ConnectTaskPaneEvents(ctp);
        }

        /// <summary>
        /// Update all active taskpanes with the new settings.
        /// </summary>
        /// <param name="newOptions">An integer specifying the settings to be applied.</param>
        internal static void UpdateSettings(int newOptions)
        {
            foreach (CustomTaskPane ctp in Globals.ThisAddIn.CustomTaskPanes)
            {
                ((Controls.ControlMain)ctp.Control).RefreshSettings(newOptions);
            }
        }

        /// <summary>
        /// A list of all task panes in the document. Key = Word Window object; Value = CustomTaskPane object
        /// </summary>
        internal IDictionary<Word.Window, CustomTaskPane> TaskPaneList
        {
            get
            {
                return m_dicTaskPanes;
            }
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
