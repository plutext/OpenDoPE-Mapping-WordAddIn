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
using System;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using XmlMappingTaskPane.Controls;

namespace XmlMappingTaskPane.Forms
{
    public partial class FormAddPart : Form
    {
        private enum CurrentPage { Intro, FromFile, FromString, FromXSD };
        private CurrentPage m_currentPage = CurrentPage.Intro;

        private string m_strXml;

        public FormAddPart()
        {
            InitializeComponent();
        }

        public string XmlString
        {
            get
            {
                return m_strXml;
            }
        }

        private void buttonNext_Click(object sender, EventArgs e)
        {
            if (m_currentPage == CurrentPage.Intro)
            {
                switch (wizardIntroduction.UserChoice)
                {
                    case WizardIntroduction.RadioSelection.FromFile:
                        wizardFromFile.Visible = true;
                        m_currentPage = CurrentPage.FromFile;
                        break;
                    case WizardIntroduction.RadioSelection.FromString:
                        wizardFromString.Visible = true;
                        m_currentPage = CurrentPage.FromString;
                        wizardFromString.SetFocus();
                        break;
                    default:
                        break;
                }

                wizardIntroduction.Visible = false;
                buttonBack.Enabled = true;
                buttonNext.Text = Properties.Resources.Finish;
            }
            else
            {
                //finish up
                if (m_currentPage == CurrentPage.FromFile)
                {
                    //validate file
                    //make sure file actually exists before continuing
                    if (!File.Exists(wizardFromFile.FilePath))
                    {
                        ControlBase.ShowErrorMessage(this, Properties.Resources.FileNotFound);
                        return;
                    }

                    //verify that the file contains valid xml, if not, tell them
                    XmlDocument xdoc = new XmlDocument();
                    xdoc.XmlResolver = null;
                    try
                    {
                        xdoc.Load(wizardFromFile.FilePath);
                    }
                    catch (XmlException ex)
                    {
                        ControlBase.ShowErrorMessage(this, string.Format(CultureInfo.CurrentCulture, Properties.Resources.FileNotValidXml, ex.Message));
                        return;
                    }

                    //send back the choice and file contents
                    m_strXml = xdoc.OuterXml;

                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else if (m_currentPage == CurrentPage.FromString)
                {
                    //validate string
                    //strip tabs and carriage returns from the string
                    string strTempXMLString = wizardFromString.XmlString;
                    strTempXMLString = strTempXMLString.Replace(Environment.NewLine, "");
                    strTempXMLString = strTempXMLString.Replace("\t", "");
                    
                    //make sure XML is well formed
                    XmlDocument xdoc = new XmlDocument();
                    try
                    {
                        xdoc.LoadXml(strTempXMLString);
                    }
                    catch (XmlException ex)
                    {
                        ControlBase.ShowErrorMessage(this, string.Format(CultureInfo.CurrentCulture, Properties.Resources.StringNotValidXml, ex.Message));
                        return;
                    }

                    //send back the choice and the XML
                    m_strXml = xdoc.OuterXml;
                    xdoc = null;

                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
            }
        }

        private void buttonBack_Click(object sender, EventArgs e)
        {
            if (m_currentPage == CurrentPage.FromFile)
            {
                wizardFromFile.Visible = false;
            }
            else if (m_currentPage == CurrentPage.FromString)
            {
                wizardFromString.Visible = false;
            }

            m_currentPage = CurrentPage.Intro;
            wizardIntroduction.Visible = true;

            buttonNext.Text = Properties.Resources.Next;
            buttonBack.Enabled = false;
        }
    }
}
