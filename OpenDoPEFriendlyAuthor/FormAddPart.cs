//Copyright (c) Microsoft Corporation.  All rights reserved.
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
