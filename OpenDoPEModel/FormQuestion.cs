/*
 * (c) Copyright Plutext Pty Ltd, 2012
 * 
 * All rights reserved.
 * 
 * This source code is the proprietary information of Plutext
 * Pty Ltd, and must be kept confidential.
 * 
 * You may use, modify and distribute this source code only
 * as provided in your license agreement with Plutext.
 * 
 * If you do not have a license agreement with Plutext:
 * 
 * (i) you must return all copies of this source code to Plutext, 
 * or destroy it.  
 * 
 * (ii) under no circumstances may you use, modify or distribute 
 * this source code.
 * 
 */
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
//using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using NLog;

using Office = Microsoft.Office.Core;


namespace OpenDoPEModel
{
    public partial class FormQuestion : Form
    {
        static Logger log = LogManager.GetLogger("OpenDoPE_Wed");

        private Office.CustomXMLPart questionsPart;

        private question q;

        public FormQuestion(Office.CustomXMLPart questionsPart, string xpath, string xpathId)
        {
            InitializeComponent();
            this.questionsPart = questionsPart;

            this.textBoxXPath.Text = xpath;
            this.textBoxQID.Text = "q" + xpathId;
        }

        private void buttonNext_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(this.textBoxQID.Text)
                || string.IsNullOrWhiteSpace(this.textBoxQText.Text))
            {
                Mbox.ShowSimpleMsgBoxError("Required data missing!");
                return;
            }

            q = new question();
            q.id = this.textBoxQID.Text;
            questionText qt = new questionText();
            qt.Value = this.textBoxQText.Text;
            q.text = qt;

            // Responses
            response responses = new response();
            q.response = responses;

            if (this.radioButtonMCYes.Checked)
            {
                // MCQ: display response form
                responseFixed responseFixed = new responseFixed();
                responses.Item = responseFixed;
                FormResponses formResponses = new FormResponses(responseFixed);
                formResponses.ShowDialog();

                // TODO - handle cancel
            }
            else
            {
                // Not MCQ
                responseFree responseFree = new responseFree();
                responses.Item = responseFree;

                // Free - just text for now
                // later, configure type
                responseFree.format = responseFreeFormat.text;
            }

            // Finally, add to part
            //updateQuestionsPart(q);

            this.Close();

        }

        public question getQuestion()
        {
            return q;
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            Mbox.ShowSimpleMsgBoxError("Not implemented yet!");
        }

        public void updateQuestionsPart(question q)
        {

            questionnaire questionnaire = new questionnaire();
            questionnaire.Deserialize(questionsPart.XML, out questionnaire);

            questionnaire.questions.Add(q);

            // Save it in docx
            string result = questionnaire.Serialize();
            log.Info(result);
            CustomXmlUtilities.replaceXmlDoc(questionsPart, result);

        }

    }
}
