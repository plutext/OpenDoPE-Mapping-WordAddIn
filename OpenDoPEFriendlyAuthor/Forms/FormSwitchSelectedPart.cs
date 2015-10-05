/*
 *  OpenDoPE authoring Word AddIn
    Copyright (C) Plutext Pty Ltd, 2012
 * 
    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools;

namespace XmlMappingTaskPane.Forms
{
    public partial class FormSwitchSelectedPart : Form
    {
        public FormSwitchSelectedPart()
        {
            InitializeComponent();

            //CustomTaskPane ctpPaneForThisWindow = Utilities.FindTaskPaneForCurrentWindow();
            //Controls.ControlMain ccm = (Controls.ControlMain)ctpPaneForThisWindow.Control;

            this.controlPartList.controlMain = ccm;
        }

        private Controls.ControlMain _ccm;
        public Controls.ControlMain ccm
        {
            get { return _ccm; }
            set { _ccm = value; }
        }

        /// <summary>
        /// Just want to hide the form.  
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormSwitchSelectedPart_FormClosing(object sender, FormClosingEventArgs e)
        {
            //If the user is simply hitting the X in the window the form hides, 
            // if any thing else such as task manager, application.exit, 
            // or windows shutdown the form is properly closed, since the 
            // return statement would be executed.
            // From http://stackoverflow.com/questions/2021681/c-sharp-hide-form-instead-of-close
            if (e.CloseReason != CloseReason.UserClosing) return;
            e.Cancel = true; // this cancels the close event.
            Hide(); 
        }

        private void buttonHide_Click(object sender, System.EventArgs e)
        {
            Hide();
        }
    }
}

