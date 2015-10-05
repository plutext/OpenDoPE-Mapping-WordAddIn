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

namespace OpenDoPEModel
{
    public partial class FormResponses : Form
    {
        private responseFixed responseFixed;

        public FormResponses(responseFixed responseFixed)
        {
            InitializeComponent();
            this.responseFixed = responseFixed;

            // Pre-populate with true/false
            this.dataGridView1.Rows.Add(2);
            this.dataGridView1.Rows[0].Cells[0].Value = "true";
            this.dataGridView1.Rows[0].Cells[1].Value = "yes";
            this.dataGridView1.Rows[1].Cells[0].Value = "false";
            this.dataGridView1.Rows[1].Cells[1].Value = "no";
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            if (this.dataGridView1.Rows.Count < 3) // auto last row
            {
                Mbox.ShowSimpleMsgBoxError("You must provide at least 2 choices!");
                return;
            }

            int last = this.dataGridView1.Rows.Count -1;
            foreach (DataGridViewRow row in this.dataGridView1.Rows)
            {
                // Last row is added automatically
                if (row == this.dataGridView1.Rows[last]) continue;

                DataGridViewCell c = row.Cells[0];
                if (string.IsNullOrWhiteSpace((string)c.Value))
                {
                    Mbox.ShowSimpleMsgBoxError("You must enter data in each cell!");
                    return;
                }
                c = row.Cells[1];
                if (string.IsNullOrWhiteSpace((string)c.Value))
                {
                    Mbox.ShowSimpleMsgBoxError("You must enter data in each cell!");
                    return;
                }
            }

            // OK

            foreach (DataGridViewRow row in this.dataGridView1.Rows)
            {
                if (row == this.dataGridView1.Rows[last]) continue;

                responseFixedItem item = new OpenDoPEModel.responseFixedItem();

                item.value = (string)row.Cells[0].Value;
                item.label = (string)row.Cells[1].Value;

                responseFixed.item.Add(item);
            }

            if (this.radioButtonYes.Checked)
            {
                responseFixed.canSelectMany = true;
            }
            else
            {
                responseFixed.canSelectMany = false;
            }
            responseFixed.canSelectManySpecified = true; // have to set this!

            this.Close();

        }

        public appearanceType getAppearanceType()
        {
            // AppearanceType
            if (this.radioButtonAppearanceFull.Checked)
            {
                return appearanceType.full;
            }
            else if (this.radioButtonAppearanceCompact.Checked)
            {
                return appearanceType.compact;
            }
            else if (this.radioButtonAppearanceMinimal.Checked)
            {
                return appearanceType.minimal;
            }
            return appearanceType.full;
        }

        public string getDataType()
        {
            // TODO: values to be validated against selected data type
            if (this.radioTypeBoolean.Checked)
            {
                return "boolean";
            }
            else if (this.radioTypeDate.Checked)
            {
                return "date";
            }
            else if (this.radioTypeNumber.Checked)
            {
                return "decimal"; // allows integer
            }
            else if (this.radioTypeText.Checked)
            {
                return "string";
            }
            else
            {   // default
                return "string";
            }
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            Mbox.ShowSimpleMsgBoxError("Not implemented yet!");
        }
    }
}
