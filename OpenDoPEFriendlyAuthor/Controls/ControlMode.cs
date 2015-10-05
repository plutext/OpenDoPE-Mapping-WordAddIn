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
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NLog;


namespace XmlMappingTaskPane.Controls
{
    public partial class ControlMode :  Controls.ControlBase
    {
        static Logger log = LogManager.GetLogger("ControlMode");

        public ControlMode()
        {
            InitializeComponent();
        }

        public Boolean isModeBind()
        {
            return radioModeBind.Checked;
        }

        public Boolean isModeCondition()
        {
            return radioModeCondition.Checked;
        }

        public Boolean isModeRepeat()
        {
            return radioModeRepeat.Checked;
        }

    }
}
