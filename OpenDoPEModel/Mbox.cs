//-------------------------------------------------------------------------------------------------
// <copyright company="Microsoft">
//    Author: Matt Scott (mrscott). Copyright (c) Microsoft Corporation.  All rights reserved.
//
//    The use and distribution terms for this software are covered by the
//    Microsoft Limited Permissive License: 
//    http://www.microsoft.com/resources/sharedsource/licensingbasics/limitedpermissivelicense.mspx
//    which can be found in the file license_mslpl.txt at the root of this distribution.
//    By using this software in any fashion, you are agreeing to be bound by
//    the terms of this license.
//
//    You must not remove this notice, or any other, from this software.
//
// </copyright>
//-------------------------------------------------------------------------------------------------

using System;
using System.Collections;
using System.Xml;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using System.Net;
using System.Threading;

namespace OpenDoPEModel
{	
    /// <summary>
    /// Message Box -- Utility methods for launching branded message boxes
    /// </summary>
	public class Mbox
	{  
		public static void ShowSimpleMsgBoxError(string msg)
		{
            MessageBox.Show(msg, "OpenDoPE", MessageBoxButtons.OK, MessageBoxIcon.Error);
		}
        
        public static void ShowSimpleMsgBoxWarning(string msg)
		{
            MessageBox.Show(msg, "OpenDoPE", MessageBoxButtons.OK, MessageBoxIcon.Warning);
		}

		public static void ShowSimpleMsgBoxInfo(string msg)
		{
            MessageBox.Show(msg, "OpenDoPE", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}
	}
}
