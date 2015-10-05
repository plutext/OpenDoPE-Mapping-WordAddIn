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
using System.Diagnostics;
using System.Globalization;
using System.Security;
using Microsoft.Win32;

namespace XmlMappingTaskPane
{
    static class SchemaLibrary
    {
        /// <summary>
        /// Get the alias for the schema.
        /// </summary>
        /// <param name="strNamespace">A string specifying the root namespace of the schema.</param>
        /// <param name="intLCID">An integer specifying the current UI language in LCID format.</param>
        /// <returns></returns>
        public static string GetAlias(string strNamespace, int intLCID)
        {
            try
            {
                //try to get the HKLM hive
                RegistryKey regHKLMKey = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Schema Library\" + strNamespace + @"\0", false);

                if (regHKLMKey != null && string.Equals(regHKLMKey.GetValue("Key").ToString(), strNamespace, StringComparison.CurrentCulture))
                {
                    //this is the key, use it
                    RegistryKey regAlias = regHKLMKey.OpenSubKey(@"Alias", false);
                    string strHKLMName = (string)regAlias.GetValue(intLCID.ToString(CultureInfo.InvariantCulture), string.Empty);

                    //if it's non empty, return it
                    if (!string.IsNullOrEmpty(strHKLMName))
                        return strHKLMName;

                    //check for a culture-invariant one
                    strHKLMName = (string)regAlias.GetValue("0", string.Empty);

                    //if it's non empty, return it
                    if (!string.IsNullOrEmpty(strHKLMName))
                        return strHKLMName;
                }
            }
            catch (SecurityException ex)
            {
                Debug.WriteLine("Failed to use HKLM: " + ex.Message);
            }

            try
            {
                //HKLM was no good, try HKCU
                RegistryKey regHKCUKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Schema Library\" + strNamespace + @"\0", false);

                if (regHKCUKey != null && string.Equals(regHKCUKey.GetValue("Key").ToString(), strNamespace, StringComparison.CurrentCulture))
                {
                    //this is the key, use it
                    RegistryKey regAlias = regHKCUKey.OpenSubKey(@"Alias", false);

                    if (regAlias == null)
                        return string.Empty;

                    string strHKCUName = (string)regAlias.GetValue(intLCID.ToString(CultureInfo.InvariantCulture), string.Empty);

                    //if it's non empty, return it
                    if (!string.IsNullOrEmpty(strHKCUName))
                        return strHKCUName;

                    //check for a culture-invariant one
                    strHKCUName = (string)regAlias.GetValue("0", string.Empty);

                    //if it's non empty, return it
                    if (!string.IsNullOrEmpty(strHKCUName))
                        return strHKCUName;
                }
            }
            catch (SecurityException ex)
            {
                Debug.WriteLine("Failed to use HKCU: " + ex.Message);
            }

            return string.Empty;
        }

        /// <summary>
        /// Set the alias for the schema
        /// </summary>
        /// <param name="strNamespace">A string specifying the root namespace of the schema.</param>
        /// <param name="strValue">A string specifying the alias.</param>
        /// <param name="intLCID">An integer specifying the current UI language in LCID format.</param>
        /// <returns>True if the alias was saved in the registry, False otherwise.</returns>
        public static bool SetAlias(string strNamespace, string strValue, int intLCID)
        {
            try
            {
                //HKLM was no good, try HKCU
                RegistryKey regHKCUKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Schema Library\" + strNamespace + @"\0", true);

                if (regHKCUKey == null)
                {
                    //create it
                    regHKCUKey = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Schema Library\" + strNamespace + @"\0", RegistryKeyPermissionCheck.ReadWriteSubTree);
                    regHKCUKey.SetValue("Key", strNamespace, RegistryValueKind.String);
                }


                if (regHKCUKey != null && string.Equals(regHKCUKey.GetValue("Key").ToString(), strNamespace, StringComparison.CurrentCulture))
                {
                    //this is the key, use it
                    RegistryKey regAlias = regHKCUKey.OpenSubKey(@"Alias", true);

                    if (regAlias == null)
                    {
                        regAlias = regHKCUKey.CreateSubKey(@"Alias");
                    }

                    regAlias.SetValue(intLCID.ToString(CultureInfo.InvariantCulture), strValue, RegistryValueKind.String);
                    return true;
                }
            }
            catch (SecurityException ex)
            {
                Debug.WriteLine("Failed to write to HKCU: " + ex.Message);
            }

            return false;
        }
    }
}
