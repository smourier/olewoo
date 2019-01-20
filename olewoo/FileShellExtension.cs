/*
 * Cheap n cheerful non-com(shell extension) file registration for context menu.
 * 
 * From
 * 
 * http://www.codeproject.com/KB/shell/SimpleContextMenu.aspx
 * 
 */
using System;
using Microsoft.Win32;

namespace olewoo_cs
{
    static class FileShellExtension
    {
        public static void Register(string fileType,
               string shellKeyName, string menuText, string menuCommand)
        {
            try
            {

                // create path to registry location

                string regPath = string.Format(@"{0}\shell\{1}",
                                               fileType, shellKeyName);

                // add context menu to the registry

                using (RegistryKey key =
                       Registry.ClassesRoot.CreateSubKey(regPath))
                {
                    key.SetValue(null, menuText);
                }

                // add command that is invoked to the registry

                using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(
                    string.Format(@"{0}\command", regPath)))
                {
                    key.SetValue(null, menuCommand);
                }
            }
            catch (System.UnauthorizedAccessException)
            {
                System.Windows.Forms.MessageBox.Show("Please run as administrator to perform this task.", "Insufficient Privileges.", System.Windows.Forms.MessageBoxButtons.OK);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Error:", System.Windows.Forms.MessageBoxButtons.OK);
            }
        }

        public static void Unregister(string fileType, string shellKeyName)
        {
            try
            {
                if (string.IsNullOrEmpty(fileType) ||
                    string.IsNullOrEmpty(shellKeyName)) return;

                // path to the registry location

                string regPath = string.Format(@"{0}\shell\{1}",
                                               fileType, shellKeyName);

                // remove context menu from the registry

                Registry.ClassesRoot.DeleteSubKeyTree(regPath);
            }
            catch (System.UnauthorizedAccessException)
            {
                System.Windows.Forms.MessageBox.Show("Please run as administrator to perform this task.", "Insufficient Privileges.", System.Windows.Forms.MessageBoxButtons.OK);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Error:", System.Windows.Forms.MessageBoxButtons.OK);
            }
        }
    }
}
