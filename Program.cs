using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;

namespace data_extractor
{
    class Program
    {
        static void Main(string[] args)
        {

            Console.WriteLine("PST file path");
            var pstFilePath = Console.ReadLine();

            Console.WriteLine("Output attachments path");
            var outputPath = Console.ReadLine();

            Application outlookApp = new Application();
            NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");

            // Open the PST file
            outlookNamespace.Logon("", "", false, false);
            outlookNamespace.AddStore(pstFilePath);

            // Get the root folder of the PST file
            MAPIFolder rootFolder = outlookNamespace.Folders.GetLast() as MAPIFolder;

            foreach (MAPIFolder folder in rootFolder.Folders)
            {
                bool exists = Directory.Exists(Path.Combine(outputPath, folder.Name));

                if (!exists)
                   Directory.CreateDirectory(Path.Combine(outputPath, folder.Name));

                // Recursively search for all items in the PST file
                foreach (object item in folder.Items)
                {
                    // If the item is an email message
                    if (item is MailItem mailItem)
                    {
                        // Iterate through all attachments in the email message
                        foreach (Attachment attachment in mailItem.Attachments)
                        {
                            // Save the attachment to disk
                            string saveFilePath = Path.Combine(outputPath, folder.Name, attachment.FileName);
                            attachment.SaveAsFile(saveFilePath);
                            Console.WriteLine("Saving " + saveFilePath);
                        }
                    }
                }
            }

            // Close the PST file
            outlookNamespace.RemoveStore(rootFolder);
            Marshal.ReleaseComObject(rootFolder);
            outlookNamespace.Logoff();

            Console.WriteLine("Attachments extracted successfully!");
            Console.WriteLine("Press enter to close");
            Console.ReadLine();
        }
    }
}