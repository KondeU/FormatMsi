using System;
using System.IO;
using WindowsInstaller;

namespace FormatMsi
{
    class Program
    {
        static void Main(string[] args)
        {
            String inputFile = null;

            if (args.Length == 1)
            {
                inputFile = args[0];
            }
            else
            {
                Console.WriteLine("Please enter the msi file:");
                inputFile = Console.ReadLine();
            }

            String productName = null;
            String productVersion = null;
            String formatMsiFileName = null;
            try
            {
                if (inputFile.EndsWith(".msi", StringComparison.OrdinalIgnoreCase))
                {
                    productName = GetPropertyFromMsi(inputFile, "ProductName");
                    productVersion = GetPropertyFromMsi(inputFile, "ProductVersion");
                }
                else
                {
                    Console.WriteLine("Error: Invalid input file!");
                    return;
                }

                formatMsiFileName = String.Format("{0}_{1}.msi", productName, productVersion);
                Console.WriteLine("Format msi file name: " + formatMsiFileName);

                File.Copy(inputFile, formatMsiFileName);
                File.Delete(inputFile);
            }
            catch (Exception exception)
            {
                Console.WriteLine("Exception: " + exception.Message);
            }
        }

        static String GetPropertyFromMsi(String msi, String property)
        {
            String ret = null;

            // WindowsInstaller from [SYSTEM]:\Windows\System32\msi.dll
            Type classType = Type.GetTypeFromProgID("WindowsInstaller.Installer");
            Installer installer = Activator.CreateInstance(classType) as Installer;

            // Open the msi file for reading, 0 means read, 1 means read and write
            Database database = installer.OpenDatabase(msi, 0);

            // The requested property fetching command
            String sql = String.Format("SELECT Value FROM Property WHERE Property='{0}'", property);

            // Open the database view and then execute SQL command
            View view = database.OpenView(sql);
            view.Execute(null);

            // Read from the fetched record
            Record record = view.Fetch();
            if (record != null)
            {
                ret = record.get_StringData(1);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(record);
            }

            // Close the database view
            view.Close();

            // Release the view's and the database's COM object
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(view);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(database);

            return ret;
        }
    }
}
