using System;
using System.Runtime.InteropServices;

namespace No360.Models;

public sealed class MsiMeta
{
    public string? Manufacturer { get; init; }

    public string? ProductName { get; init; }

    public static MsiMeta? TryLoad(string msiPath)
    {
        dynamic? installer = null;
        dynamic? database = null;
        dynamic? view = null;
        try
        {
            var type = Type.GetTypeFromProgID("WindowsInstaller.Installer");
            if (type == null)
            {
                return null;
            }

            installer = Activator.CreateInstance(type)!;
            database = installer.OpenDatabase(msiPath, 0);
            const string query =
                "SELECT `Property`,`Value` FROM `Property` WHERE `Property` = 'Manufacturer' OR `Property` = 'ProductName'";
            view = database.OpenView(query);
            view.Execute(null);

            string? manufacturer = null;
            string? product = null;
            while (true)
            {
                dynamic record = view.Fetch();
                if (record == null)
                {
                    break;
                }

                try
                {
                    string property = record.StringData(1);
                    string value = record.StringData(2);
                    if (property == "Manufacturer")
                    {
                        manufacturer = value;
                    }
                    else if (property == "ProductName")
                    {
                        product = value;
                    }
                }
                finally
                {
                    ReleaseComObject(record);
                }
            }

            return new MsiMeta
            {
                Manufacturer = manufacturer,
                ProductName = product
            };
        }
        catch
        {
            return null;
        }
        finally
        {
            try
            {
                view?.Close();
            }
            catch
            {
                // A partially initialized view may not be open yet.
            }
            finally
            {
                // These COM objects are local to this inspection. Release them
                // in reverse order, even when opening or querying the MSI fails.
                ReleaseComObject(view);
                ReleaseComObject(database);
                ReleaseComObject(installer);
            }
        }
    }

    private static void ReleaseComObject(object? value)
    {
        if (value != null && Marshal.IsComObject(value))
        {
            Marshal.ReleaseComObject(value);
        }
    }
}
