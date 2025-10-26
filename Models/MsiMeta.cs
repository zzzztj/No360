using System;

namespace No360.Models;

public sealed class MsiMeta
{
    public string? Manufacturer { get; init; }

    public string? ProductName { get; init; }

    public static MsiMeta? TryLoad(string msiPath)
    {
        try
        {
            var type = Type.GetTypeFromProgID("WindowsInstaller.Installer");
            if (type == null)
            {
                return null;
            }

            dynamic installer = Activator.CreateInstance(type)!;
            dynamic database = installer.OpenDatabase(msiPath, 0);
            const string query =
                "SELECT `Property`,`Value` FROM `Property` WHERE `Property` IN ('Manufacturer','ProductName')";
            dynamic view = database.OpenView(query);
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
    }
}
