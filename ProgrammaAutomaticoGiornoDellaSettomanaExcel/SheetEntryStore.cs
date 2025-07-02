using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using System;

public static class SheetEntryStore
{
    private static readonly string FilePath =
        Path.Combine(
            AppContext.BaseDirectory,
            "voci.json");

    public static List<string> Load()
    {
        try
        {
            if (File.Exists(FilePath))
                return JsonConvert.DeserializeObject<List<string>>(File.ReadAllText(FilePath)) ?? [];
        }
        catch { /* log se vuoi */ }
        return []; // elenco vuoto al primo avvio
    }

    public static void Save(IEnumerable<string> voci)
    {
        File.WriteAllText(FilePath,
            JsonConvert.SerializeObject(voci, Formatting.Indented));
    }
}
