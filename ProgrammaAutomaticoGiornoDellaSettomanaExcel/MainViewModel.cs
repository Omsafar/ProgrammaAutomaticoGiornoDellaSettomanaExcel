using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ClosedXML.Excel;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;

namespace ProgrammaAutomaticoGiornoDellaSettomanaExcel;
    public partial class MainViewModel : ObservableObject
    {
        public ObservableCollection<string> Voci { get; } =
            new(SheetEntryStore.Load());

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(AggiungiCommand))]
    private string? nuovaVoce;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(ElaboraCommand))]
    private string? filePath;

    /* ========================  Comandi  ======================== */

    [RelayCommand(CanExecute = nameof(CanAggiungi))]
        private void Aggiungi()
        {
        if (string.IsNullOrWhiteSpace(NuovaVoce))
            return;

        var voce = NuovaVoce.Trim();
        bool giaPresente = Voci.Any(v =>
            string.Equals(v, voce, StringComparison.OrdinalIgnoreCase));
        if (!giaPresente)
            Voci.Add(voce);
        NuovaVoce = string.Empty;
    }
    private bool CanAggiungi() => !string.IsNullOrWhiteSpace(NuovaVoce);

        [RelayCommand]
        private void Rimuovi(string voce)
        {
            Voci.Remove(voce);
        }

        [RelayCommand]
        private void SfogliaExcel()
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Cartelle di lavoro Excel (*.xlsx)|*.xlsx",
                Title = "Seleziona il file Excel"
            };
            if (dlg.ShowDialog() == true)
                FilePath = dlg.FileName;
        }

        [RelayCommand(CanExecute = nameof(HasFile))]
        private void Elabora()
        {
            if (!File.Exists(FilePath)) return;

            using var wb = new XLWorkbook(FilePath);
            var oggi = DateTime.Today;
            var giorno = CultureInfo.GetCultureInfo("it-IT")
                                    .DateTimeFormat
                                    .GetDayName(oggi.DayOfWeek);
            var header = $"Programma di {giorno} {oggi:dd/MM/yyyy}";

            foreach (var voce in Voci)
            {
            var ws = wb.Worksheets
                         .FirstOrDefault(s => string.Equals(s.Name, voce,
                                                 StringComparison.OrdinalIgnoreCase));
            if (ws is not null)
                ws.Cell("A1").Value = header;
            }
            wb.Save(); // sovrascrive il file
        }
        private bool HasFile() => !string.IsNullOrWhiteSpace(FilePath);

        /* =======  Chiamato alla chiusura finestra (App.xaml.cs) ======= */
        public void Persisti() => SheetEntryStore.Save(Voci);
    }
