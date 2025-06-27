using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ClosedXML.Excel;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;

namespace ProgrammaAutomaticoGiornoDellaSettomanaExcel;
    public partial class MainViewModel : ObservableObject
    {
        public ObservableCollection<string> Voci { get; } =
            new(SheetEntryStore.Load());

        [ObservableProperty] private string? nuovaVoce;
        [ObservableProperty] private string? filePath;

        /* ========================  Comandi  ======================== */

        [RelayCommand(CanExecute = nameof(CanAggiungi))]
        private void Aggiungi()
        {
            if (string.IsNullOrWhiteSpace(NuovaVoce)) return;
            if (!Voci.Contains(NuovaVoce))
                Voci.Add(NuovaVoce);
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
                if (wb.Worksheets.TryGetWorksheet(voce, out var ws))
                    ws.Cell("A1").Value = header;
            }
            wb.Save(); // sovrascrive il file
        }
        private bool HasFile() => !string.IsNullOrWhiteSpace(FilePath);

        /* =======  Chiamato alla chiusura finestra (App.xaml.cs) ======= */
        public void Persisti() => SheetEntryStore.Save(Voci);
    }
