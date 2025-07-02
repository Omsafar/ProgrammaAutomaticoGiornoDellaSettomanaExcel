using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ClosedXML.Excel;
using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;

namespace ProgrammaAutomaticoGiornoDellaSettomanaExcel
{
    public partial class MainViewModel : ObservableObject
    {
        public ObservableCollection<string> Voci { get; } = new(SheetEntryStore.Load());

        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(AggiungiCommand))]
        private string? nuovaVoce;

        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(ElaboraCommand))]
        private string? filePath;

        // Proprietà selezione data (default domani)
        [ObservableProperty]
        private DateTime selectedDate = DateTime.Today.AddDays(1);

        // Proprietà status per il binding
        [ObservableProperty]
        private string status = string.Empty;

        /* ========================  Comandi  ======================== */

        [RelayCommand(CanExecute = nameof(CanAggiungi))]
        private void Aggiungi()
        {
            if (string.IsNullOrWhiteSpace(NuovaVoce)) return;
            var voce = NuovaVoce.Trim();
            if (!Voci.Any(v => string.Equals(v, voce, StringComparison.OrdinalIgnoreCase)))
                Voci.Add(voce);
            NuovaVoce = string.Empty;
        }
        private bool CanAggiungi() => !string.IsNullOrWhiteSpace(NuovaVoce);

        [RelayCommand]
        private void Rimuovi(string voce) => Voci.Remove(voce);

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
            Status = "Elaborazione...";

            if (!File.Exists(FilePath))
            {
                Status = "File non trovato";
                return;
            }

            try
            {
                using var wb = new XLWorkbook(FilePath);
                var data = SelectedDate;
                var giorno = CultureInfo.GetCultureInfo("it-IT")
                                        .DateTimeFormat
                                        .GetDayName(data.DayOfWeek);
                var header = $"Programma di {giorno} {data:dd/MM/yyyy}";

                foreach (var voce in Voci)
                {
                    var ws = wb.Worksheets
                                 .FirstOrDefault(s => string.Equals(s.Name, voce,
                                                      StringComparison.OrdinalIgnoreCase));
                    if (ws is not null)
                        ws.Cell("A1").Value = header;
                }
                wb.Save();
                Status = "Finito";
            }
            catch (Exception ex)
            {
                Status = $"Errore: {ex.Message}";
            }
        }
        private bool HasFile() => !string.IsNullOrWhiteSpace(FilePath);

        // Salvataggio voci alla chiusura
        public void Persisti() => SheetEntryStore.Save(Voci);
    }
}
