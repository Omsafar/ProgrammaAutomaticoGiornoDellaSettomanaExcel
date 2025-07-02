using System.Windows;

namespace ProgrammaAutomaticoGiornoDellaSettomanaExcel
{
    public partial class App : Application
    {
        protected override void OnExit(ExitEventArgs e)
        {
            if (MainWindow?.DataContext is MainViewModel vm)
                vm.Persisti();
            base.OnExit(e);
        }
    }
}