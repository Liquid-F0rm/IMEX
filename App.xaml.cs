using System;
using System.Configuration;
using System.Data;
using System.Windows;
using System.Globalization;
using System.Windows.Data;

namespace WpfImageToText
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : System.Windows.Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            
            // Initialiser les API Windows modernes si nécessaire
            // Cette étape n'est pas toujours obligatoire mais peut aider à éviter certains problèmes
            try
            {
                // Pas d'initialisation spécifique nécessaire pour une application WPF .NET 8
                // Les API Windows Runtime ne sont pas directement accessibles sans packages supplémentaires
            }
            catch (Exception ex)
            {
                // Les API peuvent ne pas être disponibles sur tous les systèmes
                // Gérez cela gracieusement
                Console.WriteLine($"Initialisation Windows API échouée: {ex.Message}");
            }
        }
    }
}
