using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CorrecteurComparateur
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Object _config = null;

        public MainWindow()
        {
            InitializeComponent();
            ChargerConfig();
        }

        private void ChargerConfig()
        {
            string fichier = CorrecteurComparateur.Properties.Resources.ConfigFile;
            if (!File.Exists(fichier))
            {
                return;
            }
            this._config = JsonConvert.DeserializeObject(File.ReadAllText(fichier));

            Comparateurs.ComparateurAccess comparateur = 
                new Comparateurs.ComparateurAccess() 
                { 
                    URIAttendu = "C:\\Users\\nrichard\\Google Drive\\Cours\\2019-A\\420-533\\TP\\TP 1\\TP 1 - Requetes.accdb" 
                };
            comparateur.Comparer(null);
        }
    }
}
