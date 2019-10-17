using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
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
        private JObject _config = null;

        public MainWindow()
        {
            InitializeComponent();
            ChargerConfig();
        }

        private void ChargerConfig()
        {
            string cheminConfig = CorrecteurComparateur.Properties.Resources.ConfigFile;
            if (!File.Exists(cheminConfig))
            {
                return;
            }
            this._config = JObject.Parse(File.ReadAllText(cheminConfig));

            dynamic configAccess = _config["ComparateurAccess"];
            Comparateurs.ComparateurAccess comparateur =
                new Comparateurs.ComparateurAccess()
                {
                    URIAttendu = configAccess.URIAttendu
                };
            comparateur.Comparer(null);
        }
    }
}
