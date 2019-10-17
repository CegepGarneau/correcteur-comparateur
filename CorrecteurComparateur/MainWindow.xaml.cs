using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
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
        private JObject          _config = null;
        private BackgroundWorker _worker = null;

        public MainWindow()
        {
            InitializeComponent();
            ChargerConfig();
            InitialiserWorker();
        }

        /// <summary>
        /// S'occupe de charger le fichier de config
        /// </summary>
        private void ChargerConfig()
        {
            string cheminConfig = CorrecteurComparateur.Properties.Resources.ConfigFile;
            if (!File.Exists(cheminConfig))
            {
                return;
            }
            this._config = JObject.Parse(File.ReadAllText(cheminConfig));
        }

        /// <summary>
        /// S'occupe d'initialiser la tâche en arrière-plan
        /// </summary>
        /// <see cref="https://docs.microsoft.com/en-us/dotnet/framework/winforms/controls/how-to-implement-a-form-that-uses-a-background-operation"/>
        private void InitialiserWorker()
        {
            _worker = new BackgroundWorker();

            _worker.WorkerReportsProgress      = true;
            _worker.WorkerSupportsCancellation = true;

            _worker.DoWork             += new DoWorkEventHandler(DemarrerTraitement);
            _worker.ProgressChanged    += new ProgressChangedEventHandler(GererProgression);
            _worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(GererCompletion);
            
        }

        /// <summary>
        /// Pour afficher un message durant le traitement
        /// </summary>
        private void Log(string message)
        {
            this._Log.AppendText(message);
            this._Log.AppendText(Environment.NewLine);
        }

        /// <summary>
        /// Pour démarrer le traitement en arrière-plan
        /// </summary>
        private void DemarrerTraitement(object sender, DoWorkEventArgs e)
        {
            // Get the BackgroundWorker that raised this event.
            BackgroundWorker worker = sender as BackgroundWorker;

            dynamic configAccess = _config["ComparateurAccess"];
            Comparateurs.ComparateurAccess comparateur =
                new Comparateurs.ComparateurAccess()
                {
                    URIAttendu   = configAccess.URIAttendu,
                    URIAComparer = configAccess.URIAComparer
                };
            e.Result = comparateur.Comparer(worker, e);
        }

        /// <summary>
        /// Est appelée durant la progression de la tâche
        /// </summary>
        private void GererProgression(object sender, ProgressChangedEventArgs e)
        {
            this._Progression.Value = e.ProgressPercentage;
            this.Log(e.UserState.ToString());
        }

        /// <summary>
        /// Est appelée lors de la complétion de la tâche
        /// </summary>
        private void GererCompletion(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                this.Log(e.Error.Message);
            }
            else if (e.Cancelled)
            {
                this.Log("Annulé");
            }
            else
            {
                this.Log("Résultat : " + e.Result.ToString());
            }

            // Mettre le UI à jour
            this._btnComparer.IsEnabled = true;
        }

        /// <summary>
        /// Pour gérer le clic du bouton
        /// </summary>
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this._btnComparer.IsEnabled = false;
            this._worker.RunWorkerAsync();
        }
    }
}
