using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Reflection;
using Access = Microsoft.Office.Interop.Access;
using System.Windows;
using System.Threading.Tasks;
using System.ComponentModel;

namespace CorrecteurComparateur.Comparateurs
{
    /// <summary>
    /// Permet de comparer deux bases de données Access
    /// </summary>
    public class ComparateurAccess
    {
        private const int FACTEUR_TABLE = 100;

        public string URIAttendu { get; set; }
        public string URIAComparer { get; set; }

        private int _progression       = 0;
        private int _progressionTotale = 100;
        private BackgroundWorker _worker = null;

        /// <summary>
        /// Pour rapporter la progression
        /// </summary>
        private void RapporterProgression(string message)
        {
            _worker.ReportProgress(_progression * 100 / _progressionTotale, message);
        }

        /// <summary>
        /// Compare et retourne une note pour la comparaison de URIAttendu avec URIAComparer
        /// </summary>
        /// <param name="worker">Tâche en arrière plan</param>
        /// <param name="e">Événement pour interagir avec la tâche</param>
        /// <returns>Une note</returns>
        public decimal Comparer(BackgroundWorker worker, DoWorkEventArgs e)
        {
            this._worker = worker;

            Access.Application appAttendue  = null;
            Access.Application appAComparer = null;

            RapporterProgression("Démarrage");

            try
            {
                appAttendue = OuvrirBd(URIAttendu);
                appAComparer = OuvrirBd(URIAComparer);
                RapporterProgression("Bds ouvertes");

                ComparerTables(appAttendue, appAComparer);
            }
            catch (Exception exp)
            {
                RapporterProgression(exp.Message);
            }
            finally
            {
                if (appAttendue != null)
                {
                    appAttendue.Quit(Access.AcQuitOption.acQuitSaveNone);
                    Marshal.ReleaseComObject(appAttendue);
                    appAttendue = null;
                    RapporterProgression("Bd attendue fermée");
                }
                if (appAComparer != null)
                {
                    appAComparer.Quit(Access.AcQuitOption.acQuitSaveNone);
                    Marshal.ReleaseComObject(appAComparer);
                    appAComparer = null;
                    RapporterProgression("Bd à comparer fermée");
                }
            }
            return 0;
        }

        /// <summary>
        /// Pour ouvrir une bd
        /// </summary>
        /// <param name="uri">Chemin pour la bd</param>
        /// <returns>Access ouvert pour la bd demandée</returns>
        private Access.Application OuvrirBd(string uri)
        {
            RapporterProgression("Ouverture de " + uri);

            Access.Application app = null;
            app = new Access.Application();
            
            app.OpenCurrentDatabase(uri, false, null);

            return app;
        }

        /// <summary>
        /// Pour comparer toutes les tables
        /// </summary>
        private void ComparerTables(Access.Application appAttendue, Access.Application appAComparer)
        {
            RapporterProgression("Lecture des tables");
            _progressionTotale = appAttendue.CurrentData.AllTables.Count;
            List<Access.AccessObject> tables = new List<Access.AccessObject>();

            // Trouver les vraies tables et mettre de côté les tables internes d'Access (commançant par "MS")
            foreach (Access.AccessObject table in appAttendue.CurrentData.AllTables)
            {
                if (!table.Name.StartsWith("MS"))
                {
                    tables.Add(table);
                    RapporterProgression("Table : " + table.Name + " ");
                }
            }
            RapporterProgression(tables.Count + " tables à comparer");
            _progressionTotale = tables.Count * FACTEUR_TABLE;

            foreach (Access.AccessObject tableAttendue in tables)
            {
                Access.AccessObject tableAComparer = appAComparer.CurrentData.AllTables[tableAttendue.Name];
                ComparerTables(tableAttendue, tableAComparer);
                _progression += FACTEUR_TABLE;
            }
        }

        /// <summary>
        /// Pour comparer deux tables
        /// </summary>
        private void ComparerTables(Access.AccessObject tableAttendue, Access.AccessObject tableAComparer)
        {
            RapporterProgression("Table " + tableAttendue.Name + " trouvée : " + (tableAComparer != null));

        }
    }
}
