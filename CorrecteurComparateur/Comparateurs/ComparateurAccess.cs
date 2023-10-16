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


            OleDbConnection connAttendue = new OleDbConnection();
            connAttendue.ConnectionString = $"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={URIAttendu};Persist Security Info=True";

            RapporterProgression("Démarrage");

            try
            {
                connAttendue.Open();
                RapporterProgression("Bds ouvertes");

                ComparerTables(connAttendue, connAttendue);
            }
            catch (Exception exp)
            {
                RapporterProgression(exp.Message);
            }
            finally
            {
                connAttendue.Close();
                RapporterProgression("Bd attendue fermée");
                /*
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
                */
            }
            return 0;
        }

        /// <summary>
        /// Pour comparer toutes les tables
        /// </summary>
        private void ComparerTables(OleDbConnection connAttendue, OleDbConnection conAComparer)
        {
            RapporterProgression("Lecture des tables");

            string[] restrictions = new string[4];
            restrictions[3] = "Table";

            // Get list of user tables
            DataTable tablesAttendues = connAttendue.GetSchema("Tables", restrictions);

            RapporterProgression(tablesAttendues.Rows.Count + " tables à comparer");
            _progressionTotale = tablesAttendues.Rows.Count * FACTEUR_TABLE;

            foreach (DataRow dtRow in tablesAttendues.Rows)
            {
                ComparerTables(connAttendue, conAComparer, dtRow["TABLE_NAME"].ToString());
                _progression += FACTEUR_TABLE;
            }

            /*
            _progressionTotale = connAttendue.CurrentData.AllTables.Count;
            List<Access.AccessObject> tables = new List<Access.AccessObject>();

            // Trouver les vraies tables et mettre de côté les tables internes d'Access (commançant par "MS")
            foreach (Access.AccessObject table in connAttendue.CurrentData.AllTables)
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
                Access.AccessObject tableAComparer = conAComparer.CurrentData.AllTables[tableAttendue.Name];
                ComparerTables(tableAttendue, tableAComparer);
                _progression += FACTEUR_TABLE;
            }
            */
        }

        /// <summary>
        /// Pour comparer deux tables
        /// </summary>
        private void ComparerTables(OleDbConnection connAttendue, OleDbConnection conAComparer, string nomTable)
        {
            RapporterProgression("Table " + nomTable + " à comparer");
            string[] restrictions = new string[4];
            restrictions[2] = nomTable;

            DataTable tableAttendue = connAttendue.GetSchema("Columns", restrictions);

            foreach (DataRow row in tableAttendue.Rows)
            {
                foreach (DataColumn col in tableAttendue.Columns)
                {
                    if (!String.IsNullOrWhiteSpace(row[col].ToString()))
                    {
                        RapporterProgression("\t" + col.ColumnName + " : " + row[col].ToString());
                    }
                }
            }


                //RapporterProgression("Table " + nomTable + " trouvée : " + tableAComparer.ToString());
        }
    }
}
