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

            Access.Application access = null;
            worker.ReportProgress(0, "Démarrage");

            try
            {
                access = new Access.Application();
                RapporterProgression("Access démarré");

                access.OpenCurrentDatabase(URIAttendu, false, null);
                RapporterProgression("Bd ouverte");
                RapporterProgression("Lecture des tables : " + access.CurrentData.AllTables.Count);
            }
            catch (Exception exp)
            {
                RapporterProgression(exp.Message);
            }
            finally
            {
                if (access != null)
                {
                    access.Quit(Access.AcQuitOption.acQuitSaveNone);
                    Marshal.ReleaseComObject(access);
                    access = null;
                    RapporterProgression("Access fermé");
                }
            }
            return 0;
        }
    }
}
