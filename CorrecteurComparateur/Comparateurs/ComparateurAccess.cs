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

namespace CorrecteurComparateur.Comparateurs
{
    public abstract class AbsComparaison
    {
        public Object Attendu { get; set; }
        public Object AComparer { get; set; }
        public decimal NbPoints { get; set; }
        public String Titre { get; set; }

        public abstract decimal Comparer();
    }

    public interface ILoggueur
    {
        void Log(string message);
    }

    public class ComparateurAccess
    {
        public static string GetNom()
        {
            return "Comparateur Access";
        }

        public string URIAttendu { get; set; }

        public string URIAComparer { get; set; }

        public decimal Comparer(ILoggueur loggueur)
        {
            Access.Application access = null;
            try
            {
                access = new Access.Application();
                access.OpenCurrentDatabase(URIAttendu, false, null);
                Console.WriteLine(access.CurrentData.AllTables.Count);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (access != null)
                {
                    access.Quit(Access.AcQuitOption.acQuitSaveNone);
                    Marshal.ReleaseComObject(access);
                    access = null;
                }
            }
            return 0;
        }
    }
}
