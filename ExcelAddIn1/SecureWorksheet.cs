using System;
using System.Windows.Forms;
using ExcelAddIn1.PricerObjects;

namespace ExcelAddIn1
{
    public static class SecureWorksheet
    {
        public static bool SecuriseNewWs(Parameters details)
        {
            //on securuse la WS en cas d'erreur de l'user
            try
            {
                if (details.newWorksheet == null) throw new Exception("WORKSHEET");
                if (details.newWorksheet.Range["B1"].Value == null) throw new Exception("TICKER");
                if (details.newWorksheet.Range["B3"].Value == null) throw new Exception("OPTION");
            }
            catch (Exception exception)
            {
                switch (exception.Message)
                {
                    case "WORKSHEET":
                        MessageBox.Show("Merci de créer une nouvelle feuille.");
                        Globals.ThisAddIn.Application.ScreenUpdating = true;
                        return true;
                    case "TICKER":
                        MessageBox.Show("Vous devez saisir un ticker.");
                        Globals.ThisAddIn.Application.ScreenUpdating = true;
                        return true;
                    case "OPTION":
                        MessageBox.Show("Vous devez saisir le type d'Option(Call ou Put).");
                        Globals.ThisAddIn.Application.ScreenUpdating = true;
                        return true;
                    default:
                        MessageBox.Show("Il y a une erreur");
                        return true;
                }
            }
            return false;
        }
    }
}