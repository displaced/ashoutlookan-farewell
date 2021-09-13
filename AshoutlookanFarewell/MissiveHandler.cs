using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;

namespace AshoutlookanFarewell
{

    /// <summary>
    /// Co-ordinates the whole plugin.  Responds to and triggers Outlook events to make the silliness happen.
    /// </summary>
    public static class MissiveHandler
    {

        /// <summary>
        /// Start composing a new Missive.
        /// </summary>
        public static void NewMissive()
        {
            // Set phasers to Longing
            MusicController.StartPlayback();

            // Create a new email
            Outlook.MailItem mailItem = (Outlook.MailItem)
            AddIn.app.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.Subject = "";
            mailItem.To = "";
            mailItem.Body = "";

            // Set a custom MAPI/SMTP Envelope header so that the plugin running on the recipient's machine can detect a missive.
            string header = "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/X-AshoutlookanFarewell";
            mailItem.PropertyAccessor.SetProperty(header, true);

            // Set the mail's default text
            mailItem.HTMLBody = $"<FONT face=\"{Settings.MissiveFontFace}\" size=+1>" +
                $"<p>{Settings.GetRandomFragment(Fragments.Salutation)}</p><p></p><p></p><p>" +
                $"{Settings.GetRandomFragment(Fragments.Goodbye)}<br>{Settings.GetRandomFragment(Fragments.NomDePlume)}</p>";
            mailItem.Importance = Outlook.OlImportance.olImportanceLow;
            mailItem.Display(false);

            // Hook the close event so we know when to stop the music
            var inspector = mailItem.GetInspector;
            ((InspectorEvents_10_Event)inspector).Close += MissiveClosedHandler;
        }

        /// <summary>
        /// A Missive has closed - either a draft has been saved, the missive has been sent, or the user has stopped reading a missive.
        /// </summary>
        private static void MissiveClosedHandler()
        {
            MusicController.StopPlayback();
        }
    }

    
}
