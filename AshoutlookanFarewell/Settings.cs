using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AshoutlookanFarewell
{
    public enum Fragments
    {
        Salutation,
        Goodbye,
        NomDePlume
    }

    public static class Settings
    {
        public static string DataDir = "";
        public static string MediaUrl = "https://www.youtube.com/watch?v=uZmxZThb084";
        public static List<string> Salutations = new List<string>();
        public static List<string> Goodbyes = new List<string>();
        public static List<string> NomsDePlume = new List<string>();
        public static string MissiveFontFace = "";
        public static bool RandomisePlayhead = true;
        public static bool FadeIn = true;

        static Random rnd = new Random();

        public static void LoadDefaults()
        {
            // Defaut list of opening lines for the missive
            Salutations.Add("Dearest ");

            // Default list of heavy-hearted goodbyes
            Goodbyes.AddRange(new string[] { "With all my heart,", "'Till we meet once more,", "Yours Forever,", "I remain your humble servant," });

            // Default name of the sender (get the user's first name according to their Exchange profile)
            NameSpace ns = AddIn.app.Application.GetNamespace("MAPI");
            var exchangeUser = ns.CurrentUser.AddressEntry.GetExchangeUser();

            NomsDePlume.Add(exchangeUser.FirstName);

            // Set the default font to a nice cursive one
            MissiveFontFace = "Blackadder ITC";
        }

        
        public static string GetRandomFragment(Fragments fragmentType)
        {
            List<string> list = new List<string>();

            switch (fragmentType)
            {
                case Fragments.Goodbye:
                    list = Goodbyes;
                    break;
                case Fragments.NomDePlume:
                    list = NomsDePlume;
                    break;
                case Fragments.Salutation:
                    list = Salutations;
                    break;
            }

            int r = rnd.Next(list.Count);
            return list[r];
        }

    }
}
