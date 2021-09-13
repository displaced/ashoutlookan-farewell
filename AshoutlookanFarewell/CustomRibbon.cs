using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AshoutlookanFarewell
{
    public partial class CustomRibbon
    {
        private void CustomRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }


        private void sbtnNewMissive_Click(object sender, RibbonControlEventArgs e)
        {
            MissiveHandler.NewMissive();
        }
    }
}
