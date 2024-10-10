using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentAgent;
using System.Windows.Forms;


namespace ConsoleTest
{
    class Program
    {
        static void Main(string[] args)
        {
            DocumentAgent.ConfigureSettings FormSettings = new DocumentAgent.ConfigureSettings();
            FormSettings.ShowDialog();
        }
    }
}
