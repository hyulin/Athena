using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Athena
{
    public partial class Athena : Form
    {
        CBotClient BotClient = new CBotClient();

        public Athena()
        {
            BotClient.InitBotClient();

            InitializeComponent();
        }

        private void Athena_Load(object sender, EventArgs e)
        {
            BotClient.telegramAPIAsync();

            BotClient.setTelegramEvent();
        }
    }
}
