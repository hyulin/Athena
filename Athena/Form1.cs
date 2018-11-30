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
    public partial class Form1 : Form
    {
        CBotClient BotClient = new CBotClient();

        public Form1()
        {
            BotClient.InitBotClient();

            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            BotClient.telegramAPIAsync();

            BotClient.setTelegramEvent();
        }
    }
}
