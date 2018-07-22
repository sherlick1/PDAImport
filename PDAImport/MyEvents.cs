using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDAImport
{
    public delegate void MyHandler1(object sender, MyEvent e);

    public class MyEvent : EventArgs
    {
        public string message;
    }
}
