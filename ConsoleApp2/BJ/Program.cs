using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BJ
{
    class Program
    {
        static void Main(string[] args)
        {
            string ipaddr = "127.0.0.1";
            int prot = 6000;            
            int iSoket = 0;          
            try
            {
                iSoket = CServer.InitServer(ipaddr, prot);
                Console.WriteLine("iSoket=" + iSoket);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);

                throw;
            }
        }
    }
}
