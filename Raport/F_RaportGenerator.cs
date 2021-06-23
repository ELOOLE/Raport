﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Raport
{
    public partial class F_RaportGenerator : Form
    {
        private string filePath, fileContent;
        

        public F_RaportGenerator()
        {
            InitializeComponent();
        }

        private void portycsvToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            string host, port, port2, proto;
            string[] array_linijka;
            List<Host> list_wynik_host = new List<Host>();
            
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //Get the path of specified file
                filePath = openFileDialog1.FileName;

                //Read the contents of the file into a stream
                var fileStream = openFileDialog1.OpenFile();
                string linijka;
                System.IO.StreamReader file = new System.IO.StreamReader(filePath);
                //using (StreamReader reader = new StreamReader(fileStream))
                //{
                //fileContent = reader.ReadToEnd();
                while ((linijka = file.ReadLine()) != null)
                {
                    //MessageBox.Show(linijka);
                    array_linijka = linijka.Split(',');
                    host = array_linijka[0].Trim(' ', '"');
                    port = array_linijka[1].Trim(' ', '"');
                    proto = array_linijka[2].Trim(' ', '"');
                    //linijka = linijka.Trim(new Char[] { ' ', '"', '.' });
                    //fileContent += linijka.ToString();
                    //MessageBox.Show($"host: {host}, port: {port}, proto: {proto}");
                    
                    if(list_wynik_host.Exists(x => x.HostIP.Equals(host)))
                    {
                        //MessageBox.Show("istnieje");
                        if(proto.ToLower() == "tcp")
                        {
                            if(!string.IsNullOrEmpty(list_wynik_host.Find(x => x.HostIP == host).PortTCP.ToString()))
                            {
                                port = list_wynik_host.Find(x => x.HostIP == host).PortTCP.ToString() + ", " + port;
                            }
                                                        
                            port2 = list_wynik_host.Find(x => x.HostIP == host).PortUDP.ToString();

                            list_wynik_host.RemoveAll(x => x.HostIP == host);
                            list_wynik_host.Add(new Host() { HostIP = host, PortTCP =  port, PortUDP = port2 });
                        }
                        else
                        {
                            if(!string.IsNullOrEmpty(list_wynik_host.Find(x => x.HostIP == host).PortUDP.ToString()))
                            {
                                port = list_wynik_host.Find(x => x.HostIP == host).PortUDP.ToString() + ", " + port;
                            }
                                                        
                            port2 = list_wynik_host.Find(x => x.HostIP == host).PortTCP.ToString();
                            
                            list_wynik_host.RemoveAll(x => x.HostIP == host);
                            list_wynik_host.Add(new Host() { HostIP = host, PortTCP = port2, PortUDP = port });
                        }
                    }
                    else
                    {
                        if (proto.ToLower() == "tcp")
                        {
                            list_wynik_host.Add(new Host() { HostIP = host, PortTCP = port, PortUDP = "" });
                        }
                        else if(proto.ToLower() == "udp")
                        {
                            list_wynik_host.Add(new Host() { HostIP = host, PortTCP = "", PortUDP = port });
                        }
                        else
                        {
                            list_wynik_host.Add(new Host() { HostIP = host, PortTCP = "tcp", PortUDP = "udp" });
                        }
                    }
                }
                //}
                fileContent = "";

                string porttcp, portudp;
                foreach (Host host1 in list_wynik_host)
                {
                    host = host1.HostIP.ToString().Trim();
                    porttcp = host1.PortTCP.ToString().Trim();
                    portudp = host1.PortUDP.ToString().Trim();

                    fileContent += host + ";" + porttcp + ";" + portudp;
                    fileContent += System.Environment.NewLine;
                }

                //MessageBox.Show(fileContent);
                Clipboard.Clear();
                Clipboard.SetText(fileContent);

                MessageBox.Show("Zakończono parsowanie pliku. Wynik w schowku.");
                textBox1.Text = fileContent;
                //fileContent = fileContent.Trim(' ', '"');                
            }
        }

        public class Host : IEquatable<Host>
        {
            public string HostIP { get; set; }
            public string PortTCP { get; set; }
            public string PortUDP { get; set; }

            public bool Equals(Host other)
            {
                throw new NotImplementedException();
            }
        }

        private void zrodlaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Plik zródłowy Metasploit > Porty (csv) \n- polecenie w konsoli msfconsole: \n[services -u -c name,proto -o /home/user/export_wynik.csv] \n- wynik działania programu zapisuje się w pamięci podręcznej (schowek).");
        }

        private void WalidujMetasploitPortyCSV()
        {
            //services -u -c name,proto -o /home/user/export_wynik.csv

        }

    }
}
