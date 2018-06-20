using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft;
using System.Net;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Newtonsoft.Json.Linq;

namespace TestingDaFileExcel
{
    class Program
    {
        static string[] Chiamata(string ind, string cap, string citta)
        {
            string[] risp = new string[9];
            
            WebResponse response;
            Stream dataStream;
            WebRequest wR = WebRequest.Create(//API connection
                      "https://geocoder.cit.api.here.com/6.2/geocode.json?app_id=[KEY_HERE]&app_code=[KEY_HERE]&searchtext="
                      + ind //passaggio dei dati 
                      + "%20" + citta + "%20" + cap);
            Console.WriteLine("Richiesta con " + ind + " " + citta + " " + cap);
            response = wR.GetResponse();//Lettura risposta
            Console.WriteLine(response);
            dataStream = response.GetResponseStream(); //Creazion datastream
            StreamReader reader = new StreamReader(dataStream); //creazione di un lettore per lo stream
            JObject o = JObject.Parse(reader.ReadToEnd()); //Parsing del JSON in oggetto
            response.Close();

            //Console.WriteLine(o);
            try
            {
                risp[0] = "Città: " + (string)o["Response"]["View"][0]["Result"][0]["MatchQuality"]["City"];
                risp[0] = risp[0] + " Via: " + (string)o["Response"]["View"][0]["Result"][0]["MatchQuality"]["Street"][0];
                risp[0] = risp[0] + " Civico: " + (string)o["Response"]["View"][0]["Result"][0]["MatchQuality"]["HouseNumber"];
                risp[0] = risp[0] + " Cod Postale: " + (string)o["Response"]["View"][0]["Result"][0]["MatchQuality"]["PostalCode"];
                risp[1] = (string)o["Response"]["View"][0]["Result"][0]["Location"]["LocationId"];
                risp[2] = (string)o["Response"]["View"][0]["Result"][0]["Location"]["Address"]["Label"];
                risp[3] = (string)o["Response"]["View"][0]["Result"][0]["Location"]["Address"]["Street"];
                risp[4] = (string)o["Response"]["View"][0]["Result"][0]["Location"]["Address"]["HouseNumber"];
                risp[5] = (string)o["Response"]["View"][0]["Result"][0]["Location"]["Address"]["City"];
                risp[6] = (string)o["Response"]["View"][0]["Result"][0]["Location"]["Address"]["County"];
                risp[7] = (string)o["Response"]["View"][0]["Result"][0]["Location"]["Address"]["PostalCode"];
                risp[8] = (string)o["Response"]["View"][0]["Result"][0]["Location"]["Address"]["Country"];

            }
            catch (Exception ex)
            {
                Console.WriteLine("ERRORE");
                risp[1] = "Errore, non trovato";
            }
            return risp;
        }

        static void Main(string[] args)
        {
            Excel.Application xlApp = new Excel.Application()
            {
                Visible = false
            };
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\temp\IndirizziDaGeoLocalizzare.xlsx", false);

            // ******************************** CAMBIA QUI LE PAGINE ****************************
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[3]; //Pagina Excel da leggere


            Excel.Range xlRange = xlWorksheet.UsedRange;
            // ******************************* QUI FINO A CHE RIGA DEVE LEGGERE *****************
            int rowCount = 211; //Righe da leggere
            int colCount = 2;
            string[] risposta = new string[9];
            string indirizzo = "";
            string cap = "";
            string citta = "";
            int j = 1;
            int ct = 0;

            for (int i = 2; i <= rowCount; i++)//i deve sempre partire minimo da 1
            {
                j = 2;
                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null) {
                    Console.WriteLine("Leggo da excel riga" + i + " colonna" + j);
                    Console.Write("Che vale:" + xlRange.Cells[i, j].Value2.ToString() + "\n");
                    indirizzo = xlRange.Cells[i, j].Value2.ToString();
                }
                j = 6;
                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                {
                    Console.WriteLine("Leggo da excel riga" + i + " colonna" + j);
                    Console.Write("Che vale:" + xlRange.Cells[i, j].Value2.ToString() + "\n");
                    cap = xlRange.Cells[i, j].Value2.ToString();
                }
                j = 7;
                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                {
                    Console.WriteLine("Leggo da excel riga" + i + " colonna" + j);
                    Console.Write("Che vale:" + xlRange.Cells[i, j].Value2.ToString() + "\n");
                    citta = xlRange.Cells[i, j].Value2.ToString();
                }
                risposta = Chiamata(indirizzo, cap, citta); //richiamo la funzione che eseguirà la chiamata
                
                //scrivo nelle celle excel
                xlRange.Cells[i, 11].Value = risposta[2]; //risultato Here completo
                xlRange.Cells[i, 12].Value = risposta[1]; //location ID
                xlRange.Cells[i, 13].Value = risposta[0]; //match quality
                xlRange.Cells[i, 14].Value = risposta[3]; //via
                xlRange.Cells[i, 15].Value = risposta[4]; //civico
                xlRange.Cells[i, 16].Value = risposta[5]; //città
                xlRange.Cells[i, 17].Value = risposta[7]; //CAP
                xlRange.Cells[i, 18].Value = risposta[6]; //provincia
                xlRange.Cells[i, 19].Value = risposta[8]; //stato

                xlWorkbook.Save(); //salvo le modifiche sul file excel
            }

            xlWorkbook.Close(); //chiudo il wb
            xlApp.Quit(); //chiudo excel
            Console.ReadLine();
        }
    }
}
