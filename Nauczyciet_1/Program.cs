using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace Nauczyciel1
{
    class WczytajExcel
    {
        private string m_xlFileName = @"E:\plik_in.xls";
        private Excel.Application m_xlApp;                      // obiekt aplikacji
        private Excel.Range m_projectRange;                     // zakres danych do wczytania
        private Excel.Workbook m_xlWorkbook;                    // dokument
        private Excel.Worksheet m_xlWorksheet;                  // arkusz
        private System.Object m_xx = System.Type.Missing;

        public string[,] CzytajDane()
        {
            m_xlApp = new Excel.Application();
            m_xlApp.DisplayAlerts = false;

            m_xlWorkbook = m_xlApp.Workbooks.Open(m_xlFileName,
                m_xx, m_xx, m_xx, m_xx, m_xx, m_xx, m_xx,
                m_xx, m_xx, m_xx, m_xx, m_xx, m_xx, m_xx);

            m_xlWorksheet = (Excel.Worksheet)m_xlWorkbook.Worksheets[1];   // 0 wskazuje na pierwszy arkusz

            string startCell = "A1";  // zakres danych do wczytania
            string endCell = "F6";
            m_projectRange = m_xlWorksheet.get_Range(startCell, endCell);

            Array projectCells = (Array)m_projectRange.Cells.Value2;

            int col = m_projectRange.Columns.Count;
            int row = m_projectRange.Rows.Count;
            string[,] tab1 = new string[col, row];
            for (int i = 0; i < col; i++)
            {
                for (int j = 0; j < row; j++)
                {
                    tab1[i, j] = " " + projectCells.GetValue(i + 1, j + 1);
                }
            }
            m_xlApp.Quit();

            Console.Write("Wczytana tablica z pliku Excela   BIJACZ \n");

            for (int i = 0; i < col; i++)
            {
                for (int j = 0; j < row; j++)
                {
                    Console.Write(tab1[i, j] + "\t");
                }
                Console.WriteLine();
            }
           
            return tab1;
        }
    }

    class WczytajTxt
    {
        public void CzytajDaneTabTxt()
        {
            Console.Write("\n" + "Podaj nazwe pliku: ");
            string szSrcFile = Console.ReadLine();

            // czytanie zawartości pliku wejściowego
            string szSrcLine;
            ArrayList szContents = new ArrayList();
            FileStream fsInput = new FileStream(szSrcFile, FileMode.Open, FileAccess.Read);
            StreamReader srInput = new StreamReader(fsInput);
            while ((szSrcLine = srInput.ReadLine()) != null)
            {
                // dołączanie do tablicy
                szContents.Add(szSrcLine);
            }
            srInput.Close();
            fsInput.Close();

            Console.Write("\n" + "Plik wczytaniu do tabeli." + "\n");

            //wyswietlenie zmian II
            Console.Write("\n" + "Zmiany po wczytaniu pliku." + "\n");
            Console.WriteLine(szContents[0]);
            Console.WriteLine(szContents[1]);
            Console.WriteLine(szContents[2]);
        }    
    }

    /*
    class Pytania
    {
        public void Odpowiedz(tab[i,j],k)
        {
            for (int i = 0; i < k; i++)
            {

                Console.Write("Przetlumacz słowo: " + tab[i, 0] + "\n" + "Odpowiedz: ");
                string slowko_podane = Console.ReadLine();
                if (tab[i, 1] == slowko_podane)
                {
                    Console.Write("Poprawnie: " + tab[i, 1] + " = " + tab[i, 0] + "\n");
                    tab[i, 2] = "1";       // poprawnie
                    tab[i, 3] = "1";       // nie powtarzaj
                }
                else
                {
                    tab[i, 2] = "2";      // nie poprawnie
                    tab[i, 3] = "2";      // powtarzaj
                    Console.Write("Nie Poprawnie\n" + "Prawidlowa odopwiedz: " + tab[i, 1] + " = " + tab[i, 0] +
                        "\n" + "Twoja bledna odpowiedz: " + slowko_podane + " = " + tab[i, 0] + "\n");
                }
            }
        }
    }
    */

    class KlasaGlowna
    {         
        static void Main()
        {
            Console.WriteLine("Czy na pewno wczytać plik Excel? (y/n) \n");
            string czywczytacexcel = Console.ReadLine();
            if (czywczytacexcel == "y")
            {
                WczytajExcel wczytane = new WczytajExcel();
                string[,] tab_excel = wczytane.CzytajDane();
                Console.WriteLine("\nElemt 1,1= \n" + tab_excel[2, 1]);


            }
                          
            /* INNA TABLICA - SPRAWDZIC
            ArrayList szContents  = new ArrayList();
            string szElement0 = "Tabela Pole 0";
            ArrayList szArray = new ArrayList ();
            szArray.Add (szElement0);
            string szElement1 = "Tabela Pole 1";
            szArray.Add(szElement1);
            Console.WriteLine (szArray[1]);
            */

            Console.WriteLine("Czy na pewno wczytać plik Txt? (y/n) \n");
            string czywczytactxt = Console.ReadLine();
            if (czywczytactxt == "y")
            {
                WczytajTxt wczytane_txt = new WczytajTxt();
                wczytane_txt.CzytajDaneTabTxt();
            }

            Console.WriteLine("Czy uruchomic symulacje? (y/n) \n");
            string czysymulacja = Console.ReadLine();
            if (czysymulacja == "y")
            {
                int m = 4;
                int n = 10;
                string wiersz;
                string[,] tab = new string[n, m];

                //ZAPISYWANIE DO TABLICY INNE
                //string[] wierszTab = new string[5] { "jeden", "dwa", "trzy", "cztery", "piec" };
                

                for (int i = 0; i < n; i++)
                {
                    for (int j = 0; j < m; j++)
                    {
                        string pozycja = "tab" + i + "," + j;
                        tab[i, j] = pozycja;
                    }
                }

                tab[0, 0] = "samochód";
                tab[0, 1] = "car";
                tab[0, 2] = "0";
                tab[0, 3] = "0";

                tab[1, 0] = "dupa";
                tab[1, 1] = "ass";
                tab[1, 2] = "0";
                tab[1, 3] = "0";

                tab[2, 0] = "czarny";
                tab[2, 1] = "black";
                tab[2, 2] = "0";
                tab[2, 3] = "0";

                //WYSWIETLENIE tab
                for (int i = 0; i < n; i++)
                {
                    for (int j = 0; j < m; j++)
                    {
                        Console.Write(tab[i, j] + "\t");
                    }
                    Console.WriteLine();
                }


                for (int i = 0; i < n; i++)
                {

                    Console.Write("Przetlumacz słowo: " + tab[i, 0] + "\n" + "Odpowiedz: ");
                    string slowko_podane = Console.ReadLine();
                    if (tab[i, 1] == slowko_podane)
                    {
                        Console.Write("Poprawnie: " + tab[i, 1] + " = " + tab[i, 0] + "\n");
                        tab[i, 2] = "1";       // poprawnie
                        tab[i, 3] = "1";       // nie powtarzaj
                    }
                    else
                    {
                        tab[i, 2] = "2";      // nie poprawnie
                        tab[i, 3] = "2";      // powtarzaj
                        Console.Write("Nie Poprawnie\n" + "Prawidlowa odopwiedz: " + tab[i, 1] + " = " + tab[i, 0] +
                            "\n" + "Twoja bledna odpowiedz: " + slowko_podane + " = " + tab[i, 0] + "\n");
                    }
                }
            }


            /*
                // zapisujemy posortowane linie
            FileStream fsOutput = new FileStream (szDestFile,
                FileMode.Create, FileAccess.Write);
            StreamWriter srOutput = new StreamWriter (fsOutput);
            for (int nIndex = 0; nIndex < szContents.Count; nIndex++)
            {
                // zapisanie linii do pliku wyjściowego
                srOutput.WriteLine (szContents[nIndex]);
            }
            srOutput.Close ();
            fsOutput.Close ();
            */
 
            /*
   
    class Slowko
    {
        public string slowko_org;
        public string slowko_pol;
        public int czy_poprawnie;
        public int czy_powtorzyc;

        public void Powtorz()
        {
            czy_powtorzyc = 1;
        }
        
        public void Nie_Powtarzaj()
        {
            czy_powtorzyc = 2;
        }
        
        public void Poprawnie()
        {
            czy_poprawnie = 1;
        }

        public void Nie_Poprawnie()
        {
            czy_poprawnie = 2;
        }
    }
              
         
            Slowko slowko1 = new Slowko();
            slowko1.slowko_pol      = tab[0,0];
            slowko1.slowko_org      = tab[0,1];
            slowko1.czy_powtorzyc   = int.Parse(tab[0,2]);
            slowko1.czy_poprawnie   = int.Parse(tab[0,3]);

            Console.Write("Przetlumacz słowo: " + slowko1.slowko_pol + "\n" + "Odpowiedz: ");
            string slowko_podane = Console.ReadLine();
            if (slowko1.slowko_org == slowko_podane)
            {
                Console.Write("Poprawnie: " + slowko1.slowko_org + " = " + slowko1.slowko_pol + "\n");
            }
            else
            {
                slowko1.Nie_Poprawnie();
                slowko1.Powtorz();
                Console.Write("Nie Poprawnie\n" + "Prawidlowa odopwiedz: " + slowko1.slowko_org + " = " + slowko1.slowko_pol + "\n" + "Twoja bledna odpowiedz: " + slowko_podane + " = " + slowko1.slowko_pol + "\n");
            }
             */ 
        }
    }
}
