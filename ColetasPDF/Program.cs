using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ColetasPDF.Entities;
using ColetasPDF.Services;

namespace ColetasPDF
{
    class Program
    {
        static void Main(string[] args)
        {
            Process aProcess = Process.GetCurrentProcess();
            string aProcName = aProcess.ProcessName;

            if (Process.GetProcessesByName(aProcName).Length > 1)
            {
                Console.WriteLine("O programa já está em execução!");
                System.Threading.Thread.Sleep(5000);
                return;
            }

            while (ListaArquivos() != 0)
            {
                Console.WriteLine(@"\t\t\t\tRelatorio impresso...");
            }
        }

        private static int ListaArquivos()
        {

            System.Threading.Thread.Sleep(2000);

            Config config = new Config();
            config.GetConfig();

            Console.Clear();
            Console.WriteLine("\t\t\tMinas Ferramentas Ltda.");
            Console.WriteLine();
            Console.WriteLine("\t      Impressao e Envio de emails de Coletas em PDF");
            Console.WriteLine();
            Console.WriteLine("\t     Belo Horizonte, {0}", DateTime.Now.ToString("dd 'de' MMMM 'de' yyyy HH:mm:ss"));
            Console.WriteLine();


            DirectoryInfo directory = new DirectoryInfo(config.SourcePath);
            if (!directory.Exists)
            {
                Console.WriteLine("Não encontrado o diretorio especificado!");
                return 0;
            }

            FileInfo[] files = directory.GetFiles("*.txt");
            Console.WriteLine();
            Console.WriteLine(@"Arquivos.: {0}", files.Count());
            Console.WriteLine();
            if (files.Count() == 0)
            {
                try
                {
                    files = directory.GetFiles("*.pdf");
                    if (files.Count() > 0)
                    {
                        foreach (FileInfo fileinfo in files)
                        {
                            System.Threading.Thread.Sleep(5000);
                            File.Move(config.SourcePath + fileinfo, config.TargetPath + fileinfo);
                            Console.WriteLine(@"Movido o Arquivo {0} para a pasta Impressos.", fileinfo);

                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    System.Threading.Thread.Sleep(10000);

                }
                return 1;
            }

            foreach (FileInfo fileinfo in files)
            {
                decimal FileLenth = fileinfo.Length / 1024;
                string DadosDoArquivo = fileinfo.Name + "\t" + FileLenth.ToString("##,##0.00") + " kb" +
                    "\t" + fileinfo.LastWriteTime;
                if (FileLenth > 0)
                {
                    Console.WriteLine(@"Txt {0}", DadosDoArquivo);
                    CreatePDF gPdf = new CreatePDF("", "", 28);
                    gPdf.GerarPDF(fileinfo.Name);
                }
                else
                {
                    fileinfo.Delete();
                }
            }

            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine("\t\t\t\tAguarde....");
            return 1;
        }
    }
}
