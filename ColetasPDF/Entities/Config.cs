using System;
using System.Data;
using System.IO;

namespace ColetasPDF.Entities
{
    class Config
    {

        public string SourcePath { get; set; }
        public string TargetPath { get; set; }
        public string SmtpServer { get; set; }
        public string EmailAccount { get; set; }
        public string EmailBody { get; set; }
        public string Password { get; set; }
        public int SmtpPort { get; set; }
        public string Printer1 { get; set; }
        public string Printer2 { get; set; }

        public Config()
        {
        }

        public void GetConfig()
        {
            try
            {
                System.Threading.Thread.Sleep(2000);
                string CaminhoXML = Directory.GetCurrentDirectory() + @"\Config.xml";

                DataSet Ds = new DataSet();
                Ds.ReadXml(CaminhoXML);
                DataTable Dt = Ds.Tables[0];
                DataRow Dr = Dt.Rows[0];
                if (Dr != null)
                {
                    SourcePath = Dr["CaminhoOrigem"].ToString();
                    TargetPath = Dr["CaminhoDestino"].ToString();
                    SmtpServer = Dr["ServidorSMTP"].ToString();
                    EmailAccount = Dr["ContaEmail"].ToString();
                    Password = Dr["Pass"].ToString();
                    SmtpPort = int.Parse(Dr["PortaSMTP"].ToString());
                    Printer1 = Dr["CaminhoImpressora1"].ToString();
                    Printer2 = Dr["CaminhoImpressora2"].ToString();
                    EmailBody = Dr["CorpoEmail"].ToString();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Ocorreu um erro: " + ex.Message);
            }
        }

    }
}
