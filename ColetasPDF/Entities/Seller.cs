using System.Data;
using System.IO;
using System.Linq;

namespace ColetasPDF.Entities
{
    class Seller
    {
        public int Code { get; set; }
        public string Name { get; set; }
        public string Email { get; set; }
        public string Password { get; set; }
        public string Phone { get; set; }

        public Seller(int code)
        {
            Code = code;
        }

        public void GetSeller()
        {
            string CaminhoXML = Directory.GetCurrentDirectory() + @"\ListaEmails.xml";
            DataSet Ds = new DataSet();
            Ds.ReadXml(CaminhoXML);
            DataTable Dt = Ds.Tables[0];

            foreach (DataRow Dr in Dt.Rows)
            {
                if (Code == int.Parse(Dr["codigo"].ToString()))
                {
                    Name = Dr["nome"].ToString();
                    Email = Dr["email"].ToString();
                    Password = Dr["senha"].ToString();
                    Phone = Dr["telefone"].ToString();
               }
            }
            if (Email == "")
            {
                Config config = new Config();
                config.GetConfig();
                Email = config.EmailAccount;
                Password = config.Password;
            }
        }
    }
}
