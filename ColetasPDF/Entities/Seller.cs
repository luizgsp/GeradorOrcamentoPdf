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

            var seller = Dt.AsEnumerable().Where(s => s.Field<string>("codigo") == Code.ToString()).FirstOrDefault();

            Name = seller[1].ToString();
            Email = seller[2].ToString();
            Password = seller[3].ToString();
            Phone = seller[4].ToString();

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
