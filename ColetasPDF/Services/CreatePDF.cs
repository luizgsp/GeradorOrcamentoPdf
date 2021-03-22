using System;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using ColetasPDF.Entities;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace ColetasPDF.Services
{
    class CreatePDF
    {
        private string FirstLettre { get; set; }
        public string Printer1 { get; set; }
        public string Printer2 { get; set; }
        private bool Printer { get; set; }
        private string EmailBody { get; set; }
        private BaseColor FontColor { get; set; }
        private bool AlternateColor { get; set; }
        public int ItemsPerPage { get; set; }
        public Config Config { get; set; } = new Config();
        private Order Order { get; set; }

        public CreatePDF(string printer1, string printer2, int itemsPerPage)
        {
            Printer1 = printer1;
            Printer2 = printer2;
            ItemsPerPage = itemsPerPage;
        }

        public void GerarPDF(string fileTxt)
        {
            

            Config.GetConfig();

            DateTime Dt1 = DateTime.Now;
            
            try
            {
                ReadingFile(fileTxt);
                if (Order == null) { throw new Exception("Erro ao tentar ler a Coleta"); }

                Document documento = new Document(PageSize.A4, 5, 5, 20, 20);
                documento.AddAuthor("Minas Ferramentas Ltda");
                documento.AddSubject("Cotação " + fileTxt.Replace(".txt", "").Replace(".TXT", ""));
                documento.SetPageSize(PageSize.A4.Rotate());

                string ColetaPDF = "MFL Cotação nº " + Order.OrderNumber.ToString("0000")
                    + " em " + Dt1.ToString("dd-MMM-yyyy")
                    + " as " + Dt1.ToString("HH-mm-ss") + ".pdf";
                string ColetaPDFAenviar = ColetaPDF;
                if (File.Exists(Config.SourcePath + ColetaPDF))
                {
                    File.Delete(ColetaPDF);
                }
                ColetaPDF = Config.SourcePath + ColetaPDF;
                PdfWriter Writer = PdfWriter.GetInstance(documento, new FileStream(ColetaPDF, FileMode.Create));

                //=================================================================================================0
                //Construindo o corpo do PDF
                //Seleciona o arquivo para a imagem de marca d'água
                string MarcaDagua = Directory.GetCurrentDirectory() + @"\LOGO_MF_SELO_PB.jpg";
                iTextSharp.text.Image ImgMarcaDagua = iTextSharp.text.Image.GetInstance(MarcaDagua);
                //Informa a posição da Marca d'Água
                ImgMarcaDagua.SetAbsolutePosition(150, 500);

                //Abre o documento
                documento.Open();
                documento.NewPage();
                iTextSharp.text.Rectangle page = documento.PageSize;


                PdfPTable tableBorder = new PdfPTable(1);
                tableBorder.WidthPercentage = 100;

                //IMPRIMINDO O CABECALHO MINAS FERRAMENTAS LTDA
                PdfPTable TableCabec = new PdfPTable(6);
                TableCabec.WidthPercentage = 98;

                string LogoMarca = Directory.GetCurrentDirectory() + @"\LogoMF2.jpg";
                iTextSharp.text.Image ImgMinas = iTextSharp.text.Image.GetInstance(LogoMarca);
                ImgMinas.Alignment = iTextSharp.text.Image.ALIGN_CENTER;
                //ImgMinas.Top = iTextSharp.text.Image.ALIGN_CENTER;

                string LogoFluke = Directory.GetCurrentDirectory() + @"\IMGFluke.jpg";
                iTextSharp.text.Image ImgFluke = iTextSharp.text.Image.GetInstance(LogoFluke);
                ImgFluke.Alignment = iTextSharp.text.Image.ALIGN_CENTER;
                ImgFluke.ScaleAbsolute(100f, 50f);
                //ImgFluke.Top = iTextSharp.text.Image.ALIGN_CENTER;

                PdfPTable tblCabec = new PdfPTable(13);
                tblCabec.WidthPercentage = 100;
                PdfPTable tableClientes = new PdfPTable(1);
                PdfPTable tableClientes2 = new PdfPTable(6);
                PdfPTable tableItens = new PdfPTable(15);
                tableItens.WidthPercentage = 95;

                PdfPTable tableTotais = new PdfPTable(5);
                PdfPTable tableObs = new PdfPTable(6);
                PdfPTable tableObs2 = new PdfPTable(2);


                PdfPCell cell;
                PdfPCell cellCabec = new PdfPCell(ImgMinas);
                cellCabec.Rowspan = 4;
                cellCabec.Colspan = 2;
                cellCabec.Border = 0;
                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                cellCabec.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                TableCabec.AddCell(cellCabec);

                cellCabec = new PdfPCell(ImgFluke);
                cellCabec.Rowspan = 4;
                cellCabec.Colspan = 2;
                cellCabec.Border = 0;
                cellCabec.VerticalAlignment = Element.ALIGN_MIDDLE;
                cellCabec.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                TableCabec.AddCell(cellCabec);

                string[] arrValores = new string[2];
                

                string[] linhaA1 = new string[10];
                int countRows = 0;
                int intRows = 0;
                AlternateColor = false;
                FontColor = BaseColor.BLACK;

                int TotalPaginas = CountPages(Config.SourcePath + fileTxt);
                int ContaPaginas = 1;

                
                BaseColor baseColor = GetColor(AlternateColor);
                //imprime o cabecalho fixo
                // Numero da Cotacao
                AddCell(ref TableCabec, "Cotacao Nr.:" + Order.OrderNumber.ToString(), 0, 2, 0, Element.ALIGN_MIDDLE, Element.ALIGN_CENTER, FontFactory.TIMES_ROMAN, 16, Font.NORMAL, baseColor);
                //Data da cotacao e pagina
                AddCell(ref TableCabec, "Data da Cotação: " + Order.DateOrder.ToString("dd/MM/yyyy") + @" Pag.: " + ContaPaginas.ToString("##0") + @"/" + TotalPaginas, 0, 2, 0, Element.ALIGN_MIDDLE, Element.ALIGN_CENTER, FontFactory.TIMES_ROMAN, 10, Font.NORMAL, baseColor);
                //Vendedor
                AddCell(ref TableCabec, "Vendedor: " + Order.SalesPerson, 0, 2, 0, Element.ALIGN_MIDDLE, Element.ALIGN_CENTER, FontFactory.TIMES_ROMAN, 10, Font.NORMAL, baseColor);
                //Referente
                AddCell(ref TableCabec, "Referente a: " + Order.OrderReference, 0, 2, 0, Element.ALIGN_MIDDLE, Element.ALIGN_LEFT, FontFactory.TIMES_ROMAN, 10, Font.NORMAL, baseColor);

                string DadosEmpresa = @"CNPJ.: 17.194.994/0001-27 IE.:0620080420094 Av. Bias Fortes, 1853 - Barro Preto - Belo Horizonte/MG Cep:30.170-012 Fones:(31)2101-6000 (31)3279-6000 Fax:(31)2101-6010";
                AddCell(ref TableCabec, DadosEmpresa, 0, 6, 0, Element.ALIGN_MIDDLE, Element.ALIGN_MIDDLE, FontFactory.TIMES_ROMAN, 10, Font.NORMAL, baseColor);


                //Dados do cliente
                //Razão Social, CNPJ, IE
                AddCelCabec(ref tableClientes2, "Razão Social", Order.Customer.Name, 4, baseColor);
                AddCelCabec(ref tableClientes2, "CNPJ/CPF", Order.Customer.CnpjCpf, 0, baseColor);
                AddCelCabec(ref tableClientes2, "Inscr. Estadual", Order.Customer.StateRegister, 0, baseColor);
                // Logradouro, Cidade, UF, Cep, Tel, Fax
                AddCelCabec(ref tableClientes2, "Logradouro", Order.Customer.Address, 3, baseColor);
                AddCelCabec(ref tableClientes2, "Cidade", Order.Customer.City, 0, baseColor);
                AddCelCabec(ref tableClientes2, "UF", Order.Customer.Uf, 0, baseColor);
                AddCelCabec(ref tableClientes2, "CEP", Order.Customer.PostalCode, 0, baseColor);
                //Telefones
                foreach (string phone in Order.Customer.Phones)
                {
                    AddCelCabec(ref tableClientes2, "Telefone", phone, 0, baseColor);
                }
                // Email, Att e Mensagem
                AddCelCabec(ref tableClientes2, "E-mail", Order.Customer.Email, 2, baseColor);
                AddCelCabec(ref tableClientes2, "Telefone", Order.OrderReference, 2, baseColor);
                AddCell(ref tableClientes2, Order.Message, 0, 2, 0, Element.ALIGN_TOP, Element.ALIGN_LEFT, FontFactory.TIMES_ROMAN, 10, Font.NORMAL, baseColor);


                // IMPRESSAO DO CABECALHO DOS ITENS
                //TableCabec
                cell = new PdfPCell(TableCabec);
                cell.Colspan = 15;
                cell.VerticalAlignment = Element.ALIGN_TOP;
                cell.HorizontalAlignment = Element.ALIGN_MIDDLE;
                cell.BackgroundColor = baseColor;
                tableItens.AddCell(cell);
                //tableClientes2
                cell = new PdfPCell(tableClientes2);
                cell.Colspan = 15;
                cell.VerticalAlignment = Element.ALIGN_TOP;
                cell.HorizontalAlignment = Element.ALIGN_MIDDLE;
                cell.BackgroundColor = baseColor;
                tableItens.AddCell(cell);
                //tableItens
                tableBorder.AddCell(tableItens);

                countRows = 0;
                intRows = Order.Items.Count;
                foreach (Item item in Order.Items)
                {
                    countRows++;
                    if (countRows > ItemsPerPage)
                    {
                        cell = new PdfPCell(TableCabec);
                        cell.Colspan = 15;
                        cell.VerticalAlignment = Element.ALIGN_TOP;
                        cell.HorizontalAlignment = Element.ALIGN_MIDDLE;
                        cell.BackgroundColor = baseColor;
                        tableItens.AddCell(cell);
                        //tableClientes2
                        cell = new PdfPCell(tableClientes2);
                        cell.Colspan = 15;
                        cell.VerticalAlignment = Element.ALIGN_TOP;
                        cell.HorizontalAlignment = Element.ALIGN_MIDDLE;
                        cell.BackgroundColor = baseColor;
                        tableItens.AddCell(cell);
                        //tableItens
                        tableBorder.AddCell(tableItens);
                    }
                    AlternateColor = !(AlternateColor);
                    baseColor = GetColor(AlternateColor);
                    //Num Item
                    AddCell(ref tableItens, item.ItemSequel.ToString("000"), 1, 1, 0, Element.ALIGN_TOP, Element.ALIGN_CENTER, FontFactory.HELVETICA, 7, Font.NORMAL, baseColor);
                    //Codigo
                    AddCell(ref tableItens, item.ItemCode, 1, 1, 0, Element.ALIGN_TOP, Element.ALIGN_CENTER, FontFactory.HELVETICA, 7, Font.NORMAL, baseColor);
                    //Discriminação
                    AddCell(ref tableItens, item.Description, 1, 6, 0, Element.ALIGN_TOP, Element.ALIGN_CENTER, FontFactory.HELVETICA, 7, Font.NORMAL, baseColor);
                    //Qtde
                    AddCell(ref tableItens, item.Quantity.ToString("F2"), 1, 1, 0, Element.ALIGN_TOP, Element.ALIGN_RIGHT, FontFactory.HELVETICA, 7, Font.NORMAL, baseColor);
                    //Class.Fiscal
                    AddCell(ref tableItens, item.TaxClassification, 1, 1, 0, Element.ALIGN_TOP, Element.ALIGN_CENTER, FontFactory.HELVETICA, 7, Font.NORMAL, baseColor);
                    //% ICMS
                    AddCell(ref tableItens, item.IcmsTax, 1, 1, 0, Element.ALIGN_TOP, Element.ALIGN_CENTER, FontFactory.HELVETICA, 7, Font.NORMAL, baseColor);
                    //Pr.Unit c/ICMS
                    AddCell(ref tableItens, item.UnitPrice.ToString("F2"), 1, 1, 0, Element.ALIGN_TOP, Element.ALIGN_RIGHT, FontFactory.HELVETICA, 7, Font.NORMAL, baseColor);
                    //Pr.Total c/ICMS
                    AddCell(ref tableItens, item.GetTotalPrice().ToString("F2"), 1, 1, 0, Element.ALIGN_TOP, Element.ALIGN_RIGHT, FontFactory.HELVETICA, 7, Font.NORMAL, baseColor);
                    //IPI
                    AddCell(ref tableItens, item.IpiTax, 1, 1, 0, Element.ALIGN_TOP, Element.ALIGN_CENTER, FontFactory.HELVETICA, 7, Font.NORMAL, baseColor);
                    //Prazo de Entrega
                    AddCell(ref tableItens, item.DeliverTime, 1, 1, 0, Element.ALIGN_TOP, Element.ALIGN_CENTER, FontFactory.HELVETICA, 7, Font.NORMAL, baseColor);

                }
                countRows++;
                if (countRows > ItemsPerPage)
                {
                    cell = new PdfPCell(TableCabec);
                    cell.Colspan = 15;
                    cell.VerticalAlignment = Element.ALIGN_TOP;
                    cell.HorizontalAlignment = Element.ALIGN_MIDDLE;
                    cell.BackgroundColor = baseColor;
                    tableItens.AddCell(cell);
                    //tableClientes2
                    cell = new PdfPCell(tableClientes2);
                    cell.Colspan = 15;
                    cell.VerticalAlignment = Element.ALIGN_TOP;
                    cell.HorizontalAlignment = Element.ALIGN_MIDDLE;
                    cell.BackgroundColor = baseColor;
                    tableItens.AddCell(cell);
                    //tableItens
                    tableBorder.AddCell(tableItens);
                }
                baseColor = GetColor(false);
                // IMPRESSO CABEÇALHO DOS TOTAIS
                Notes noteIpi = Order.Notes[0];
                AddCell(ref tableTotais, noteIpi.Value, 1, 2, 2, Element.ALIGN_TOP, Element.ALIGN_CENTER, FontFactory.HELVETICA, 8, Font.NORMAL, baseColor);

                //Valor da Mão de Obra
                AddCell(ref tableTotais, @"Valor da Mão de Obra", 1, 1, 0, Element.ALIGN_TOP, Element.ALIGN_CENTER, FontFactory.HELVETICA, 8, Font.BOLD, baseColor);
                //Total do IPI
                AddCell(ref tableTotais, @"Valor do IPI", 1, 1, 0, Element.ALIGN_TOP, Element.ALIGN_CENTER, FontFactory.HELVETICA, 8, Font.BOLD, baseColor);
                //Total de Mercadorias
                AddCell(ref tableTotais, @"Total de Mercadorias", 1, 1, 0, Element.ALIGN_TOP, Element.ALIGN_CENTER, FontFactory.HELVETICA, 8, Font.BOLD, baseColor);
                //Total Geral
                AddCell(ref tableTotais, @"Total Geral", 1, 1, 0, Element.ALIGN_TOP, Element.ALIGN_CENTER, FontFactory.HELVETICA, 8, Font.BOLD, baseColor);
                //Valor da Mão de Obra
                AddCell(ref tableTotais, Order.LaborValue.ToString("F2"), 1, 1, 0, Element.ALIGN_TOP, Element.ALIGN_CENTER, FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLUE);
                //Total do IPI
                AddCell(ref tableTotais, "0,00", 1, 1, 0, Element.ALIGN_TOP, Element.ALIGN_CENTER, FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLUE);
                //Total de Mercadorias
                AddCell(ref tableTotais, Order.GetTotais().ToString("F2"), 1, 1, 0, Element.ALIGN_TOP, Element.ALIGN_CENTER, FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLUE);
                //Total Geral
                AddCell(ref tableTotais, Order.GetTotais().ToString("F2"), 1, 1, 0, Element.ALIGN_TOP, Element.ALIGN_CENTER, FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLUE);
                tableBorder.AddCell(tableTotais);

                foreach (Notes note in Order.Notes)
                {
                    if (!string.IsNullOrEmpty(note.Name))
                    {
                        AddCelCabec(ref tableObs, note.Name, note.Value, 0, baseColor);
                    }
                    else
                    {
                        AddCell(ref tableObs2, note.Value, 1, 1, 0, Element.ALIGN_TOP, Element.ALIGN_LEFT, FontFactory.HELVETICA, 7, Font.NORMAL, baseColor);
                    }
                }

                cell = new PdfPCell(tableObs);
                cell.Border = 0;
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                tableBorder.AddCell(cell);

                AddCell(ref tableObs2, @"DISTRIBUIDOR AUTORIZADO de ferramentas de metal duro  LAMINA, Consulte-nos", 1, 1, 0, Element.ALIGN_TOP, Element.ALIGN_LEFT, FontFactory.HELVETICA, 7, Font.NORMAL, baseColor);
                AddCell(ref tableObs2, @"ESPECIALIZADA EM FERRAMENTAS PARA INDUSTRIA E MECANICA EM GERAL", 1, 1, 0, Element.ALIGN_TOP, Element.ALIGN_LEFT, FontFactory.HELVETICA, 7, Font.NORMAL, baseColor);

                cell = new PdfPCell(tableObs2);
                cell.Border = 0;
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                tableBorder.AddCell(cell);

                documento.Add(tableBorder);
                // fecha o documento
                Writer.Flush();
                //documento.Close();

                PrintDocumment(FirstLettre, fileTxt, Order.SellerCode, ColetaPDFAenviar);

            }
            catch (IOException e)
            {
                throw new IOException(fileTxt + " Erro->" + e.Message);
            }
            catch (Exception e)
            {
                throw new Exception(fileTxt + " Erro->" + e.Message);
            }

        }

        private void ReadingFile(string fileTxt)
        {
            string strLettrePrevious = "";
            string strDetails, strLettre, PrimeiroItemCabec = "";
            int intRows = 0;

            try
            {
                if (Config == null) { Config = new Config(); }
                Config.GetConfig();
                if (Order == null) { Order = new Order(); }
                Customer customer = new Customer();
                Item items;

                StreamReader srArquivo = new StreamReader(Config.SourcePath + fileTxt);
                while (!srArquivo.EndOfStream)
                {
                    string strRows = srArquivo.ReadLine();
                    if (!string.IsNullOrEmpty(strRows))
                    {
                        strLettre = strRows.Substring(0, 1);
                        if (strLettrePrevious != strLettre) { strLettrePrevious = strLettre; }
                        strDetails = strRows.Substring(1, (strRows.Length - 1)); //Conteúdo

                        if (intRows == 0)
                        {
                            Order.SellerCode = int.Parse(strRows.Substring(1, 3));
                            FirstLettre = strLettre;
                        }
                        if (strRows.Length > 1 && intRows != 0)
                        {
                            switch (strLettre)
                            {
                                case "A": // Cabeçalho da Cotação
                                    string[] lineA = strDetails.Split(char.Parse(":"));
                                    if (strDetails.Contains("Cotacao Nr.:")) { Order.OrderNumber = int.Parse(lineA[1].Trim()); }
                                    if (strDetails.Contains("Data Cotacao.:")) { Order.DateOrder = DateTime.Parse(lineA[1].Trim()); }
                                    if (strDetails.Contains("Vendedor:")) { Order.SalesPerson = lineA[1].Trim(); }
                                    if (strDetails.Contains("Referente a:")) { Order.OrderReference = lineA[1].Trim(); }
                                    break;
                                case "B": // Dados do cliente
                                    string[] strCustomer = strDetails.Split(char.Parse(";"));
                                    foreach (string lineB in strCustomer)
                                    {
                                        string[] cols = lineB.Split(char.Parse(":"));
                                        if (cols[0].Contains("Razao Social")) { customer.Name = cols[1].Trim(); }
                                        if (cols[0].Contains("CNPJ/CPF")) { customer.CnpjCpf = cols[1].Trim(); }
                                        if (cols[0].Contains("Inscr. Estadual")) { customer.StateRegister = cols[1].Trim(); }
                                        if (cols[0].Contains("Logradouro")) { customer.Address = cols[1].Trim(); }
                                        if (cols[0].Contains("Cidade")) { customer.City = cols[1].Trim(); }
                                        if (cols[0].Contains("UF")) { customer.Uf = cols[1].Trim(); }
                                        if (cols[0].Contains("CEP")) { customer.PostalCode = cols[1].Trim(); }
                                        if (cols[0].Contains("Tel.")) { customer.Phones.Add(cols[1].Trim()); }
                                        if (cols[0].Contains("Fax.")) { customer.Phones.Add(cols[1].Trim()); }
                                        if (cols[0].Contains("E-mail")) { customer.Email = cols[1].Trim(); }
                                        if (cols[0].Contains("ATT.Sr(a).")) { customer.Contact = cols[1].Trim(); }
                                    }
                                    Order.Message = strDetails.Trim();


                                    break;
                                case "C": // Itens da Cotação
                                    if (Order.Customer == null) { Order.Customer = customer; }

                                    PrimeiroItemCabec = strDetails;
                                    if (!PrimeiroItemCabec.Contains("Item"))
                                    {
                                        string[] strItems = strDetails.Split(char.Parse(";"));

                                        items = new Item(
                                                int.Parse(strItems[0]),
                                                strItems[1].Trim(),
                                                strItems[2].Trim(),
                                                double.Parse(strItems[3]),
                                                strItems[4].Trim(),
                                                strItems[5].Trim(),
                                                 double.Parse(strItems[6]),
                                                strItems[8].Trim(),
                                                strItems[9].Trim()
                                            );
                                        Order.Items.Add(items);
                                    }

                                    break;
                                case "D": // Totalizador
                                    string[] lineD = strDetails.Split(char.Parse(";"));
                                    Order.Notes.Add(new Notes("IpiObs", lineD[0].Trim()));
                                    Order.LaborValue = double.Parse(lineD[1]);
                                    break;
                                case "E": //

                                    break;
                                case "F": // Informações Gerais
                                    string[] lineF = strDetails.Split(char.Parse(":"));
                                    Order.Notes.Add(new Notes(lineF[0], lineF[1].Trim()));
                                    break;
                                case "G": // Observações para o cliente
                                    Order.Notes.Add(new Notes("Obs", strDetails.Trim()));
                                    break;

                            }
                        }
                        intRows++;
                    }
                }

                srArquivo.Close();


            }
            catch (IOException e)
            {
                throw new Exception(fileTxt + " Erro->" + e.Message);
            }
            catch (Exception e)
            {
                throw new Exception(fileTxt + " Erro->" + e.Message);
            }
        }

        private void AddCell(ref PdfPTable table, string content, int border, int colspan, int rowspan, int verticalAlignment, int horizontalAlignment, string fontFactory, float tamanho, int fontStyle, BaseColor color)
        {
            PdfPCell cell = new PdfPCell(new Phrase(content, FontFactory.GetFont(fontFactory, tamanho, fontStyle)));
            cell.Border = border;
            if (colspan > 0) cell.Colspan = colspan;
            if (rowspan > 0) cell.Rowspan = rowspan;
            cell.VerticalAlignment = verticalAlignment;
            cell.HorizontalAlignment = horizontalAlignment;
            cell.BackgroundColor = color;
            table.AddCell(cell);
        }

        private void AddCelCabec(ref PdfPTable table, string title, string content, int colspan, BaseColor color)
        {
            Chunk c1 = new Chunk(title, FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL));
            Chunk c2 = new Chunk(content, FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8, Font.NORMAL));
            Phrase NovaLinha = new Phrase(Environment.NewLine);
            Paragraph paragrafo = new Paragraph();
            paragrafo.Add(c1);
            paragrafo.Add(NovaLinha);
            paragrafo.Add(c2);
            PdfPCell cell = new PdfPCell(paragrafo);
            if (colspan > 0) cell.Colspan = colspan;
            cell.VerticalAlignment = Element.ALIGN_TOP;
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            cell.BackgroundColor = color;
            table.AddCell(cell);
        }

        private void PrintDocumment(string OndeGerar, string ColetaTXT, int CodVendedor, string ColetaPDFAenviar)
        {
            string enviaMensagem = "";
            int stepsEmail = 0;
            string Motivo = "";

            Seller seller = new Seller(CodVendedor);
            seller.GetSeller();

            try
            {
                if (Config == null) { Config.GetConfig(); }

                string sourceFile = Config.SourcePath + @"\" + ColetaPDFAenviar;


                string targetFile = @"F:\Aenviar\" + Order.Customer.CnpjCpf.Replace(".", "").Replace("-", "").Replace("/", "").Replace(" ", "") +
                    @"\" + Order.OrderNumber + "-" + DateTime.Now.Hour.ToString("00") + "-" +
                    DateTime.Now.Minute.ToString("00") + "-" + DateTime.Now.Second.ToString("00") + ".pdf";

                string existeCaminhoDestino = @"F:\Aenviar\" + Order.Customer.CnpjCpf.Replace(".", "").Replace("-", "").Replace("/", "").Replace(" ", "");
                if (!Directory.Exists(existeCaminhoDestino))
                {
                    Directory.CreateDirectory(existeCaminhoDestino);
                }

                //============== imprime na impressora 1 (Recepção) ======================================================================================
                if (OndeGerar == "C" || OndeGerar == "Q" || OndeGerar == "Y")
                {
                    if (File.Exists(targetFile)) { File.Delete(targetFile); }
                    try
                    {
                        File.Move(sourceFile, targetFile);
                        Console.WriteLine();
                        Console.WriteLine("Movido o PDF para a pasta do Cliente {0} ", Order.Customer.CnpjCpf);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(sourceFile + " ++++> " + e.Message);
                        return;
                    }
                    System.Threading.Thread.Sleep(1500);
                    Console.WriteLine("Aguarde imprimindo coleta...{0}", ColetaPDFAenviar);
                    Process proc = new Process();
                    proc.StartInfo.CreateNoWindow = false;
                    proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    proc.StartInfo.Verb = "print";
                    proc.StartInfo.FileName = targetFile;
                    proc.Start();
                    proc.WaitForInputIdle();
                    proc.CloseMainWindow();
                    proc.Close();
                    Console.WriteLine();
                    Console.WriteLine("Enviado o Arquivo {0} para impressora.", ColetaPDFAenviar);
                }
                else
                {
                    //================== move para o aenviar a coleta do vendedor =================================================================================
                    //System.Threading.Thread.Sleep(2000);
                    if (File.Exists(targetFile)) { File.Delete(targetFile); }
                    try
                    {
                        File.Move(sourceFile, targetFile);
                        Console.WriteLine();
                        Console.WriteLine("Movido o PDF para a pasta do Cliente {0} ", Order.Customer.CnpjCpf);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(sourceFile + " ++++> " + e.Message);
                        return;
                    }


                    if (IsValidEmail(Order.Customer.Email))
                    {

                        if (Order.Customer.Contact.Trim() == "") { Order.Customer.Contact = "Comprador"; }
                        if (Order.Customer.Email.Trim() == "") { Order.Customer.Email = "vendas@myportal.com.br"; }

                        string Subject = Order.OrderNumber + " " + Order.Customer.Name;
                        string[] Note1 = new string[3];
                        int i = 0;
                        foreach (Notes note in Order.Notes)
                        {
                            if (note.Name == "Obs")
                            {
                                Note1[i] = note.Value;
                                i++;
                            }
                        }
                        enviaMensagem = Config.EmailBody + "<br /><br />" +
                                Note1[0] +
                                "<br /><br />" +
                                Note1[1] +
                                "Atenciosamente," + "<br />" +
                                  seller.Name + "<br />" +
                                "Departamento de Vendas." + "<br />" +
                                seller.Email + "<br />" +
                                "Tel.: " +
                                seller.Phone +
                                " (31) 2101.6000  / Fax: (31) 2101.6010<br />" +
                                "Av. Bias Fortes, 1853 | B. Barro Preto | Belo Horizonte - MG | Cep 30170-012 ";

                        // envia a mensagem para o cliente
                        SendMail sendMail = new SendMail(seller.Email, enviaMensagem, Order.Customer.Email, Subject, targetFile, seller.Email, seller.Password, Priority.Normal);
                        stepsEmail = sendMail.Mailing() ? 1 : 0;

                        // envia uma copia para a conta copiadeemail@myportal.com.br
                        sendMail = new SendMail(
                            seller.Email,
                            enviaMensagem,
                            emailCustomer: "copiadeemail@myportal.com.br",
                            Subject,
                            targetFile,
                            seller.Name,
                            seller.Password,
                            Priority.Normal
                        );
                        stepsEmail = sendMail.Mailing() ? 2 : 0;

                        // envia confirmacao para o vendedor de mensagem enviada OK
                        Subject = "OK - " + Order.OrderNumber + " " + Order.Customer.Name + " " + DateTime.Now.ToString();
                        enviaMensagem = Order.OrderNumber + "<br />" + Order.Customer.CnpjCpf + "-" + Order.Customer.Name
                            + "<br /><b>Enviada com sucesso para:</b><br />" + Order.Customer.Email
                             + "<br />Em: " + DateTime.Now.ToString() + ".<br />" +
                             Note1 + ".";

                        sendMail = new SendMail(
                            emailAccount: Config.EmailAccount,
                            emailBody: enviaMensagem,
                            emailCustomer: seller.Email,
                            Subject,
                            targetFile,
                            sellerName: "Vendas",
                            password: Config.Password,
                            Priority.Normal
                        );
                        stepsEmail = sendMail.Mailing() ? 3 : 0;

                        Console.WriteLine();
                        Console.WriteLine(Subject);
                    }
                    else
                    {
                        //caso o email do cliente nao seja valido envia um email para o vendedor avisando
                        string Subject = "((ERRO)) - " + Order.OrderNumber + " " + Order.Customer.Name + " " + DateTime.Now.ToString();
                        enviaMensagem = Order.OrderNumber + "<br />" + Order.Customer.CnpjCpf + "-" + Order.Customer.Name + " - Vendedor: " + Order.SalesPerson
                        + "<br /><b>NAO enviada para:</b><br />" + Order.Customer.Email + "<br />Motivo: <b>Email Invalido.</b>" +
                        "<br />Em: " + DateTime.Now.ToString() + ".";
                        SendMail sendMail = new SendMail(
                            emailAccount: Config.EmailAccount,
                            emailBody: enviaMensagem,
                            emailCustomer: seller.Email,
                            Subject,
                            targetFile,
                            sellerName: "Vendas",
                            password: Config.Password,
                            Priority.High
                        );
                        stepsEmail = sendMail.Mailing() ? 3 : 0;
                    }
                }
            }
            catch (Exception e)
            {
                string Subject = "((ERRO)) - " + Order.OrderNumber + " " + Order.Customer.Name + " " + DateTime.Now.ToString();
                enviaMensagem = Order.OrderNumber + "<br />" + Order.Customer.CnpjCpf + "-" + Order.Customer.Name + " - Vendedor: " + Order.SalesPerson
                + "<br /><b>NAO enviada para:</b><br />" + Order.Customer.Email + "<br />Motivo: " + e.Message +
                        "<br />Em: " + DateTime.Now.ToString() + ".";

                SendMail sendMail = new SendMail(
                            emailAccount: Config.EmailAccount,
                            emailBody: enviaMensagem,
                            emailCustomer: seller.Email,
                            Subject,
                            "",
                            sellerName: "Vendas",
                            password: Config.Password,
                            Priority.High
                        );
                stepsEmail = sendMail.Mailing() ? 4 : 0;

                Motivo = e.Message;
            }
            finally
            {
                // ========= copia o arquivo txt para a pasta impressos
                string sourceFile = Path.Combine(Config.SourcePath, ColetaTXT);
                string destFile = Path.Combine(Config.TargetPath, ColetaTXT);
                try
                {
                    if (!Directory.Exists(Config.TargetPath))
                    {
                        Directory.CreateDirectory(Config.TargetPath);
                    }
                    if (File.Exists(destFile))
                    {
                        File.Delete(destFile);
                    }
                    File.Move(sourceFile, destFile);
                    Console.WriteLine("Movido arquivo de {0} para {1}", sourceFile, destFile);
                }
                catch (Exception e)
                {
                    Console.WriteLine(sourceFile + " >>> " + e.Message);
                }

                if (stepsEmail != 0)
                {
                    //gera txt com os dados do envio.
                    StreamWriter s = File.AppendText(Config.TargetPath + @"email.txt");
                    string cnpjCpf = (Order.Customer.CnpjCpf.Replace(".", "").Replace("-", "").Replace("/", "").Replace(" ", "") + "00").Trim();
                    if (cnpjCpf.Length == 13)
                    {
                        cnpjCpf = cnpjCpf + "000";
                    }

                    string linha = "|" + CodVendedor.ToString("0000") + cnpjCpf + Order.OrderNumber.ToString("0000") + FillFields(Order.Customer.Email.Trim(), 50) +
                            DateTime.Now.ToString("ddMMyy") + stepsEmail + DateTime.Now.ToString("HHmmss") + "00";

                    if (stepsEmail == 1)
                        linha = linha + ("Enviada com sucesso para: " + FillFields(Order.Customer.Email.Trim(), 50).Trim() + " " +
                            cnpjCpf + "-" + Order.Customer.Name.Trim() + " Em: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") +
                            ".=================================================").Substring(0, 100);

                    else
                        linha = linha + ("Motivo: " + Motivo + " " +
                        cnpjCpf + "-" + Order.Customer.Name.Trim() + " Em: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") +
                        ".==========================================================================================================" +
                        "=================================================================================================").Substring(0, 100);
                    s.WriteLine(linha);
                    s.Close();
                }
            }
        }

        private int CountPages(string caminho)
        {
            StreamReader srArquivo = new StreamReader(caminho);
            int ContaLinhas = 0;
            int ContaPaginas = 0;
            while (srArquivo.Peek() != -1)
            {
                string strLinha = srArquivo.ReadLine();
                if (strLinha != "")
                {
                    int TamanhoLinha = strLinha.Length;
                    string Letra = strLinha.Substring(0, 1);
                    if (TamanhoLinha > 1)
                    {
                        switch (Letra)
                        {
                            case "C":
                                ContaLinhas++;
                                break;
                        }
                    }
                }
            }
            if (ContaLinhas > 28)
            {
                ContaPaginas = (int)ContaLinhas / 28;
                if ((ContaLinhas % 28) != 0) { ContaPaginas++; }
            }
            else
                ContaPaginas = 1;

            return ContaPaginas;
        }

        private BaseColor GetColor(bool alternar)
        {
            if (alternar)
            {
                //CorFonte = BaseColor.WHITE;
                FontColor = BaseColor.BLACK;
                return BaseColor.WHITE;
            }
            else
            {
                FontColor = BaseColor.BLACK;
                return BaseColor.LIGHT_GRAY;
            }
        }

        private bool IsValidEmail(string email)
        {
            try
            {
                //define a expressão regulara para validar o email
                string texto_Validar = email;
                Regex expressaoRegex = new Regex(@"\w+@[a-zA-Z_0-9-]+?\.[a-zA-Z]{2,3}");

                // testa o email com a expressão
                if (expressaoRegex.IsMatch(texto_Validar))
                {
                    // o email é valido
                    return true;
                }
                else
                {
                    // o email é inválido
                    return false;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private string FillFields(string campo, int tamanho)
        {
            int resto = tamanho - campo.Length;
            while (resto != 0)
            {
                campo += " ";
                resto--;
            }
            return campo;
        }
    }
}
