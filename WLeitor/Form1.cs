using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace WLeitor
{
    public partial class Form1 : Form
    {

        public List<string> listaArquivo = new List<string>();

        public Form1()
        {
            InitializeComponent();
        }

        private void btnSelecionar_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult result = fbd.ShowDialog();

            //Marca o diretório a ser listado
            //DirectoryInfo diretorio = new DirectoryInfo(txtArquivo.Text);
            //Executa função GetFile(Lista os arquivos desejados de acordo com o parametro)
            
            var Arquivos = Directory.EnumerateFiles(fbd.SelectedPath, "*.xml", SearchOption.AllDirectories);

            foreach (var y in Arquivos)
            {
                lstArquivos.Items.Add(Path.GetFileName(y));
                listaArquivo.Add(y);
            }

        }

        private void btnGerar_Click(object sender, EventArgs e)
        {
            if (lstArquivos.Items.Count > 0)
            {
                List<string> lista = new List<string>();

                foreach (string x in listaArquivo)
                {
                    lista.Add(x);
                }

                XmlNamespaceManager nsmgr = new XmlNamespaceManager(new NameTable());
                nsmgr.AddNamespace("n", "http://www.portalfiscal.inf.br/nfe");

                XmlDocument oXML = new XmlDocument();

                XmlNode root;

                XmlNodeList cnpj, nf, codProd, desc, unidade, quantidade, vlr;
                FolderBrowserDialog fbd = new FolderBrowserDialog();

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                int counter = 1, i = 1, l = 2, k = 2, h = 2, g = 2, f = 2, a = 1;


                using (var excelPackage = new ExcelPackage())
                {
                    excelPackage.Workbook.Properties.Title = "Relatório";
                    var sheet = excelPackage.Workbook.Worksheets.Add("Relatório");
                    string caminho;

                    string[] titulos = new string[] {"NUM_NF", "CNPJ_EMITENTE", "SEQ_ITEM", "COD_PRODUTO",
                    "DESCRICAO", "UNIDADE_MEDIDA", "QUANTIDADE", "VLR_BRUTO"};

                    foreach (var titulo in titulos)
                    {
                        sheet.Cells[1, i++].Value = titulo;
                    }

                    i = 2;

                    foreach (var y in lista)
                    {
                        oXML.Load(y);
                        root = oXML.DocumentElement;
                        cnpj = root.SelectNodes("//n:CNPJ", nsmgr);
                        nf = root.SelectNodes("//n:nNF", nsmgr);
                        codProd = root.SelectNodes("//n:cProd", nsmgr);
                        desc = root.SelectNodes("//n:xProd", nsmgr);
                        unidade = root.SelectNodes("//n:uCom", nsmgr);
                        quantidade = root.SelectNodes("//n:qCom", nsmgr);
                        vlr = root.SelectNodes("//n:vProd", nsmgr);

                        foreach (XmlNode oNo in cnpj)
                        {
                            foreach (XmlNode oNo1 in nf)
                            {
                                for (int z = 0; z < codProd.Count; z++)
                                {
                                    sheet.Cells[i, 1].Value = oNo1.ChildNodes.Item(0).InnerText;
                                    sheet.Cells[i, 2].Value = oNo.ChildNodes.Item(0).InnerText;
                                    i++;
                                }
                            }

                            foreach (XmlNode oNo1 in codProd)
                            {
                                sheet.Cells[l, 4].Value = oNo1.ChildNodes.Item(0).InnerText;
                                sheet.Cells[l, 3].Value = counter;
                                counter++;
                                l++;
                            }

                            foreach (XmlNode oNo1 in desc)
                            {
                                sheet.Cells[k, 5].Value = oNo1.ChildNodes.Item(0).InnerText;
                                k++;
                            }


                            foreach (XmlNode oNo1 in unidade)
                            {
                                sheet.Cells[f, 6].Value = oNo1.ChildNodes.Item(0).InnerText;
                                f++;
                            }

                            foreach (XmlNode oNo1 in quantidade)
                            {
                                sheet.Cells[h, 7].Value = oNo1.ChildNodes.Item(0).InnerText;
                                h++;
                            }

                            foreach (XmlNode oNo1 in vlr)
                            {
                                if (a < vlr.Count)
                                {
                                    sheet.Cells[g, 8].Value = oNo1.ChildNodes.Item(0).InnerText;
                                    g++;
                                    a++;
                                }
                            }
                            a = 1;
                            counter = 1;
                            break;
                        }
                    }

                    caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    string path = caminho + @"\Relatorio.xlsx";
                    File.WriteAllBytes(path, excelPackage.GetAsByteArray());
                    MessageBox.Show("Concluído. Verifique em " + Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                }
            }
            else
            {
                MessageBox.Show("Não é possivel gerar o relatório sem selecionar ao menos 1 XML!", "Aviso", MessageBoxButtons.OK);
            }

        }
    }
}
