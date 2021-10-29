﻿using OfficeOpenXml;
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
        public string a2;
        public Form1()
        {
            InitializeComponent();
        }

        private void btnSelecionar_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult result = fbd.ShowDialog();

            var Arquivos = Directory.EnumerateFiles(fbd.SelectedPath, "*.xml", SearchOption.AllDirectories);

            foreach (var y in Arquivos)
            {
                lstArquivos.Items.Add(Path.GetFileName(y));
                listaArquivo.Add(y);
            }

        }

        private void btnGerar_Click(object sender, EventArgs e)
        {
            try
            {
                if (lstArquivos.Items.Count > 0)
                {
                    List<string> lista = new List<string>();
                    StreamWriter pod;
                    string caminho, path;

                    caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    path = caminho + @"\Relatorio.txt";
                    pod = File.CreateText(path);

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

                    List<string> a1, b1, c1, d1, e1, f1, g1, h1;

                    a1 = new List<string>();
                    b1 = new List<string>();
                    c1 = new List<string>();
                    d1 = new List<string>();
                    e1 = new List<string>();
                    f1 = new List<string>();
                    g1 = new List<string>();
                    h1 = new List<string>();

                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    int counter = 1, i = 1, l = 2, k = 2, h = 2, g = 2, f = 2, a = 1, w1 = 0;

                    pod.WriteLine("|NUM_NF|CNPJ_EMITENTE|SEQ_ITEM|COD_PRODUTO|DESCRICAO|UNIDADE_MEDIDA|QUANTIDADE|VLR_BRUTO|");

                    i = 2;

                    foreach (var y in lista)
                    {
                        a1.Clear();
                        b1.Clear();
                        c1.Clear();
                        d1.Clear();
                        e1.Clear();
                        f1.Clear();
                        g1.Clear();
                        h1.Clear();

                        a2 = y.ToString();
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
                                    a1.Add(oNo1.ChildNodes.Item(0).InnerText.ToString());
                                    b1.Add(oNo.ChildNodes.Item(0).InnerText.ToString());
                                    i++;
                                }
                            }

                            foreach (XmlNode oNo2 in codProd)
                            {
                                d1.Add(oNo2.ChildNodes.Item(0).InnerText.ToString());
                                c1.Add(counter.ToString());
                                counter++;
                                l++;
                            }

                            foreach (XmlNode oNo3 in desc)
                            {
                                e1.Add(oNo3.ChildNodes.Item(0).InnerText.ToString());
                                k++;
                            }

                            foreach (XmlNode oNo4 in unidade)
                            {
                                f1.Add(oNo4.ChildNodes.Item(0).InnerText.ToString());
                                f++;
                            }

                            foreach (XmlNode oNo5 in quantidade)
                            {
                                g1.Add(oNo5.ChildNodes.Item(0).InnerText.ToString());
                                h++;
                            }

                            foreach (XmlNode oNo6 in vlr)
                            {
                                if (a < vlr.Count)
                                {
                                    h1.Add(oNo6.ChildNodes.Item(0).InnerText.ToString());
                                    g++;
                                    a++;
                                }
                            }
                            a = 1;
                            counter = 1;

                            //writer here
                            foreach (var val in a1)
                            {
                                var s9 = "|" + a1[w1].ToString() + "|" + b1[w1].ToString() + "|" + c1[w1].ToString() + "|" + d1[w1].ToString() + "|" + e1[w1].ToString() + "|" + f1[w1].ToString() + "|" + g1[w1].ToString() + "|" + h1[w1].ToString() + "|";
                                pod.WriteLine(s9);
                                pod.Flush();
                                w1++;
                            }
                            w1 = 0;

                            break;
                        }
                    }

                    MessageBox.Show("Concluído. Verifique em " + Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

                }
                else
                {
                    MessageBox.Show("Não é possivel gerar o relatório sem selecionar ao menos 1 XML!", "Aviso", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Arqivo que apresentou erro:" +  a2, "Erro", MessageBoxButtons.OKCancel);
            }
        }
    }
}
