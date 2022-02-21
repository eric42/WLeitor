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
using WLeitor.ECD;

namespace WLeitor
{
    public partial class frmParametros : Form
    {
        public List<string> dados = new List<string>();
        public bool deuErro = false;
        string sUser, sPwd, sVBanco;
        public string sRetornoErro = "";
        public DataTable oDt = new DataTable();

        public frmParametros()
        {
            InitializeComponent();
        }

        public frmParametros(List<string> dado)
        {
            InitializeComponent();
            dado = dados;
        }

        private void dtIni_ValueChanged(object sender, EventArgs e)
        {
            if (dtIni.Value > dtFim.Value)
            {
                MessageBox.Show("Data inicial não pode ser maior que a data final", "Aviso", MessageBoxButtons.OK);
                dtIni.Focus();
            }
        }

        private void dtFim_ValueChanged(object sender, EventArgs e)
        {
            if (dtFim.Value < dtIni.Value)
            {
                MessageBox.Show("Data final não pode ser menor que a data inicial", "Aviso", MessageBoxButtons.OK);
                dtFim.Focus();
            }
        }

        private void btnProcessar_Click(object sender, EventArgs e)
        {
            if (cmbBanco.SelectedIndex != 0)
            {
                if (!string.IsNullOrEmpty(txtFilial.Text))
                {
                    if (!string.IsNullOrEmpty(txtMatriz.Text))
                    {
                        if (dtIni.Value != DateTime.Today && dtFim.Value != DateTime.Today)
                        {
                            StreamWriter pod;
                            string caminho, path;

                            caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                            path = caminho + @"\RelatorioBanco.txt";
                            pod = File.CreateText(path);

                            int i = 0;
                            ECD.dao obj = new ECD.dao();
                            string sBanco = cmbBanco.SelectedItem.ToString();

                            estabeleceConexao();

                            //oDt = obj.executaComando();

                            //obj.ConectaBanco(txtLogin.Text, txtPassword.Text, sBanco, txtMatriz.Text, txtFilial.Text, dtIni.Value.ToString(), dtFim.Value.ToString());

                            foreach (DataRow linhas in oDt.Rows)
                            {
                                dados.Add("|" + linhas["chave_nf"].ToString() + "|" + linhas["id_item"].ToString() + "|");
                                dados.Add(dados.ToString());
                                i++;
                            }

                            pod.WriteLine("|ID_NF_ENTRADA|CHAVE_NF|NUM_NF|DT_EMISSAO|DT_ENTRADA|CGC_CPF|ID_ITEM|COD_PRODUTO|DESCRICAO_NOTA|COD_UNID_MEDIDA|QUANTIDADE|VLR_BRUTO|");

                            foreach (DataRow linha in oDt.Rows)
                            {
                                pod.WriteLine(linha);
                                pod.Flush();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Data inicial e final devem ser preenchidas", "Aviso", MessageBoxButtons.OK);
                            dtIni.Focus();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Matriz deve ser preenchida", "Aviso", MessageBoxButtons.OK);
                        txtMatriz.Focus();
                    }
                }
                else
                {
                    MessageBox.Show("Filial deve ser preenchido", "Aviso", MessageBoxButtons.OK);
                    txtFilial.Focus();
                }
            }
            else
            {
                MessageBox.Show("Deve ser selecionado um banco", "Aviso", MessageBoxButtons.OK);
                cmbBanco.Focus();
            }
        }

        private void cmbBanco_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbBanco.SelectedIndex != 0)
            {
                if (!string.IsNullOrEmpty(txtLogin.Text))
                {
                    dao.getStrConn(cmbBanco.SelectedIndex, txtLogin.Text, txtPassword.Text);
                }
                else
                {
                    MessageBox.Show("Login ou Senha não pode estar em branco!", "Aviso", MessageBoxButtons.OK);
                    txtLogin.Focus();
                    txtLogin.Clear();
                    txtPassword.Clear();
                    deuErro = true;
                    cmbBanco.SelectedIndex = 0;
                }
            }
            else
            {
                if (deuErro == false && sRetornoErro != "")
                {
                    MessageBox.Show("Deve ser selecionado um banco!", "Aviso", MessageBoxButtons.OK);
                    cmbBanco.Focus();
                    cmbBanco.SelectedIndex = 0;
                    deuErro = false;
                }
            }
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            Form1 frm = new Form1();
            this.Hide();
            frm.Show();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {


            this.ofd.Multiselect = false;
            this.ofd.Title = "Selecionar TNS";
            ofd.InitialDirectory = @"C:\";
            ofd.CheckFileExists = true;
            ofd.CheckPathExists = true;
            ofd.FilterIndex = 2;
            ofd.RestoreDirectory = true;
            ofd.ReadOnlyChecked = true;
            ofd.ShowReadOnly = true;

            DialogResult dr = this.ofd.ShowDialog();

            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                txtTns.Text = ofd.FileName;
                string sCaminhoTNS = txtTns.Text;
                var TNSNamesReader = new dao();
                cmbBanco.DataSource = TNSNamesReader.carregaCombo(sCaminhoTNS);
            }
        }

        private void estabeleceConexao()
        {
            sUser = txtLogin.Text.Trim();
            sPwd = txtPassword.Text.Trim();

            try
            {
                sVBanco = cmbBanco.SelectedItem.ToString();
                string sTeste = "";
                sTeste = cmbBanco.SelectedIndex.ToString();
                sRetornoErro = null;

                if (sVBanco == "")
                {
                    return;
                }

                string sql = "select ent.id_nf_entrada, ent.chave_nf, ent.num_nf, ent.dt_emissao," +
                " ent.dt_entrada, ent.cgc_cpf, itm.id_item, itm.cod_produto, itm.descricao_nota," +
                " itm.cod_unid_medida, itm.quantidade, itm.vlr_bruto" +
                "from lf_nf_entrada ent, lf_nf_entrada_item" +
                "where ent.cod_matriz = '" + txtMatriz.Text.Trim() + "'" +
                "and ent.cod_filial = '" + txtFilial.Text.Trim() + "'" +
                "and ent.dt_entrada between '" + dtIni.Value.ToString().Trim() + "' and '" + dtFim.Value.ToString().Trim() + "'" +
                "and.ent.cod_modelo = '55'" +
                "and ent.cod_status = '01'" +
                "and ent.cod_matriz = itm.cod_matriz" +
                "and ent.cod_filial = itm.cod_filial" +
                "and ent.id_nf_entrada = itm.id_nf_entrada;";

                dao.carregaBanco(cmbBanco, sql, sVBanco, sRetornoErro, sUser, sPwd);

                if (sRetornoErro == "")
                {
                    MessageBox.Show("Dados Carregados");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            { }
        }
    }
}