using System;
using System.Data;
using System.Data.SqlServerCe;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Configuration;
using System.ComponentModel;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace ManipulacaoBanco
{
    public partial class frmPrincipal : Form
    {
        public frmPrincipal()
        {
            InitializeComponent();

            //progressBar1.Visible = false;
            //lblProgress.Visible = false;

            #region CUSTOMIZAÇÃO DO DATAGRIDVIEW

            // linhas alternadas
            dataBanco.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(234, 234, 234);
            dataImport.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(234, 234, 234);
            dataExport.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(234, 234, 234);

            // linha selecionada
            dataBanco.DefaultCellStyle.SelectionBackColor = Color.FromArgb(230, 125, 33);
            dataExport.DefaultCellStyle.SelectionBackColor = Color.FromArgb(230, 125, 33);
            dataImport.DefaultCellStyle.SelectionBackColor = Color.FromArgb(230, 125, 33);
            dataBanco.DefaultCellStyle.SelectionForeColor = Color.White;
            dataImport.DefaultCellStyle.SelectionForeColor = Color.White;
            dataExport.DefaultCellStyle.SelectionForeColor = Color.White;

            // fonte
            //dataGridView2.DefaultCellStyle.Font = new Font("Century Gothic",8);

            // bordas
            dataBanco.CellBorderStyle = DataGridViewCellBorderStyle.None;
            dataImport.CellBorderStyle = DataGridViewCellBorderStyle.None;
            dataExport.CellBorderStyle = DataGridViewCellBorderStyle.None;

            // cabeçalho
            dataBanco.ColumnHeadersDefaultCellStyle.Font = new Font("Century Gothic", 8, FontStyle.Bold);
            dataImport.ColumnHeadersDefaultCellStyle.Font = new Font("Century Gothic", 8, FontStyle.Bold);
            dataExport.ColumnHeadersDefaultCellStyle.Font = new Font("Century Gothic", 8, FontStyle.Bold);

            dataBanco.EnableHeadersVisualStyles = false; // habilita a edição do cabeçalho
            dataImport.EnableHeadersVisualStyles = false; // habilita a edição do cabeçalho
            dataExport.EnableHeadersVisualStyles = false; // habilita a edição do cabeçalho

            dataBanco.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(211, 84, 21);
            dataImport.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(211, 84, 21);
            dataExport.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(211, 84, 21);
            dataBanco.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataImport.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataExport.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            #endregion

            CarregarCaminhoBase();

        }

        public void CarregarBanco()
        {
            string caminhoBanco = txtCaminhoBanco.Text;
            CarregarArquivoImportado(caminhoBanco, dataBanco);
        }

        private void btnPastaImport_Click(object sender, EventArgs e)
        {

            try
            {
                OpenFileDialog file = new OpenFileDialog();
                file.Filter = "Arquivo (*.CSV)|*.CSV";
                if (file.ShowDialog() == DialogResult.OK)
                {
                    #region CARREGAR DATAGRID

                    CarregarArquivoImportado(file.FileName, dataImport);
                    txtPastaImport.Text = file.FileName;

                    #endregion

                    #region REMOVER E INSERIR PEDIDOS

                    RemoverPedidos();
                    lblBancoDeDados.Text = "Registros duplicados removidos.";
                    InserirPedidos(dataExport, dataBanco);
                    InserirPedidos(dataExport, dataImport);
                    lblImport.Text = "CSV carregado.";

                    #endregion

                    #region MODIFICAR ARQUIVO

                    new Exportador().ExportarCSV(dataExport, txtCaminhoBanco.Text);
                    MessageBox.Show(@"Banco de dados modificado com sucesso." + "\n\n" + txtCaminhoBanco.Text, "", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    #endregion
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro:" + "\n\n" + ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void CarregarArquivoImportado(string ArquivoCSV, DataGridView data)
        {
            DataTable dt = new DataTable();

            try // verificar se o arquivo existe ou está aberto
            {
                string[] linhas = System.IO.File.ReadAllLines(ArquivoCSV);
                if (linhas.Length > 0)
                {
                    // para o cabeçalho
                    string primeiralinha = linhas[0];
                    string[] cabecalho = primeiralinha.Split(';');
                    foreach (string celulaCabecalho in cabecalho)
                    {
                        dt.Columns.Add(new DataColumn(celulaCabecalho));
                    }

                    //progressBar1.BeginInvoke(new Action(() =>
                    //{
                    //    progressBar1.Maximum = linhas.Length;
                    //}
                    //));
                    
                    // para as celulas
                    for (int r = 1; r < linhas.Length; r++) //1
                    {
                        //backgroundWorker1.ReportProgress(r);

                        string[] celulas = linhas[r].Split(';');
                        DataRow dr = dt.NewRow();
                        int indice = 0;

                        foreach (string celula in cabecalho)
                        {
                            try
                            {
                                dr[celula] = celulas[indice++];
                            }
                            catch (Exception)
                            {
                                MessageBox.Show($"Erro na linha {r} e coluna {indice - 1} do banco de dados. \nNão foi possível carregar o arquivo.");
                                return; // encerra o metodo
                            }
                        }
                        if (dr[2].ToString() != "")
                        {
                            dt.Rows.Add(dr);
                        }
                    }
                }

                if (dt.Rows.Count > 0)
                {
                    //data.BeginInvoke(new Action(() =>
                    //{
                    //    data.DataSource = dt;
                    //}
                    //));
                    data.DataSource = dt;
                }

                lblBancoDeDados.Text = "Banco de dados carregado.";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro: " + ex.Message);
            }
        }

        public void RemoverPedidos()
        {
            foreach (DataGridViewRow linhaImport in dataImport.Rows)
            {
                string pedidoImport = linhaImport.Cells[2].Value.ToString();

                foreach (DataGridViewRow linhaBanco in dataBanco.Rows)
                {
                    string pedidoBanco = linhaBanco.Cells[2].Value.ToString();
                    if (pedidoBanco == pedidoImport)
                    {
                        dataBanco.Rows.Remove(linhaBanco);
                    }
                }
            }
        }

        public void InserirPedidos(DataGridView dataRecebe, DataGridView dataTransfere)
        {
            dataRecebe.AllowUserToAddRows = true;
            DataGridViewRow row = new DataGridViewRow();

            for (int i = 0; i < dataTransfere.Rows.Count; i++)
            {
                row = (DataGridViewRow)dataTransfere.Rows[i].Clone();
                int intColIndex = 0;
                foreach (DataGridViewCell cell in dataTransfere.Rows[i].Cells)
                {
                    row.Cells[intColIndex].Value = cell.Value;
                    intColIndex++;
                }
                dataRecebe.Rows.Add(row);
            }
            dataRecebe.AllowUserToAddRows = false;
            dataRecebe.Refresh();
        }

        private void btnAtualizarBanco_Click(object sender, EventArgs e)
        {
            AlterarCaminhoBase();
            CarregarBanco();
            MessageBox.Show("Banco atualizado.");
        }

        #region USADO SÓ DA PRIMEIRA VEZ

        public void CriarBase()
        {
            string baseDados = @"\\paris\eng\Usuarios\Lorenzo\BancoCaminho.sdf";
            string strConnection = @"DataSource = " + baseDados + ";Password = '1234'";
            SqlCeEngine db = new SqlCeEngine(strConnection);
            if (!File.Exists(baseDados))
            {
                db.CreateDatabase();
            }
            db.Dispose();
            SqlCeConnection conexao = new SqlCeConnection();
            conexao.ConnectionString = strConnection;
            try
            {
                conexao.Open();
                SqlCeCommand comando = new SqlCeCommand();
                comando.Connection = conexao;

                comando.CommandText = "CREATE TABLE tabelaCaminho (id INT NOT NULL PRIMARY KEY, caminho NVARCHAR(100))";
                comando.ExecuteNonQuery();

                //label1.Text = "Tabela criada.";
            }
            catch (Exception ex)
            {
                //label1.Text = ex.Message;
            }
            finally
            {
                conexao.Close();
            }
        }

        public void InserirCaminhoBase()
        {
            string baseDados = @"\\paris\eng\Usuarios\Lorenzo\BancoCaminho.sdf";
            string strConection = @"DataSource = " + baseDados + "; Password = '1234'";

            SqlCeConnection conexao = new SqlCeConnection(strConection);

            try
            {
                conexao.Open();

                SqlCeCommand comando = new SqlCeCommand();
                comando.Connection = conexao;

                int id = 0;
                string caminho = @"\\paris\eng\Usuarios\Lorenzo\banco.csv";

                comando.CommandText = "INSERT INTO tabelaCaminho VALUES (" + id + ", '" + caminho + "')";

                comando.ExecuteNonQuery();

                lblBancoDeDados.Text = "Registro inserido.";
                comando.Dispose();
            }
            catch (Exception ex)
            {
                lblBancoDeDados.Text = ex.Message;
            }
            finally
            {
                conexao.Close();
            }
        }

        #endregion

        public void AlterarCaminhoBase()
        {
            string baseDados = @"\\paris\eng\Usuarios\Lorenzo\BancoCaminho.sdf";
            string strConection = @"DataSource = " + baseDados + "; Password = '1234'";

            SqlCeConnection conexao = new SqlCeConnection(strConection);

            try
            {
                conexao.Open();

                SqlCeCommand comando = new SqlCeCommand();
                comando.Connection = conexao;

                int id = 0;
                string caminho = txtCaminhoBanco.Text;

                string query = "UPDATE tabelaCaminho SET caminho = '" + caminho + "' WHERE id LIKE '" + id + "'";

                comando.CommandText = query;

                comando.ExecuteNonQuery();

                lblBancoDeDados.Text = "Caminho alterado";
                comando.Dispose();
            }
            catch (Exception ex)
            {
                lblBancoDeDados.Text = ex.Message;
            }
            finally
            {
                conexao.Close();
            }
        }

        public void CarregarCaminhoBase()
        {
            string baseDados = @"\\paris\eng\Usuarios\Lorenzo\BancoCaminho.sdf";
            string strConnection = @"DataSource = " + baseDados + ";Password = '1234'";

            SqlCeConnection conexao = new SqlCeConnection(strConnection);
            conexao.Open();

            try
            {
                int id = 0;
                string query = "SELECT caminho FROM tabelaCaminho WHERE id = '" + id + "'  ";

                SqlCeCommand comando = new SqlCeCommand(query, conexao);

                string caminho = comando.ExecuteScalar().ToString();

                txtCaminhoBanco.Text = caminho;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conexao.Close();
            }
        }

        private void btnCarregarBanco_Click(object sender, EventArgs e)
        {
            //backgroundWorker1.WorkerReportsProgress = true;
            //backgroundWorker1.RunWorkerAsync();

            CarregarBanco();
        }

        //PROGRESS BAR

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            CarregarBanco();
        }

        private void backgroundWorker1_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            //progressBar1.Value = e.ProgressPercentage;
            //lblProgress.Text = e.ProgressPercentage.ToString() + "%";
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            //progressBar1.Value =0;
            //lblProgress.Text = "100%";
        }
    }
}
