using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExamenResuelto
{
    public partial class Form1 : Form
    {
        OleDbConnection miCnx;
        OleDbCommand miCmd;
        OleDbDataAdapter sAdapter;
        OleDbCommandBuilder sBuilder;
        DataSet sDs;
        DataTable sTable;


        public Form1()
        {
            InitializeComponent();
        }

        void eliminarpaneles()
        {
            panel1.Visible = true;
            panelBlanco.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;
            panel5.Visible = false;
            panel6.Visible = false;
            panel7.Visible = false;
            panel8.Visible = false;
            panel9.Visible = false;
            panel10.Visible = false;
        }

        void cargarCentros()
        {
            string connetionString = null;
            OleDbConnection connection;
            OleDbCommand command;
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            DataSet ds = new DataSet();
            string sql = null;
            connetionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=|DataDirectory|\\Database1.mdb";

            sql = "SELECT Id_Centro, Centro FROM Centros";
            connection = new OleDbConnection(connetionString);
            try
            {
                connection.Open();
                command = new OleDbCommand(sql, connection);
                adapter.SelectCommand = command;
                adapter.Fill(ds);
                adapter.Dispose();
                command.Dispose();
                comboBox1.DataSource = ds.Tables[0];
                comboBox1.ValueMember = "Id_Centro";
                comboBox1.DisplayMember = "Centro";
                comboBox2.DataSource = ds.Tables[0];
                comboBox2.ValueMember = "Id_Centro";
                comboBox2.DisplayMember = "Centro";
                comboBox3.DataSource = ds.Tables[0];
                comboBox3.ValueMember = "Id_Centro";
                comboBox3.DisplayMember = "Centro";
                comboBox6.DataSource = ds.Tables[0];
                comboBox6.ValueMember = "Id_Centro";
                comboBox6.DisplayMember = "Centro";
                listBox1.DataSource = ds.Tables[0];
                listBox1.ValueMember = "Id_Centro";
                listBox1.DisplayMember = "Centro";
                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can not open connection ! ");
            }

        }

        //Centros

        //Inicio
        private void inicioToolStripMenuItem_Click(object sender, EventArgs e)
        {
            eliminarpaneles();
        }

        //Añadir
        private void añadirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            eliminarpaneles();
            panel2.Visible = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string conexion = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=|DataDirectory|\\Database1.mdb";
            miCnx = new OleDbConnection(conexion);
            string sql = "SELECT COUNT(*) FROM Centros WHERE Centro = '" + textBox1.Text + "'";
            miCmd = new OleDbCommand(sql, miCnx);
            miCnx.Open();
            int cantidad = (int)miCmd.ExecuteScalar();
            
            if (cantidad == 0)
            {
                string sql2 = "INSERT INTO Centros(Centro) VALUES('" + textBox1.Text+"');";
                miCmd = new OleDbCommand(sql2, miCnx);
                miCmd.ExecuteNonQuery();
            }

            miCnx.Close();
        }

        //Listar
        private void listarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            eliminarpaneles();
            panel3.Visible = true;

            string conexion = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=|DataDirectory|\\Database1.mdb";

            string sql = "SELECT * FROM Centros";
            OleDbConnection connection = new OleDbConnection(conexion);
            connection.Open();
            miCmd = new OleDbCommand(sql, connection);
            sAdapter = new OleDbDataAdapter(miCmd);
            sBuilder = new OleDbCommandBuilder(sAdapter);
            sDs = new DataSet();
            sAdapter.Fill(sDs, "Centros");
            sTable = sDs.Tables["Centros"];
            connection.Close();
            dataGridView1.DataSource = sDs.Tables["Centros"];
            dataGridView1.ReadOnly = true;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        //Borrar
        private void borrarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            eliminarpaneles();
            cargarCentros();
            panel4.Visible = true;


            string conexion = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=|DataDirectory|\\Database1.mdb";
            string sentencia = "DELETE FROM Centros WHERE Centro='" + comboBox1.Text + "'";
            OleDbConnection miCnx = new OleDbConnection(conexion);
            OleDbCommand miCmd = new OleDbCommand(sentencia, miCnx);
            miCnx.Open();
            miCmd.ExecuteNonQuery();
            miCnx.Close();
            miCnx.Dispose();
        }

        //Departamentos
        
        //Inicio
        private void inicioToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            eliminarpaneles();
        }

        //Añadir
        private void añadirToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            eliminarpaneles();
            cargarCentros();
            panel5.Visible = true;
            


        }

        private void button3_Click(object sender, EventArgs e)
        {
            string conexion = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=|DataDirectory|\\Database1.mdb";
            miCnx = new OleDbConnection(conexion);
            string sql = "SELECT COUNT(*) FROM Departamentos WHERE Departamento = '" + textBox2.Text + "'";
            miCmd = new OleDbCommand(sql, miCnx);
            miCnx.Open();
            int cantidad = (int)miCmd.ExecuteScalar();

            if (cantidad == 0)
            {
                string sql2 = "INSERT INTO Departamentos(Id_Centro,Departamento) VALUES(" + comboBox2.SelectedValue + ", '" + textBox2.Text + "');";
                miCmd = new OleDbCommand(sql2, miCnx);
                miCmd.ExecuteNonQuery();
            }

            miCnx.Close();
        }

        //Listar
        private void listarToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            eliminarpaneles();
            panel6.Visible = true;

            string conexion = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=|DataDirectory|\\Database1.mdb";

            string sql = "SELECT A.Id_Departamento, A.Departamento, B.Centro FROM Departamentos A, Centros B WHERE A.Id_Centro = B.Id_Centro";
            OleDbConnection connection = new OleDbConnection(conexion);
            connection.Open();
            miCmd = new OleDbCommand(sql, connection);
            sAdapter = new OleDbDataAdapter(miCmd);
            sBuilder = new OleDbCommandBuilder(sAdapter);
            sDs = new DataSet();
            sAdapter.Fill(sDs, "Departamentos");
            sTable = sDs.Tables["Departamentos"];
            connection.Close();
            dataGridView2.DataSource = sDs.Tables["Departamentos"];
            dataGridView2.ReadOnly = true;
            dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        //Borrar
        private void borrarToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            eliminarpaneles();
            cargarCentros();
            panel7.Visible = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string connetionString = null;
            OleDbConnection connection;
            OleDbCommand command;
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            DataSet ds = new DataSet();
            string sql = null;
            connetionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=|DataDirectory|\\Database1.mdb";

            sql = "SELECT Id_Departamento, Departamento FROM Departamentos WHERE Id_Centro = " + comboBox3.SelectedValue;
            connection = new OleDbConnection(connetionString);
            try
            {
                connection.Open();
                command = new OleDbCommand(sql, connection);
                adapter.SelectCommand = command;
                adapter.Fill(ds);
                adapter.Dispose();
                command.Dispose();
                comboBox4.DataSource = ds.Tables[0];
                comboBox4.ValueMember = "Id_Departamento";
                comboBox4.DisplayMember = "Departamento";
                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can not open connection ! ");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string conexion = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=|DataDirectory|\\Database1.mdb";
            string sentencia = "DELETE FROM Departamentos WHERE Departamento='" + comboBox4.Text + "'";
            OleDbConnection miCnx = new OleDbConnection(conexion);
            OleDbCommand miCmd = new OleDbCommand(sentencia, miCnx);
            miCnx.Open();
            miCmd.ExecuteNonQuery();
            miCnx.Close();
            miCnx.Dispose();
        }

        //Empleados

        //Inicio
        private void inicioToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            eliminarpaneles();
        }

        //Añadir
        private void añadirToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            eliminarpaneles();
            cargarCentros();
            panel8.Visible = true;



        }

        private void button6_Click(object sender, EventArgs e)
        {
            string connetionString = null;
            OleDbConnection connection;
            OleDbCommand command;
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            DataSet ds = new DataSet();
            string sql = null;
            connetionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=|DataDirectory|\\Database1.mdb";

            sql = "SELECT Id_Departamento, Departamento FROM Departamentos WHERE Id_Centro = " + comboBox3.SelectedValue;
            connection = new OleDbConnection(connetionString);
            try
            {
                connection.Open();
                command = new OleDbCommand(sql, connection);
                adapter.SelectCommand = command;
                adapter.Fill(ds);
                adapter.Dispose();
                command.Dispose();
                comboBox5.DataSource = ds.Tables[0];
                comboBox5.ValueMember = "Id_Departamento";
                comboBox5.DisplayMember = "Departamento";
                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can not open connection ! ");
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string conexion = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=|DataDirectory|\\Database1.mdb";
            miCnx = new OleDbConnection(conexion);
            string sql = "SELECT COUNT(*) FROM Empleados WHERE Empleado = '" + textBox3.Text + "'";
            miCmd = new OleDbCommand(sql, miCnx);
            miCnx.Open();
            int cantidad = (int)miCmd.ExecuteScalar();

            if (cantidad == 0)
            {
                string sql2 = "INSERT INTO Empleados(Id_Departamento,Empleado) VALUES(" + comboBox5.SelectedValue + ", '" + textBox3.Text + "');";
                miCmd = new OleDbCommand(sql2, miCnx);
                miCmd.ExecuteNonQuery();
            }

            miCnx.Close();
        }

        //Listar
        private void listarToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            eliminarpaneles();
            panel9.Visible = true;

            string conexion = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=|DataDirectory|\\Database1.mdb";

            string sql = "SELECT A.Empleado, B.Departamento, C.Centro FROM Empleados A, Departamentos B, Centros C WHERE A.Id_Departamento = B.Id_Departamento AND B.Id_Centro = C.Id_Centro";
            OleDbConnection connection = new OleDbConnection(conexion);
            connection.Open();
            miCmd = new OleDbCommand(sql, connection);
            sAdapter = new OleDbDataAdapter(miCmd);
            sBuilder = new OleDbCommandBuilder(sAdapter);
            sDs = new DataSet();
            sAdapter.Fill(sDs, "Empleados");
            sTable = sDs.Tables["Empleados"];
            connection.Close();
            dataGridView3.DataSource = sDs.Tables["Empleados"];
            dataGridView3.ReadOnly = true;
            dataGridView3.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

        }

        //Borrar
        private void borrarToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            eliminarpaneles();
            cargarCentros();
            panel10.Visible = true;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string connetionString = null;
            OleDbConnection connection;
            OleDbCommand command;
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            DataSet ds = new DataSet();
            string sql = null;
            connetionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=|DataDirectory|\\Database1.mdb";

            sql = "SELECT Id_Departamento, Departamento FROM Departamentos WHERE Id_Centro = " + comboBox6.SelectedValue;
            connection = new OleDbConnection(connetionString);
            try
            {
                connection.Open();
                command = new OleDbCommand(sql, connection);
                adapter.SelectCommand = command;
                adapter.Fill(ds);
                adapter.Dispose();
                command.Dispose();
                comboBox7.DataSource = ds.Tables[0];
                comboBox7.ValueMember = "Id_Departamento";
                comboBox7.DisplayMember = "Departamento";
                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can not open connection ! ");
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string connetionString = null;
            OleDbConnection connection;
            OleDbCommand command;
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            DataSet ds = new DataSet();
            string sql = null;
            connetionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=|DataDirectory|\\Database1.mdb";

            sql = "SELECT Id_Empleado, Empleado FROM Empleados WHERE Id_Departamento = " + comboBox7.SelectedValue;
            connection = new OleDbConnection(connetionString);
            try
            {
                connection.Open();
                command = new OleDbCommand(sql, connection);
                adapter.SelectCommand = command;
                adapter.Fill(ds);
                adapter.Dispose();
                command.Dispose();
                listBox2.DataSource = ds.Tables[0];
                listBox2.ValueMember = "Id_Empleado";
                listBox2.DisplayMember = "Empleado";
                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can not open connection ! ");
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            string conexion = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=|DataDirectory|\\Database1.mdb";
            string sentencia = "DELETE FROM Empleados WHERE Empleado='" + listBox2.Text + "'";
            OleDbConnection miCnx = new OleDbConnection(conexion);
            OleDbCommand miCmd = new OleDbCommand(sentencia, miCnx);
            miCnx.Open();
            miCmd.ExecuteNonQuery();
            miCnx.Close();
            miCnx.Dispose();
        }
    }
}
