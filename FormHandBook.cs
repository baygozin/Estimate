using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EstimatesAssembly
{
    public partial class FormHandBook : Form
    {
        private DataTable dt;

        public FormHandBook()
        {
            InitializeComponent();
        }

        private void FormHandBook_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "estimateDBDataSet.employees". При необходимости она может быть перемещена или удалена.
            this.employeesTableAdapter.Fill(this.estimateDBDataSet.employees);

        }

        private void bindingNavigatorSaveItems_Click(object sender, EventArgs e)
        {
            try
            {
                this.Validate();
                this.employeesBindingSource.EndEdit();
                this.employeesTableAdapter.Update(this.estimateDBDataSet.employees);
            } catch (System.Exception ex)
            {
                MessageBox.Show("Update failed");
            }
        }
//        try
//{
//    this.Validate();
//    this.customersBindingSource.EndEdit();
//    this.customersTableAdapter.Update(this.northwindDataSet.Customers);
//    MessageBox.Show("Update successful");
//}
//catch (System.Exception ex)
//{
//    MessageBox.Show("Update failed");
//}

    }
}
