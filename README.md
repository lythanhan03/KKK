##C###
```c#
using ClosedXML.Excel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace KNSQL
{
    public partial class Form1 : Form
    {
        string connectstring = @"Data Source=DESKTOP-E9VSLH9; Initial Catalog=THWD; Integrated Security=True;";

        SqlConnection con;
        SqlCommand cmd;
        SqlDataAdapter adt;
        DataTable dt = new DataTable();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            con = new SqlConnection(connectstring);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                using (con = new SqlConnection(connectstring))
                {
                    con.Open();
                    using (cmd = new SqlCommand("DELETE FROM Giangvien WHERE Magv = @Magv", con))
                    {
                        cmd.Parameters.AddWithValue("@Magv", magiangvien.Text);
                        cmd.ExecuteNonQuery();
                    }
                }
                LoadData(); // Làm mới lưới dữ liệu sau khi xóa dữ liệu
            }
            catch (SqlException ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                using (con = new SqlConnection(connectstring))
                {
                    con.Open();
                    using (cmd = new SqlCommand("INSERT INTO Giangvien (Magv, Hovaten, Ngaysinh, Gioitinh, Makhoa) VALUES (@Magv, @Hovaten, @Ngaysinh, @Gioitinh, @Makhoa)", con))
                    {
                        cmd.Parameters.AddWithValue("@Magv", magiangvien.Text);
                        cmd.Parameters.AddWithValue("@Hovaten", tengiangvien.Text);
                        cmd.Parameters.AddWithValue("@Ngaysinh", ngaysinh.Text);
                        cmd.Parameters.AddWithValue("@Gioitinh", gioitinh.Text);
                        cmd.Parameters.AddWithValue("@Makhoa", khoa.Text);

                        cmd.ExecuteNonQuery();
                    }
                }
                LoadData(); // Làm mới lưới dữ liệu sau khi chèn dữ liệu
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void LoadData()
        {
            try
            {
                using (con = new SqlConnection(connectstring))
                {
                    con.Open();
                    using (cmd = new SqlCommand("SELECT * FROM Giangvien;", con))
                    {
                        using (adt = new SqlDataAdapter(cmd))
                        {
                            dt.Clear(); // Xóa dữ liệu cũ
                            adt.Fill(dt); // Điền dữ liệu mới
                            dataGridView2.DataSource = dt; // Thiết lập nguồn dữ liệu
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ten_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void xuatexcel_Click(object sender, EventArgs e)
        {
            try
            {
                using (SaveFileDialog sfd = new SaveFileDialog { Filter = "Excel Workbook|*.xlsx" })
                {
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        using (XLWorkbook wb = new XLWorkbook())
                        {
                            DataTable dt = ((DataTable)dataGridView2.DataSource);
                            wb.Worksheets.Add(dt, "Giangvien");
                            wb.SaveAs(sfd.FileName);
                        }
                        MessageBox.Show("Xuất dữ liệu ra Excel thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xảy ra: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void sua_Click(object sender, EventArgs e)
        {
            // Kiểm tra xem có dòng nào được chọn không
            if (dataGridView2.SelectedRows.Count == 1)
            {
                // Lấy dòng được chọn
                DataGridViewRow selectedRow = dataGridView2.SelectedRows[0];

                // Hiển thị dữ liệu của dòng được chọn lên các điều khiển nhập liệu
                magiangvien.Text = selectedRow.Cells["Magv"].Value.ToString();
                tengiangvien.Text = selectedRow.Cells["Hovaten"].Value.ToString();
                ngaysinh.Text = selectedRow.Cells["Ngaysinh"].Value.ToString();
                gioitinh.Text = selectedRow.Cells["Gioitinh"].Value.ToString();
                khoa.Text = selectedRow.Cells["Makhoa"].Value.ToString();
            }
            else
            {
                MessageBox.Show("Vui lòng chọn một dòng để sửa.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void luu_Click(object sender, EventArgs e)
        {
            try
            {
                using (con = new SqlConnection(connectstring))
                {
                    con.Open();
                    using (cmd = new SqlCommand("UPDATE Giangvien SET Hovaten = @Hovaten, Ngaysinh = @Ngaysinh, Gioitinh = @Gioitinh, Makhoa = @Makhoa WHERE Magv = @Magv", con))
                    {
                        cmd.Parameters.AddWithValue("@Magv", magiangvien.Text);
                        cmd.Parameters.AddWithValue("@Hovaten", tengiangvien.Text);
                        cmd.Parameters.AddWithValue("@Ngaysinh", ngaysinh.Text);
                        cmd.Parameters.AddWithValue("@Gioitinh", gioitinh.Text);
                        cmd.Parameters.AddWithValue("@Makhoa", khoa.Text);

                        cmd.ExecuteNonQuery();
                    }
                }
                MessageBox.Show("Cập nhật dữ liệu thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadData(); // Làm mới lưới dữ liệu sau khi cập nhật
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xảy ra: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
```
