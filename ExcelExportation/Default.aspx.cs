using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ExcelExportation
{
    public partial class Default : System.Web.UI.Page
    {

        Excel importedExcel;
        Excel excel = new Excel();

        private int rows;

        private string fileName = "";
        private string msgStr = "";
        private string errorStr = "";

        protected void Page_Load(object sender, EventArgs e)
        {

        }


        public void createHeader()
        {
            excel.CreateNewFile();

            excel.WriteToCell(0, 0, "Column 1");
            excel.WriteToCell(0, 1, "Column 2");
            excel.WriteToCell(0, 2, "Column 3");
            excel.WriteToCell(0, 3, "Column 4");
            excel.WriteToCell(0, 4, "Column 5");
            excel.WriteToCell(0, 5, "Column 6");
        }

        public void writeInfo()
        {
            fileName = FileUpload1.FileName;
            copyTempFile();
            importedExcel = new Excel(Server.MapPath("tempFile\\") + fileName, 1);

            rows = importedExcel.ws.UsedRange.Rows.Count;

            /// Copy in bulk 
            /// copy from row 2 to row 10, col 1 to col 4 of the imported excel
            /// paste to row 2 to row 10, col 1 to col 4 of the exported excel file
            ///
            ///string[,] readData = importedExcel.ReadRange(2, 1, rows, 4);
            ///excel.WriteRange(2, 3, rows,6 , readData);

            for (int i = 1; i < rows; i++)
            {
                /// Copy col 1 to col 6 from the imported excel file to exporting excel file
                /// Can add any formula to edit the imported value before writing it in the exporting file
                excel.WriteToCell(i, 0, importedExcel.ws.Cells[i + 1 , 1].Value.ToString());
                excel.WriteToCell(i, 1, importedExcel.ws.Cells[i + 1, 2].Value.ToString());
                excel.WriteToCell(i, 2, importedExcel.ws.Cells[i + 1, 3].Value.ToString());
                excel.WriteToCell(i, 3, importedExcel.ws.Cells[i + 1, 4].Value.ToString());
                excel.WriteToCell(i, 4, importedExcel.ws.Cells[i + 1, 5].Value.ToString());
                excel.WriteToCell(i, 5, importedExcel.ws.Cells[i + 1, 6].Value.ToString());
            }
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            createHeader();
            writeInfo();
            excel.SaveAs(Server.MapPath("Export\\") + "Export File - " + TextBox1.Text);

            importedExcel.Close();
            deleteTempFile();
            excel.Close();

            Response.Write("<script>alert('File Exported!');</script>");
        }

        public void copyTempFile()
        {
            if (File.Exists(Server.MapPath("tempFile\\") + fileName))
            {
                File.Delete(Server.MapPath("tempFile\\") + fileName);
            }
            FileUpload1.PostedFile.SaveAs(Server.MapPath("tempFile\\") + fileName);
        }

        public void deleteTempFile()
        {
            if (File.Exists(Server.MapPath("tempFile\\") + fileName))
            {
                File.Delete(Server.MapPath("tempFile\\") + fileName);
            }
        }
    }
}