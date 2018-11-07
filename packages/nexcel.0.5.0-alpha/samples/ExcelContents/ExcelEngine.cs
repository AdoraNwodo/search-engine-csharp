using System;
using NExcel;

namespace ExcelContents
{
	/// <summary>
	/// Reads Excel files.
	/// Writes data to console. 
	/// </summary>
	public class ExcelEngine
	{

		/// <summary>
		/// Writes to console the contents of a Excel file.
		/// Writes all sheets and all cells in sheet.
		/// Otherwise does nothing.
		/// </summary>
		/// <param name="filename">the Excel file name</param>
		public void Write(string filename)
		{
			// check the input parameters
			if (filename==null) return;


			// init
			Workbook workbook = null;

			try
			{
				// open the Excel workbook
				workbook = Workbook.getWorkbook(filename);

				// for each sheet in workbook, write cell contents to console
				foreach (Sheet sheet in workbook.Sheets)
				{
					// for each row
					for (int irow = 0; irow < sheet.Rows; irow++)
					{
						// for each column
						for(int icol=0; icol<sheet.Columns; icol++)
						{
							// get current cell
							Cell cell = sheet.getCell(icol, irow);

							// write to console the cell contents
							System.Console.WriteLine("{0}[{1},{2}]:  {3}", sheet.Name, irow, icol, cell.Contents);
						}
					} 
				}

			}
			catch (Exception ex) 
			{
				Console.WriteLine(ex.Message);
			}
			finally
			{
				// Close workbook
				if ((workbook != null))
				{
					workbook.close();
				}
			}

		}


	}
}
