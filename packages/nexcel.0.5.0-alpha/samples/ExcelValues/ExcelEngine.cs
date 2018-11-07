using System;
using NExcel;

namespace ExcelValues
{
	/// <summary>
	/// Reads Excel files.
	/// Writes data to console. 
	/// </summary>
	public class ExcelEngine
	{

		/// <summary>
		/// Writes to console the values of a Excel file.
		/// Writes all sheets and all cells in sheet.
		/// For each cell writes value and value type.
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

				// for each sheet in workbook, write cell values to console
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

							// get value and value type
							object val = cell.Value;
							string strValue =  (val!=null)  ? val.ToString() : "";
							string strValueType =  (val!=null)  ? val.GetType().ToString() : "empty cell";

							// write to console the cell value type and the value 
							System.Console.WriteLine("{0}[{1},{2}]: ({3})  {4}", 
								sheet.Name, irow, icol, strValueType, strValue);
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
