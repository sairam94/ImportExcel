package helper;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import entities.Product;

public class ExcelHelper {
	
	String excelPath = null;
	
/*	public static void main(String arg[]){
		System.out.println("Main");
		ExcelHelper eh = new ExcelHelper();
		try {
			eh.readXLSXFile();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}*/
	public ExcelHelper(String excelPath) {
		super();
		this.excelPath = excelPath;
	}

	public List<Product> readXLSXFile() throws IOException
	{
		List<Product> listProduct = new ArrayList<Product>();
		System.out.println("readXLSXFile");
		InputStream ExcelFileToRead = new FileInputStream(excelPath);
		XSSFWorkbook  wb = new XSSFWorkbook(ExcelFileToRead);
		
		XSSFWorkbook test = new XSSFWorkbook(); 
		
		XSSFSheet sheet = wb.getSheetAt(0);
		XSSFRow row; 
		XSSFCell cell;

		Iterator rows = sheet.rowIterator();
		while (rows.hasNext())
		{
			row=(XSSFRow) rows.next();
			Product product = new Product();
			Iterator cells = row.cellIterator();
			int flag = 0;
			while (cells.hasNext())
			{	
				 
				cell=(XSSFCell) cells.next();
				if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING && flag==0)
				{	
					product.setId(cell.getStringCellValue());
					flag++;
				}
				else if(cell.getCellType() == XSSFCell.CELL_TYPE_STRING && flag==1)
				{
					product.setName(cell.getStringCellValue());
					flag++;
				}
				else if(cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC && flag==2)
				{
					product.setPrice(cell.getNumericCellValue());
					flag++;
				}
				else if(cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC && flag==3)
				{
					product.setQuantity(cell.getNumericCellValue());
					flag++;
				}
				else if(cell.getCellType() == XSSFCell.CELL_TYPE_BOOLEAN && flag==4)
				{
					product.setStatus(cell.getBooleanCellValue());
					flag++;
				}
				else if(cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC && flag==5)
				{
					product.setCreationDate(cell.getDateCellValue());
					flag++;
				}
				else
				{
					//U Can Handel Boolean, Formula, Errors
				}
			}
			listProduct.add(product);
			System.out.println("End");
		}
		
		return listProduct;
	}
	
}
