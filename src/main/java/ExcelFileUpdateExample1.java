import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.hssf.usermodel.HSSFWorkbook; //para crear un excel en blanco

/**
 * This program illustrates how to update an existing Microsoft Excel document.
 * Append new rows to an existing sheet.
 * 
 * @author www.codejava.net
 *
 */
public class ExcelFileUpdateExample1 {


	public static void main(String[] args) {
		String excelFilePath = "Inventario.xlsx";
		FileInputStream inputStream;

		try {
			File file = new File(excelFilePath);   // se crea una variable del archivo
			Workbook workbook;

			Sheet sheet;
			
			if (file.exists()){  //si el archivo existe se obtienen los datos de alli, y si no se crea uno nuevo
				inputStream = new FileInputStream(file);
				workbook = WorkbookFactory.create(inputStream);
				inputStream.close();
				sheet = workbook.getSheetAt(0);
			}
			else{
				Object[] titulos = {"No","Book Title","Author","Price"};
				workbook = new HSSFWorkbook();//uso el "hssfworbook" por que me permite crear un archivo desde 0 muy facilmente
				sheet = workbook.createSheet("Java Books 1");// se crea una nueva hoja 
				Row row = sheet.createRow(0);
				Cell cell;
				int columnCount = 0;
				for (Object titulo : titulos){
					cell = row.createCell(columnCount);
					cell.setCellValue((String) titulo);
					columnCount++;
				}
				
			}
			
			
			Object[][] bookData = {
					{"El que se duerme pierde", "Tom Peter", 16},
					{"Sin lugar a duda", "Ana Gutierrez", 26},
					{"El arte de dormir", "Nico", 32},
					{"Buscando a Nemo", "Humble Po", 41},
			};

			int rowCount = sheet.getLastRowNum();

			for (Object[] aBook : bookData) {
				Row row = sheet.createRow(++rowCount);

				int columnCount = 0;
				
				Cell cell = row.createCell(columnCount);
				cell.setCellValue(rowCount);
				
				for (Object field : aBook) {
					cell = row.createCell(++columnCount);
					if (field instanceof String) {
						cell.setCellValue((String) field);
					} else if (field instanceof Integer) {
						cell.setCellValue((Integer) field);
					}
				}

			}

			Iterator<Sheet> sheets = workbook.iterator();

			while (sheets.hasNext()){
				Sheet nextSheet = sheets.next();
				System.out.println("***************" + nextSheet.getSheetName() + "***************");
				
				Iterator<Row> iterator = nextSheet.iterator(); //creamos el iterador sobre la hoja

				while (iterator.hasNext()) { 			  //como 3 ciclos anidados
					Row nextRow = iterator.next();
					Iterator<Cell> cellIterator = nextRow.cellIterator(); //y tambien creamos el iterador sobre cada fila de la hoja
					
					while (cellIterator.hasNext()) { 
						Cell cell = cellIterator.next();
        		    	switch (cell.getCellType()) {   //dependiendo del tipo de datos de la celda hacemos el print correspondiente
							case Cell.CELL_TYPE_STRING:
							System.out.print(cell.getStringCellValue());
							break;
							case Cell.CELL_TYPE_BOOLEAN:
							System.out.print(cell.getBooleanCellValue());
							break;
							case Cell.CELL_TYPE_NUMERIC:
							System.out.print(cell.getNumericCellValue());
							break;
						}
						System.out.print(" - ");
					}
					System.out.println();
				}
			}
			System.out.println("*********************************");

			FileOutputStream outputStream = new FileOutputStream(excelFilePath);
			workbook.write(outputStream);
			workbook.close();
			outputStream.close();
			
		} catch (IOException | EncryptedDocumentException
				| InvalidFormatException ex) {
			ex.printStackTrace();
		}
	}

}
