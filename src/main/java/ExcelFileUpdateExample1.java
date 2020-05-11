import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Scanner;

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
				sheet = workbook.createSheet("Java Books");// se crea una nueva hoja 
				Row row = sheet.createRow(0);
				Cell cell;
				int columnCount = 0;
				for (Object titulo : titulos){
					cell = row.createCell(columnCount);
					cell.setCellValue((String) titulo);
					columnCount++;
				}
				
			}

			if (sheet.getLastRowNum() == 30)
			{

				String name_sheet = "Java Books "+String.valueOf(workbook.getActiveSheetIndex()+1); //Guardo el nombre de la nueva hoja con el índice X

				sheet = workbook.createSheet(name_sheet);	//Creo la nueva hoja con el nombre guardado.

				workbook.setActiveSheet(workbook.getActiveSheetIndex()+1);	//Cambio la hoja activa a la que acabo de crear

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

			//Menu de actualización de un registro específico del archivo según su Nro. de Identificación.

			Scanner scanner = new Scanner(System.in);
			boolean mainLoop = true;
			int choice = -1;
			int id = 0;
			while (mainLoop) {
				System.out.println("¿Desea realizar alguna modificación? (Escoja el número de su decisión).");
				System.out.println("*********************************");
				System.out.println("1. Modificar Autor.");
				System.out.println("2. Modificar Precio.");
				System.out.println("0. Salir.");

				choice = scanner.nextInt();

				switch (choice){
					case 0: //Salir
						mainLoop = false;
						break;
					case 1:
						Scanner input = new Scanner(System.in);

						System.out.println(" Indique el id del registro a actualizar \n");
						id = input.nextInt();
						updateFile(id,"Author", excelFilePath);
						break;
					case 2:
						Scanner input2 = new Scanner(System.in);

						System.out.println(" Indique el id del registro a actualizar \n");
						id = input2.nextInt();
						updateFile(id,"Price", excelFilePath);
						break;
				}
			}
			//Se cierra el archivo.

			FileOutputStream outputStream = new FileOutputStream(excelFilePath);
			workbook.write(outputStream);
			workbook.close();
			outputStream.close();

		} catch (IOException | EncryptedDocumentException
				| InvalidFormatException ex) {
			ex.printStackTrace();
		}
	}

	static void updateFile(int id, String attribute, String excelFile) {
		String value="";
		int price=0, rowCount=0;
		boolean loop = true;
		Scanner input = new Scanner(System.in);
		Scanner input2 = new Scanner(System.in);

		try {
			FileInputStream inputStream = new FileInputStream(new File(excelFile));
			Workbook workbook = WorkbookFactory.create(inputStream);
			Sheet sheet = workbook.getSheetAt(0);
			rowCount = sheet.getLastRowNum();
		}catch(Exception e){
			e.printStackTrace();
		}
		System.out.println("valor: "+excelFile);
		if (attribute == "Author" && id <= rowCount) {
			try {
				FileInputStream inputStream = new FileInputStream(new File(excelFile));
				Workbook workbook = WorkbookFactory.create(inputStream);

				while(loop) {
					System.out.print("\n Ingrese el nuevo autor del libro: ");

					value = input.nextLine();

					if (value=="") {
						System.out.print("\n No se aceptan campos vacios. Debe ingresar un nombre.");
					}else {
						loop = false;
					}
				}
				System.out.println("valor: "+value);
				Sheet sheet = workbook.getSheetAt(0);
				Cell cell2Update = sheet.getRow(id).getCell(2);
				cell2Update.setCellValue(value);


				inputStream.close();


				FileOutputStream outputStream = new FileOutputStream(excelFile);
				workbook.write(outputStream);
				workbook.close();
				outputStream.close();

			} catch (IOException | EncryptedDocumentException
					| InvalidFormatException ex) {
				ex.printStackTrace();
			}
		}else if (attribute == "Price" && id <= rowCount) {
			try {
				FileInputStream inputStream = new FileInputStream(new File(excelFile));
				Workbook workbook = WorkbookFactory.create(inputStream);

				while(loop) {
					System.out.print("\n Ingrese el nuevo valor para el Precio: ");

					price = input2.nextInt();

					if (price == 0) {
						System.out.print("\n Por favor ingrese un valor: ");
					}else {
						loop = false;
					}
				}

				Sheet sheet = workbook.getSheetAt(0);
				Cell cell2Update = sheet.getRow(id).getCell(3);
				cell2Update.setCellValue(price);

				inputStream.close();

				FileOutputStream outputStream = new FileOutputStream(excelFile);
				workbook.write(outputStream);
				workbook.close();
				outputStream.close();

			} catch (IOException | EncryptedDocumentException
					| InvalidFormatException ex) {
				ex.printStackTrace();
			}
		}else {
			System.out.println("\nError: ese id no existe. Intente nuevamente.");
		}
	}
}
