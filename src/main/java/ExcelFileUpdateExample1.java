import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.util.Scanner;

public class ExcelFileUpdateExample1 {

	static void ActualizarLibro(int id, String atributo, String archivo) {
		String nuevoValor="";
		int precio=0, rowCount=0;
		boolean ciclo = true;
		Scanner entrada = new Scanner(System.in);
		Scanner entrada2 = new Scanner(System.in);
		
		try {
			FileInputStream inputStream = new FileInputStream(new File(archivo));
			Workbook workbook = WorkbookFactory.create(inputStream);
			Sheet sheet = workbook.getSheetAt(0);
			rowCount = sheet.getLastRowNum();
		}catch(Exception e){
			e.printStackTrace();
		}
		
		if (atributo == "Author" && id <= rowCount) {
			try {
				FileInputStream inputStream = new FileInputStream(new File(archivo));
				Workbook workbook = WorkbookFactory.create(inputStream);
				
				while(ciclo) {
					System.out.print("\n Ingrese el nuevo valor para el Autor: ");

			        nuevoValor = entrada.nextLine();
			        
			        if (nuevoValor=="") {
			        	System.out.print("\n Por favor ingrese un valor: ");
			        }else {
			        	ciclo = false;
			        }
				}
				
				Sheet sheet = workbook.getSheetAt(0);
				Cell cell2Update = sheet.getRow(id).getCell(2);
				cell2Update.setCellValue(nuevoValor);
				
				
				inputStream.close();


				FileOutputStream outputStream = new FileOutputStream(archivo);
				workbook.write(outputStream);
				workbook.close();
				outputStream.close();
				
			} catch (IOException | EncryptedDocumentException
					| InvalidFormatException ex) {
				ex.printStackTrace();
			}
		}else if (atributo == "Price" && id <= rowCount) {
			try {
				FileInputStream inputStream = new FileInputStream(new File(archivo));
				Workbook workbook = WorkbookFactory.create(inputStream);
				
				while(ciclo) {
					System.out.print("\n Ingrese el nuevo valor para el Precio: ");

			        precio = entrada2.nextInt();
			        
			        if (precio == 0) {
			        	System.out.print("\n Por favor ingrese un valor: ");
			        }else {
			        	ciclo = false;
			        }
				}
				
				Sheet sheet = workbook.getSheetAt(0);
				Cell cell3Update = sheet.getRow(id).getCell(3);
				cell3Update.setCellValue(precio);
				
				inputStream.close();
				
				FileOutputStream outputStream = new FileOutputStream(archivo);
				workbook.write(outputStream);
				workbook.close();
				outputStream.close();
				
			} catch (IOException | EncryptedDocumentException
					| InvalidFormatException ex) {
				ex.printStackTrace();
			}
			
		}else {
			System.out.println("\nError: esa fila no existe, la ultima fila es:"+rowCount+" y usted ingreso:"+ id);
		}
	
		return;
	}

	static void IngresarRegistro(String nombre, String autor, int precio, String archivo) {

		try {
			FileInputStream inputStream = new FileInputStream(new File(archivo));
			Workbook workbook = WorkbookFactory.create(inputStream);

			
			int sheetCount = workbook.getNumberOfSheets();
			Sheet sheet = workbook.getSheetAt(sheetCount-1);
			int totalRowCount = sheet.getLastRowNum();

			if (totalRowCount >=30){

				Sheet newSheet = workbook.createSheet("Java Books "+sheetCount);

				Object[][] bookData = {
					{nombre,autor,precio},
				};

				int rowCount = 0;

				  Row initRow = newSheet.createRow(rowCount);
				  Cell initCell = initRow.createCell(0);
				  initCell.setCellValue("No");
				  initCell = initRow.createCell(1);
				  initCell.setCellValue("Book Title");
				  initCell = initRow.createCell(2);
				  initCell.setCellValue("Author");
				  initCell = initRow.createCell(3);
				  initCell.setCellValue("Price");
				

            	for (Object[] aBook : bookData) {
					
                	Row row = newSheet.createRow(++rowCount);
					
					int columnCount = 0;
					
					Cell newCell = row.createCell(columnCount);
					newCell.setCellValue(rowCount);
                  
               	 for (Object field : aBook) {
                   	 Cell cell = row.createCell(++columnCount);
                  	  if (field instanceof String) {
                  	      cell.setCellValue((String) field);
                  	  } else if (field instanceof Integer) {
                   	     cell.setCellValue((Integer) field);
                   	 }
                }
                  
            }       
 
            FileOutputStream outputStream = new FileOutputStream(archivo);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

			}else{

					Object[][] bookData = {
							{nombre,autor,precio},
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
					

					inputStream.close();

					FileOutputStream outputStream = new FileOutputStream(archivo);
					workbook.write(outputStream);
					workbook.close();
					outputStream.close();

				}
			
		} catch (IOException | EncryptedDocumentException
				| InvalidFormatException ex) {
			ex.printStackTrace();
		}


	}

	
	public static void main(String[] args) {
		String excelFilePath = "Inventario.xlsx";
		Scanner input = new Scanner(System.in);
	    boolean mainLoop = true;
	    
	    int opcion = 0;
	    while(mainLoop){
	        System.out.println(" Grupo 2 \n");
	        System.out.print("1. Alumno A \n");
	        System.out.print("2. Alumno B \n");
	        System.out.print("3. Alumno C.\n");
	        System.out.print("4. Salir\n");
	        System.out.print("\nSeleccione una opcion: ");

	        opcion = input.nextInt();

	    switch(opcion){

	    case 1: //Alumno A
	        
	        
	        break;

	    case 2: //Alumno B 
			String libro="";
			String autor="";
			int precio=0;
			Scanner scan = new Scanner(System.in);
			
			System.out.println(" Ingrese nuevo registro : \n");

			System.out.println(" Ingrese nombre del libro : \n");
			libro=scan.nextLine();
			System.out.println(" Ingrese autor del libro : \n");
			autor=scan.nextLine();
			System.out.println(" Ingrese precio del libro : \n");
			precio=scan.nextInt();

			IngresarRegistro(libro, autor, precio, excelFilePath);

			System.out.println(" Registro Completado \n");
		
	        break;

	    case 3: // Alumno C Jesus Cadiz
	        int id=0, choice=0, p=0;
	        Scanner registro = new Scanner(System.in);
	        
	        System.out.println(" Indique el numero del registro a actualizar \n");
	        id = registro.nextInt();
	        
	        while(p == 0) {
	        	System.out.println("Indique el atributo a actualizar: \n");
		        System.out.println(" 1. Author \n");
		        System.out.println(" 2. Price \n");
		        
		        choice = input.nextInt();
		        
		        if (choice == 1) {
		        	ActualizarLibro(id,"Author", excelFilePath);
		        	p = 1;
		        }else if (choice == 2) {
		        	ActualizarLibro(id,"Price" , excelFilePath);
		        	p = 1;
		        }else {
		        	System.out.println("Seleccione una opcion valida \n");
		        }
	        }
	    	break;
	    case 4: // salir
	    	mainLoop = false;
	    	break;
	    default :
	             System.out.println("Opcion no valida!");
	             break;
	    }


	    }
		/* 
		try {
			FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
			Workbook workbook = WorkbookFactory.create(inputStream);

			Sheet sheet = workbook.getSheetAt(0);

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

			inputStream.close();

			FileOutputStream outputStream = new FileOutputStream(excelFilePath);
			workbook.write(outputStream);
			workbook.close();
			outputStream.close();
			
		} catch (IOException | EncryptedDocumentException
				| InvalidFormatException ex) {
			ex.printStackTrace();
		}
		
	}
	
	*/
	}

}
