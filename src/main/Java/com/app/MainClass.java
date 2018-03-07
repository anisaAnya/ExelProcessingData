package com.app;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;

public class MainClass {

	static Element lastElement;
	static String[] remarks = {"Не ходит", "Постоянно", "Не выдает", "Не придерживается", "Читает лекции", 
	"Не успевает", "Некорректно", "Методичек нет", "Предмет", "Название"};

    public static void main(String[] argv) throws IOException {
		Element temp ;

		readTeacherList(); 
		temp = lastElement;
		while(temp != null) {
			System.out.println(temp.name);
			temp = temp.prev;
    	}
    	File myFolder = new File("D:/project/exel/res"); //путь к папке с всеми таблицами результатов опроса
   		File[] files = myFolder.listFiles();
   		for (int i = 0; i < files.length; i++) {
   			readDataFile(lastElement, files[i].toString());	
   		}
    	SaveInFile (lastElement);
    	
    }

    static boolean checkForWord(String line, String word){
            return line.contains(word);
    }

    public static void readTeacherList() throws IOException
	{
		InputStream ExcelFileToRead = new FileInputStream("Teacher.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);

		XSSFSheet sheet = wb.getSheetAt(0);
		XSSFRow row; 
		XSSFCell cell;

		Iterator rows = sheet.rowIterator();

		while (rows.hasNext())
		{
			row=(XSSFRow) rows.next();
			Iterator cells = row.cellIterator();
			while (cells.hasNext()) {
			
				cell=(XSSFCell) cells.next();
		
					if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING) {
						lastElement = new Element(lastElement, cell.getStringCellValue());
					} else if(cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
						System.out.print(cell.getNumericCellValue()+" ");
					} else {
					//U Can Handel Boolean, Formula, Errors
				}
			}
		}
			System.out.println();
	}

    public static void readDataFile(Element last, String file) throws IOException
	{
		Element current = null;;
		Element temp;
		InputStream ExcelFileToRead = new FileInputStream(file);
		XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);

		XSSFSheet sheet = wb.getSheetAt(0);
		XSSFRow row;  
		XSSFCell cell;
		int rowNum = 0;
		int cellNum = 0;

		row = sheet.getRow(rowNum);
		cell = row.getCell(cellNum);

		while(cell != null) {
			if (checkForWord(cell.getStringCellValue(), "Отметь галочками")) {
				temp = last;
				while (temp != null) {
					if (checkForWord(cell.getStringCellValue(), temp.name)) {
						current = temp;
						break;
					}
					temp = temp.prev;
				}
				if (current != null) {
					rowNum++;
					row = sheet.getRow(rowNum);
					cell = row.getCell(cellNum);
					for(int i = 0; i < sheet.getLastRowNum(); i++) {
						if(cell != null) {
							//имеем массив стрингов замечаний сравниваем содержание клетки
							//на каждый стринг из массива и если находим совпадение увеличиваем каунтер
							for(int a = 0; a < 10; a++) {
								if(checkForWord(cell.getStringCellValue(), remarks[a])) {
									current.countRemark[a]++;
									current.commonCount++;
								}
							}
						}
						rowNum++;
						row = sheet.getRow(rowNum);
						if( row != null) cell = row.getCell(cellNum);
					}
				}
				rowNum = 0;
				current = null;
				cellNum++;
				row = sheet.getRow(rowNum);
				cell = row.getCell(cellNum); 
			}
			if (checkForWord(cell.getStringCellValue(), "Оцени в целом")) {
				temp = last;
				while (temp != null) {
					if (checkForWord(cell.getStringCellValue(), temp.name)) {
						current = temp;
						current.respondentsNum += sheet.getLastRowNum();
						break;
					}
					temp = temp.prev;
				}
				if (current != null) {
					rowNum++;
					row = sheet.getRow(rowNum);
					cell = row.getCell(cellNum);
					for(int i = 0; i < sheet.getLastRowNum(); i++) {
						//System.out.println(sheet.getLastRowNum());
						if(cell != null) {
							if(current.average == 0) {
								current.average = cell.getNumericCellValue();
							} else {
								current.average = (current.average + cell.getNumericCellValue()) / 2;
							}
						}
						rowNum++;
						row = sheet.getRow(rowNum);
						if( row != null) cell = row.getCell(cellNum);
					}
				}
			}
			rowNum = 0;
			current = null;
			cellNum++;
			row = sheet.getRow(rowNum);
			cell = row.getCell(cellNum);
		}
	
	}

    public static void SaveInFile (Element last) throws IOException {
    	XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Просто лист");
        Element current = null;;
		Element temp;

		XSSFRow row; 
		XSSFCell cell;
		int rowNum = 0;
		int cellNum = 0;

		row = sheet.createRow(rowNum);
 		// создаем подписи к столбцам (это будет первая строчка в листе Excel файла)
        row.createCell(0).setCellValue("Фамилия");
        row.createCell(1).setCellValue("Кол-во опрошеных");
        row.createCell(2).setCellValue("Кол-во замечаний");
        row.createCell(3).setCellValue("Средний балл");
        row.createCell(4).setCellValue("Не ходит на пары");
        row.createCell(5).setCellValue("Постоянно опаздывает на пары");
        row.createCell(6).setCellValue("Не выдает критерии оценивания и список тем");
        row.createCell(7).setCellValue("Не придерживается критериев оценивания");
        row.createCell(8).setCellValue("Читает лекции непонятно");
        row.createCell(9).setCellValue("Не успевает принять лабораторные");
        row.createCell(10).setCellValue("Не корректно обращается со студентами");
        row.createCell(11).setCellValue("Методичек нет или они не качественные");
        row.createCell(12).setCellValue("Предмет не актуален");
        row.createCell(13).setCellValue("Название предмета не соответствует материалу");

        for(int a = 0; a < 15; a++) {
        	sheet.autoSizeColumn(a);
        }
 
 		temp = last;
 		rowNum = 1;
    	while(temp != null) {
    		row = sheet.createRow(rowNum);
			row.createCell(0).setCellValue(temp.name);
			row.createCell(2).setCellValue(temp.commonCount);
        	row.createCell(3).setCellValue(temp.average);
        	row.createCell(1).setCellValue(temp.respondentsNum);
        	for(int i = 0; i < 10; i++) {
        		row.createCell(i + 4).setCellValue(temp.countRemark[i]);
        	}
        	temp = temp.prev;
        	rowNum++;

    	}
 
        // записываем созданный в памяти Excel документ в файл
        try (FileOutputStream out = new FileOutputStream(new File("File.xlsx"))) {
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Excel файл успешно создан!");
    }
    
}

 class Element {
	String name; //Фамилия преподавателя
	Element next; 
	Element prev;
	double average = 0; //средний балл
	int[] countRemark = new int[10]; //кол-во замечаний из общего списка
	int respondentsNum = 0; //кол-во опрошеных студентов у которых вел этот преподаватель
	int commonCount = 0; //общее кол-во замечаний

	Element(Element last, String nameData) {
		if(last != null) {
			this.prev = last;
			last.next = this;
		}
		name = nameData;
	}
}