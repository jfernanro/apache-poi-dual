package com.dual.apachepoi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExportarExcel
{

  private static final String FILE_NAME = "/tmp/MiPrimerExcel.xlsx";

  public static void main(String[] args)
  {
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet sheet = workbook.createSheet("Datatypes in Java");
    Object[][] datos =
      {{"Nombre","Apellido 1","Apellido 2","Edad"},{"Juan","Fernandez","Rodriguez",36},{"Manuel","Perez","Gonzalez",40},
        {"Alfonso","Jimenez","Márquez",41},{"Pedro","Martinez","Alcántara",28},{"Francisco","Fernández","Alvarez",42}};

    int rowNum = 0;
    System.out.println("Exportando excel...");

    for(Object[] datatype:datos)
    {
      Row row = sheet.createRow(rowNum++);
      int colNum = 0;
      for(Object field:datatype)
      {
        Cell cell = row.createCell(colNum++);
        if(field instanceof String)
        {
          cell.setCellValue((String)field);
        }
        else if(field instanceof Integer)
        {
          cell.setCellValue((Integer)field);
        }
      }
    }

    try
    {
      FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
      workbook.write(outputStream);
      workbook.close();

      System.out.println("Exito!");
    }
    catch(FileNotFoundException e)
    {
      System.out.println("Error");
      e.printStackTrace();
    }
    catch(IOException e)
    {
      System.out.println("Error");
      e.printStackTrace();
    }

    System.out.println("Fin");
  }

}
