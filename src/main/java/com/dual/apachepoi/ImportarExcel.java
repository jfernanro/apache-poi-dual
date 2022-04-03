package com.dual.apachepoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ImportarExcel
{

  private static final String FILE_NAME = "/tmp/MiPrimerExcel.xlsx";

  public static void main(String[] args)
  {
    try
    {

      FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
      Workbook workbook = new XSSFWorkbook(excelFile);
      Sheet datatypeSheet = workbook.getSheetAt(0);
      Iterator<Row> iterator = datatypeSheet.iterator();

      System.out.println("Importando excel ...");
      while(iterator.hasNext())
      {

        Row currentRow = iterator.next();
        Iterator<Cell> cellIterator = currentRow.iterator();

        while(cellIterator.hasNext())
        {

          Cell currentCell = cellIterator.next();
          // getCellTypeEnum aparecera como obsoleto a partir de la 3.15
          // getCellTypeEnum sera renombrado a getCellType a partir de la version 4.0
          if(currentCell.getCellTypeEnum() == CellType.STRING)
          {
            System.out.print(currentCell.getStringCellValue() + "-----");
          }
          else if(currentCell.getCellTypeEnum() == CellType.NUMERIC)
          {
            System.out.print(currentCell.getNumericCellValue() + "-----");
          }

        }
      }
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
    finally
    {
      System.out.println("Fin");
    }
  }

}
