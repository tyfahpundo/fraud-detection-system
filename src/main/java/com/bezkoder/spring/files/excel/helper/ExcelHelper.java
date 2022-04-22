package com.bezkoder.spring.files.excel.helper;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFExtendedColor;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import com.bezkoder.spring.files.excel.model.Finance;

public class ExcelHelper {
  public static String TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  static String[] HEADERs = { "Id", "Description", "Sales","FirstDigit","Count","Percentage","Benford","Difference" };
  static String SHEET = "FinacialStatement";

  public static boolean hasExcelFormat(MultipartFile file) {
    if (!TYPE.equals(file.getContentType())) {
      return false;
    }
    return true;
  }

  public static ByteArrayInputStream tutorialsToExcel(List<Finance> finances) {

    try (Workbook workbook = new XSSFWorkbook();
         ByteArrayOutputStream out = new ByteArrayOutputStream()) {
      Sheet sheet = workbook.createSheet(SHEET);
      CellStyle style = workbook.createCellStyle();
      style.setFillBackgroundColor(IndexedColors.RED.getIndex());
      style.setFillPattern(FillPatternType.BIG_SPOTS);
      // Header
      Row headerRow = sheet.createRow(0);
      for (int col = 0; col < HEADERs.length; col++) {
        Cell cell = headerRow.createCell(col);
        cell.setCellValue(HEADERs[col]);
      }

      int rowIdx = 1;
      List<Long> digitList = new ArrayList<>();
      for (Finance finance : finances) {
        long number = finance.getSales();
        long firstDigit = 0;
        while (number != 0) {
          firstDigit = number % 10;
          number = number / 10;
        }
        digitList.add(firstDigit);
      }
      for (Finance finance : finances) {
        Row row = sheet.createRow(rowIdx++);

        long number = finance.getSales();
        long firstDigit = 0;
        while(number != 0){
            firstDigit = number % 10;
            number = number/10;
        }

        row.createCell(0).setCellValue(finance.getId());
        row.createCell(1).setCellValue(finance.getDescription());
        row.createCell(2).setCellValue(finance.getSales());
        row.createCell(3).setCellValue(firstDigit);
        long finalFirstDigit = firstDigit;
        long count = digitList.stream()
                .filter(x-> x == finalFirstDigit)
                .count();
        row.createCell(4).setCellValue(count);
        double percentage =((double) count/digitList.size())*100;
        row.createCell(5).setCellValue(String.valueOf(percentage));

        switch((int) row.getCell(4).getNumericCellValue()){
          case 1:
            if(!row.getCell(5).getStringCellValue().equals(String.valueOf(30.1)) ){
              row.getCell(5).setCellStyle(style);
            }
            break;

          case 2:
            if(!row.getCell(5).getStringCellValue().equals(String.valueOf(17.6)) ){
              row.getCell(5).setCellStyle(style);
            }
            break;
          case 3:
            if(!row.getCell(5).getStringCellValue().equals(String.valueOf(12.5)) ){
              row.getCell(5).setCellStyle(style);
            }
            break;
          case 4:
            if(!row.getCell(5).getStringCellValue().equals(String.valueOf(9.7)) ){
              row.getCell(5).setCellStyle(style);
            }
            break;

          case 5:
            if(!row.getCell(5).getStringCellValue().equals(String.valueOf(7.9)) ){
              row.getCell(5).setCellStyle(style);
            }
            break;

          case 6:
            if(!row.getCell(5).getStringCellValue().equals(String.valueOf(6.7)) ){
              row.getCell(5).setCellStyle(style);
            }
            break;
          case 7:
            if(!row.getCell(5).getStringCellValue().equals(String.valueOf(5.8)) ){
              row.getCell(5).setCellStyle(style);
            }
            break;
          case 8:
            if(!row.getCell(5).getStringCellValue().equals(String.valueOf(5.1)) ){
              row.getCell(5).setCellStyle(style);
            }
            break;

          case 9:
            if(!row.getCell(5).getStringCellValue().equals(String.valueOf(4.6)) ){
              row.getCell(5).setCellStyle(style);
            }
            break;

          default:
            break;
        }
        row.createCell(6).setCellValue(String.valueOf(Math.log10(((double) 1/firstDigit)+1)));
        row.createCell(7).setCellValue(String.valueOf(percentage-(Math.log10(((double) 1/firstDigit)+1))));


      }

      workbook.write(out);
      return new ByteArrayInputStream(out.toByteArray());
    } catch (IOException e) {
      throw new RuntimeException("fail to import data to Excel file: " + e.getMessage());
    }
  }

  public static List<Finance> excelToTutorials(InputStream is) {
    try {
      Workbook workbook = new XSSFWorkbook(is);

      Sheet sheet = workbook.getSheetAt(0);
      Iterator<Row> rows = sheet.iterator();

      List<Finance> finances = new ArrayList<>();

      int rowNumber = 0;
      while (rows.hasNext()) {
        Row currentRow = rows.next();

        // skip header
        if (rowNumber == 0) {
          rowNumber++;
          continue;
        }

        Iterator<Cell> cellsInRow = currentRow.iterator();

        Finance finance = new Finance();

        int cellIdx = 0;
        while (cellsInRow.hasNext()) {
          Cell currentCell = cellsInRow.next();

          switch (cellIdx) {
          case 0:
            finance.setId((long) currentCell.getNumericCellValue());
            break;

          case 1:
            finance.setDescription(currentCell.getStringCellValue());
            break;

          case 2:
            finance.setSales((long) currentCell.getNumericCellValue());
            break;

          default:
            break;
          }

          cellIdx++;
        }

        finances.add(finance);
      }

      workbook.close();

      return finances;
    } catch (IOException e) {
      throw new RuntimeException("fail to parse Excel file: " + e.getMessage());
    }
  }
}
