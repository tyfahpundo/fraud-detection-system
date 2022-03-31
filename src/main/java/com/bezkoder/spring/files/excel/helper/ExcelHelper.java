package com.bezkoder.spring.files.excel.helper;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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

      // Header
      Row headerRow = sheet.createRow(0);

      for (int col = 0; col < HEADERs.length; col++) {
        Cell cell = headerRow.createCell(col);
        cell.setCellValue(HEADERs[col]);
      }

      int rowIdx = 1;
      List<Long> digitList = new ArrayList<>();
      for (Finance finance : finances) {
        Row row = sheet.createRow(rowIdx++);

        long number = finance.getSales();
        long firstDigit = 0;
        while(number != 0){
            firstDigit = number % 10;
            number = number/10;
        }
        digitList.add(firstDigit);

        row.createCell(0).setCellValue(finance.getId());
        row.createCell(1).setCellValue(finance.getDescription());
        row.createCell(2).setCellValue(finance.getSales());
        row.createCell(3).setCellValue(firstDigit);
        long finalFirstDigit = firstDigit;
        long count = digitList.stream()
                .filter(x-> x == finalFirstDigit)
                .count();
        row.createCell(4).setCellValue(count);
        float percentage = ((count/digitList.size())*100);
        row.createCell(5).setCellValue( new BigDecimal(percentage).floatValue());
        row.createCell(6).setCellValue(new BigDecimal(Math.log10((1/firstDigit)+1)).floatValue());
        row.createCell(7).setCellValue(new BigDecimal(percentage-(Math.log10((1/firstDigit)+1))).floatValue());


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
