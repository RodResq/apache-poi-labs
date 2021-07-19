import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelPOIHelp {
	
	
	public void writeExcel() throws IOException {
		
		Workbook workbook = new XSSFWorkbook();
		
		try {
			Sheet sheet = workbook.createSheet("Persons");
			sheet.setColumnWidth(0, 6000);
			sheet.setColumnWidth(1, 4000);
			
			Row header = sheet.createRow(0);
			
			CellStyle headerStyle = workbook.createCellStyle();
			headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
			headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			
			
			XSSFFont font = ((XSSFWorkbook) workbook).createFont();
			font.setFontName("Arial");
			font.setFontHeightInPoints((short) 16);
			font.setBold(true);
			headerStyle.setFont(font);
			
			Cell headerCell = header.createCell(0);
			headerCell.setCellValue("Name");
			headerCell.setCellStyle(headerStyle);
			
			headerCell = header.createCell(1);
			headerCell.setCellValue("Age");
			headerCell.setCellStyle(headerStyle);
			
			
			/* Criacao do conteudo da tabela com stylo diferente**/
			CellStyle style = workbook.createCellStyle();
			style.setWrapText(true);
			
			Row row = sheet.createRow(1);
			
			Cell cell = row.createCell(0);
			cell.setCellValue("John Smith");
			cell.setCellStyle(style);
			
			Cell cell2 = row.createCell(1);
			cell2.setCellValue(20);
			cell2.setCellStyle(style);
			
			
			File currDir = new File(".");
			String path = currDir.getAbsolutePath();
			System.out.println("Absolut path: " + path);
			
			String fileLocation = path.subSequence(0, path.length() -1) + "temp2.xlsx";
			FileOutputStream outputStream = new FileOutputStream(fileLocation);
			workbook.write(outputStream);
		} finally {
            if (workbook != null) {
                    workbook.close();
            }
        }
	}

}
