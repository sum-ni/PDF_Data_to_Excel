import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;

public class Set_cell extends PDF_Data{
	public static void set_title_style(CellStyle title_style) {
		title_style.setAlignment(HorizontalAlignment.CENTER);
		title_style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		title_style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		title_style.setAlignment(HorizontalAlignment.CENTER);
		border_line(title_style);
	}
	public static void border_line(CellStyle style) {
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
	}
	public static void set_style(CellStyle style, int cnt) {
		for (int j = 0; j < cnt; j++) {
			row.getCell(j).setCellStyle(style);
		}
	}
	public static void set_cell_value(String[] cell_data) {
		row = sheet.createRow(cnt++);
		for (int i = 0; i < cell_data.length; i++) {
			row.createCell(i).setCellValue(cell_data[i]);
		}
	}
}
