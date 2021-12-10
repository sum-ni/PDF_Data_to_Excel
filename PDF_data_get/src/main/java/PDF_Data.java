import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PDF_Data {
	static int cnt = 0;
	static XSSFWorkbook wb = new XSSFWorkbook();
	static XSSFSheet sheet = wb.createSheet("DATA");
	static XSSFRow row;
	static Set_cell s = new Set_cell();

	private static String read_pdf(PDDocument doc) throws IOException {
		PDFTextStripper stripper = new PDFTextStripper();

		stripper.setStartPage(22);
		stripper.setEndPage(65);

		String input_text = stripper.getText(doc);

		return input_text;
	}
	private static void write_excel(String Data) {
		CellStyle title_style = wb.createCellStyle();
		s.set_title_style(title_style);
		
		Font font = wb.createFont();
		font.setBold(true);
		title_style.setFont(font);

		String[] cell_title = { "id", "name", "src", "note", "Type" };
		row = sheet.createRow(cnt++);
		for (int k = 0; k < cell_title.length; k++) {
			row.createCell(k).setCellValue(cell_title[k]);
		}
		s.set_style(title_style, cell_title.length);
		
		String id = "", name = "", src = "", note = "", type = "";
		String pattern_id = "5\\.4\\.(1|2)\\.(1|2|4)\\..+";
		String pattern_name = "name.+";
		String pattern_src = "src.+";
		String pattern_note = "note.+";
		String pattern_type = "Type.+";

		Pattern p_id = Pattern.compile(pattern_id);
		Pattern p_name = Pattern.compile(pattern_name);
		Pattern p_src = Pattern.compile(pattern_src);
		Pattern p_note = Pattern.compile(pattern_note);
		Pattern p_type = Pattern.compile(pattern_type);

		/* */
		Scanner scan = new Scanner(Data);
		ArrayList<String> id_list = new ArrayList<String>();
		ArrayList<String> name_list = new ArrayList<String>();
		ArrayList<String> src_list = new ArrayList<String>();
		ArrayList<String> note_list = new ArrayList<String>();
		ArrayList<String> type_list = new ArrayList<String>();
		while (scan.hasNext()) {
			String line = scan.nextLine();
			Boolean b_src = false, b_note = false;

			Matcher m_id = p_id.matcher(line);
			Matcher m_name = p_name.matcher(line);
			Matcher m_src = p_src.matcher(line);
			Matcher m_note = p_note.matcher(line);
			Matcher m_type = p_type.matcher(line);
			
			if (m_id.find()) {
				String str = m_id.group(0);
				if(str.charAt(str.length()-1)==',') {
					str += scan.nextLine();
				}
				id_list.add(str);
			}
			if (m_name.find()) {
				name_list.add(m_name.group(0));
			}
			b_src = m_src.find();
			b_note = m_note.find();
			if (b_src||b_note) { 
				if(b_src) {
					src_list.add(m_src.group(0));
					note_list.add("");
				}else if(b_note){
					src_list.add("");
					note_list.add(m_note.group(0));
				}
			}
			if (m_type.find()) {
				type_list.add(m_type.group(0));
			}
		}
		
		CellStyle style = wb.createCellStyle();
		s.border_line(style);
		
		if((id_list.size()==name_list.size()) && (id_list.size() == src_list.size()) &&
				(id_list.size()==note_list.size()) && (id_list.size() == type_list.size())){
			for(int i = 0;i<id_list.size();i++) {
				String id_reg = "\\d\\.\\d\\.\\d\\.\\S+ ";
				String regex = "\\S+ :      ";
				id = id_list.get(i).replaceAll(id_reg,"");
				name = name_list.get(i).replaceAll(regex, "");
				src = src_list.get(i).replaceAll(regex, "");
				note = note_list.get(i).replaceAll(regex, "");
				type = type_list.get(i).replaceAll(regex, "");
	//			
				String[] cell_input = { id, name, src, note, type };
				s.set_cell_value(cell_input);
				s.set_style(style, cell_input.length);
			}
		}

		scan.close();
	}
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		String fileName = "./input.pdf";
		String pdf_Data = null;
		File file = new File(fileName);
		PDDocument doc = PDDocument.load(file);
		pdf_Data = read_pdf(doc);

		/* */
		write_excel(pdf_Data);
		/* */

		FileOutputStream fout = new FileOutputStream("./result.xlsx");
		wb.write(fout);
		fout.close();

		doc.close();
	}

}
