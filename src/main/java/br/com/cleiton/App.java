package br.com.cleiton;

import java.io.File;
import java.io.FileFilter;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.Charset;
import java.util.Date;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.jamel.dbf.DbfReader;
import org.jamel.dbf.structure.DbfField;
import org.jamel.dbf.structure.DbfHeader;

/**
 * Hello world!
 *
 */
public class App {

	public static void main(String[] args) throws Exception {
		try {
			File diretorioDeImagens = new File("C:\\PDVARQ");
			File diretorioOrigem = new File("C:\\BI");

			if (!diretorioOrigem.exists()) {
				diretorioOrigem.mkdirs();
			}

			File[] imagensDoDiretorio = diretorioDeImagens.listFiles(new FileFilter() {
				public boolean accept(File b) {
					return b.getName().endsWith(".dbf");
				}
			});
			for (File file : imagensDoDiretorio) {
				writeToTxtFile(file,
						new File(diretorioOrigem.getAbsolutePath() + "/" + file.getName().replaceAll(".dbf", ".xlsx")),
						Charset.forName("ISO-8859-1"));
			}
		} catch (Exception e) {
			e.printStackTrace();
			FileUtils.writeByteArrayToFile(new File("C:\\BI\\erro.txt"), e.getMessage().getBytes());
		}

	}

	public static void writeToTxtFile(File dbf, File txt, Charset dbfEncoding) throws Exception {

		DbfReader reader = new DbfReader(dbf);
		DbfHeader header = reader.getHeader();

		String[] titles = new String[header.getFieldsCount()];
		for (int i = 0; i < header.getFieldsCount(); i++) {
			DbfField field = header.getField(i);
			titles[i] = field.getName();
		}

		SXSSFWorkbook workbook = new SXSSFWorkbook(1000);
		SXSSFSheet sheet = (SXSSFSheet) workbook.createSheet("Planilha");
		OutputStream out = new FileOutputStream(txt);
		criarCabecalho(titles, sheet);
		Object[] row;
		while ((row = reader.nextRecord()) != null) {
			org.apache.poi.ss.usermodel.Row poiRow = sheet.createRow(sheet.getPhysicalNumberOfRows());
			CellStyle cellStyle = workbook.createCellStyle();
			CreationHelper createHelper = workbook.getCreationHelper();
			cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/mm/YYYY"));

			for (int i = 0; i < header.getFieldsCount(); i++) {
				DbfField field = header.getField(i);
				byte dataType = field.getDataType();

				Cell createCell = poiRow.createCell(i);
				switch (dataType) {
				case 'C':
					createCell.setCellValue(new String((byte[]) row[i], dbfEncoding));
					break;
				case 'D':
					createCell.setCellStyle(cellStyle);
					createCell.setCellValue((Date) row[i]);
					break;
				case 'F':
					break;
				case 'L':
					criarBoleano(row, i, createCell);
					break;
				case 'N':
					criarNumero(row, i, createCell);
					break;
				case 'M':
					break;
				case 'T':
					break;
				default:

				}

			}
		}

		fecharArquivo(titles, workbook, sheet, out);
		reader.close();
	}

	private static void criarNumero(Object[] row, int i, Cell createCell) {
		java.lang.Number number2 = (Number) row[i];
		if (number2 != null) {
			Double valor = number2.doubleValue();
			createCell.setCellValue(valor);

		}
	}

	private static void criarBoleano(Object[] row, int i, Cell createCell) {
		Boolean valor = (Boolean) row[i];
		createCell.setCellValue(valor);
	}

	private static void criarCabecalho(String[] titles, SXSSFSheet sheet) throws IOException {
		int cellIndex = 0;
		org.apache.poi.ss.usermodel.Row poiRow = sheet.createRow(0);
		for (String value : titles) {
			Cell cell = poiRow.createCell(cellIndex++);
			cell.setCellValue(value);
		}

	}

	private static void fecharArquivo(String[] titles, SXSSFWorkbook workbook, SXSSFSheet sheet, OutputStream out)
			throws IOException {
		for (int columnIndex = 0; columnIndex < titles.length; columnIndex++) {
			sheet.trackColumnForAutoSizing(columnIndex);
			sheet.autoSizeColumn(columnIndex);
		}
		workbook.write(out);
		out.flush();
		out.close();
	}
}
