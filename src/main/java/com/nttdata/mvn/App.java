package com.nttdata.mvn;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Clase para crear y leer archivos Excel.
 *
 */
public class App {
	public static void main(String[] args) {

		EscribirEXCEL();
		LeerEXCEL();

	}

	/**
	 * Metodo para leer archivos Excel.
	 */

	private static void LeerEXCEL() {

		String nombreArchivo = "ListaLibros.xlsx";
		String hoja = "Libros";

		try (FileInputStream file = new FileInputStream(new File(nombreArchivo))) {

			// Leer archivo de Excel

			XSSFWorkbook libro = new XSSFWorkbook(file);

			// Obtener la hoja que se va a leer

			XSSFSheet sheet = libro.getSheetAt(0);

			// Obtener todas las filas de la hoja de Excel

			Iterator<Row> rowIterator = sheet.iterator();

			Row row;

			// Se recorre cada fila hasta el final

			while (rowIterator.hasNext()) {
				row = rowIterator.next();

				// Se obtienen las celdas por fila

				Iterator<Cell> cellIterator = row.cellIterator();
				Cell cell;

				// Se recorre cada celda

				while (cellIterator.hasNext()) {

					// Se obtiene la celda en especifico y se imprime

					cell = cellIterator.next();
					System.out.print(cell.getStringCellValue() + " - ");
				}
				System.out.println("");
			}

		} catch (Exception e) {
			e.getMessage();
		}
	}
	/**
	 * Método para crear Excel mediante eclipse.
	 */

	private static void EscribirEXCEL() {

		String nombreArchivo = "ListaLibros.xlsx";

		String hoja = "Libros";

		XSSFWorkbook libro = new XSSFWorkbook();
		XSSFSheet hoja1 = libro.createSheet(hoja);

		// Cabecera de la hoja de excel

		String[] header = new String[] { "NOMBRE", "GENERO", "AUTOR" };

		// Contenido de la hoja de excel

		String[][] document = new String[][] { { "Los Girasoles Ciegos", "Ficcion", "Alberto Mendez" },
				{ "Los Pilares de la Tierra", "Novela", "Ken Follett" }, { "Donde nadie te encuentre", "Misterio", "Alicia Gimenez Bartlett" } };

		// Poner en negrita la cabecera

		CellStyle style = libro.createCellStyle();
		Font font = libro.createFont();
		font.setBold(true);
		style.setFont(font);

		// Generar los datos para el documento

		for (int i = 0; i <= document.length; i++) {
			XSSFRow row = hoja1.createRow(i); // Se crea la fila
			for (int j = 0; j < header.length; j++) {
				if (i == 0) { // Para la cabecera
					XSSFCell cell = row.createCell(j); // Se crean las celdas pra la cabecera
					cell.setCellValue(header[j]); // Se añade el contenido
				} else {
					XSSFCell cell = row.createCell(j); // Se crean las celdas para el contenido
					cell.setCellValue(document[i - 1][j]); // Se añade el contenido
				}
			}
		}

		// Crear el archivo

		try (OutputStream fileOut = new FileOutputStream(nombreArchivo)) {
			System.out.println("SE CREO EL EXCEL");
			libro.write(fileOut);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
