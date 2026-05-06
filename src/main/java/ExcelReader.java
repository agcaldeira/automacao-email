import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

	// Atualizado para novo layout com colunas NF-e e deslocamento de status/timestamp
	private static final int STATUS_COLUMN_DEFAULT = 24;
	private static final int TIMESTAMP_COLUMN_DEFAULT = 25;

	public static List<ClientInvoice> readClients(String path) {
		List<ClientInvoice> clientes = new ArrayList<>();
		try (FileInputStream fis = new FileInputStream(path); Workbook wb = new XSSFWorkbook(fis)) {

			Sheet sheet = wb.getSheetAt(0); // primeira aba

			// Detectar colunas de Status e Timestamp a partir do header (linha 0)
			int statusCol = STATUS_COLUMN_DEFAULT;
			int tsCol = TIMESTAMP_COLUMN_DEFAULT;
			Row header = sheet.getRow(0);
			if (header != null) {
				for (Cell cell : header) {
					String h = getCellString(cell);
					if (h == null)
						continue;
					String lower = h.trim().toLowerCase();
					if (lower.contains("status") || lower.contains("situação") || lower.contains("situacao")) {
						statusCol = cell.getColumnIndex();
					}
					if (lower.contains("data") || lower.contains("horário") || lower.contains("horario") || lower.contains("envio") || lower.contains("timestamp")) {
						tsCol = cell.getColumnIndex();
					}
				}
			}

			// Iterar linhas (pular header)
			for (Row row : sheet) {
				if (row.getRowNum() == 0)
					continue; // pular header
				if (row == null)
					continue;

				String status = getCellString(row.getCell(statusCol));
				if (status == null || !status.equalsIgnoreCase("Pendente")) {
					// pular linhas que não estejam marcadas como Pendente
					continue;
				}

				// Leitura por índice de coluna (0-based) - conforme nova ordem informada
				String nome = getCellString(row.getCell(0));
				String cnpj = getCellString(row.getCell(1));
				String email = getCellString(row.getCell(2));
				Integer qtdConsultas = getCellInteger(row.getCell(3));
				BigDecimal valUnitConsulta = getCellDecimal(row.getCell(4));
				BigDecimal totalConsultas = getCellDecimal(row.getCell(5));
				Integer qtdSerasa = getCellInteger(row.getCell(6));
				BigDecimal valUnitSerasa = getCellDecimal(row.getCell(7));
				BigDecimal totalSerasa = getCellDecimal(row.getCell(8));
				Integer qtdNeg = getCellInteger(row.getCell(9));
				BigDecimal valUnitNeg = getCellDecimal(row.getCell(10));
				BigDecimal totalNeg = getCellDecimal(row.getCell(11));
				Integer qtdExc = getCellInteger(row.getCell(12));
				BigDecimal valUnitExc = getCellDecimal(row.getCell(13));
				BigDecimal totalExc = getCellDecimal(row.getCell(14));
				Integer qtdSms = getCellInteger(row.getCell(15));
				BigDecimal valUnitSms = getCellDecimal(row.getCell(16));
				BigDecimal totalSms = getCellDecimal(row.getCell(17));
				// NF-e
				Integer qtdNf = getCellInteger(row.getCell(18));
				BigDecimal valUnitNf = getCellDecimal(row.getCell(19));
				BigDecimal totalNf = getCellDecimal(row.getCell(20));
				BigDecimal mensalidade = getCellDecimal(row.getCell(21));
				BigDecimal totalFinal = getCellDecimal(row.getCell(22));
				String arquivo = getCellString(row.getCell(23));

				// Calcula totais quando ausentes
				if (totalConsultas == null && qtdConsultas != null && valUnitConsulta != null) {
					totalConsultas = valUnitConsulta.multiply(BigDecimal.valueOf(qtdConsultas));
				}
				if (totalSerasa == null && qtdSerasa != null && valUnitSerasa != null) {
					totalSerasa = valUnitSerasa.multiply(BigDecimal.valueOf(qtdSerasa));
				}
				if (totalNeg == null && qtdNeg != null && valUnitNeg != null) {
					totalNeg = valUnitNeg.multiply(BigDecimal.valueOf(qtdNeg));
				}
				if (totalExc == null && qtdExc != null && valUnitExc != null) {
					totalExc = valUnitExc.multiply(BigDecimal.valueOf(qtdExc));
				}
				if (totalSms == null && qtdSms != null && valUnitSms != null) {
					totalSms = valUnitSms.multiply(BigDecimal.valueOf(qtdSms));
				}
				if (totalNf == null && qtdNf != null && valUnitNf != null) {
					totalNf = valUnitNf.multiply(BigDecimal.valueOf(qtdNf));
				}
				if (totalFinal == null) {
					totalFinal = BigDecimal.ZERO;
					if (totalConsultas != null)
						totalFinal = totalFinal.add(totalConsultas);
					if (totalSerasa != null)
						totalFinal = totalFinal.add(totalSerasa);
					if (totalNeg != null)
						totalFinal = totalFinal.add(totalNeg);
					if (totalExc != null)
						totalFinal = totalFinal.add(totalExc);
					if (totalSms != null)
						totalFinal = totalFinal.add(totalSms);
					if (totalNf != null)
						totalFinal = totalFinal.add(totalNf);
					if (mensalidade != null)
						totalFinal = totalFinal.add(mensalidade);
				}

				ClientInvoice ci = new ClientInvoice(nome, cnpj, email, safeInt(qtdConsultas), valUnitConsulta, totalConsultas, safeInt(qtdSerasa), valUnitSerasa, totalSerasa, safeInt(qtdNeg), valUnitNeg, totalNeg, safeInt(qtdExc), valUnitExc, totalExc, safeInt(qtdSms), valUnitSms, totalSms, safeInt(qtdNf), valUnitNf, totalNf, mensalidade, totalFinal, arquivo, row.getRowNum(), path);
				clientes.add(ci);
			}

		} catch (Exception ex) {
			System.err.println("Erro lendo Excel: " + ex.getMessage());
			ex.printStackTrace();
		}
		return clientes;
	}

	/**
	 * Marca a linha do Excel (rowIndex, 0-based) como Enviado e escreve a data/hora na coluna adjacente.
	 * Retorna true se gravou com sucesso.
	 */
	public static boolean markRowAsSent(String path, int rowIndex) {
		try (FileInputStream fis = new FileInputStream(path); Workbook wb = new XSSFWorkbook(fis)) {
			Sheet sheet = wb.getSheetAt(0);
			// detectar colunas como em readClients
			int statusCol = STATUS_COLUMN_DEFAULT;
			int tsCol = TIMESTAMP_COLUMN_DEFAULT;
			Row header = sheet.getRow(0);
			if (header != null) {
				for (Cell cell : header) {
					String h = getCellString(cell);
					if (h == null)
						continue;
					String lower = h.trim().toLowerCase();
					if (lower.contains("status") || lower.contains("situação") || lower.contains("situacao")) {
						statusCol = cell.getColumnIndex();
					}
					if (lower.contains("data") || lower.contains("horário") || lower.contains("horario") || lower.contains("envio") || lower.contains("timestamp")) {
						tsCol = cell.getColumnIndex();
					}
				}
			}

			Row row = sheet.getRow(rowIndex);
			if (row == null) {
				row = sheet.createRow(rowIndex);
			}
			Cell statusCell = row.getCell(statusCol);
			if (statusCell == null)
				statusCell = row.createCell(statusCol);
			statusCell.setCellValue("Enviado");

			Cell tsCell = row.getCell(tsCol);
			if (tsCell == null)
				tsCell = row.createCell(tsCol);
			String now = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date());
			tsCell.setCellValue(now);

			// gravar de volta no arquivo
			try (FileOutputStream fos = new FileOutputStream(path)) {
				wb.write(fos);
			}
			return true;
		} catch (Exception e) {
			System.err.println("Erro ao atualizar status no Excel: " + e.getMessage());
			e.printStackTrace();
			return false;
		}
	}

	private static String getCellString(Cell c) {
		if (c == null)
			return null;
		try {
			if (c.getCellType() == CellType.STRING)
				return c.getStringCellValue().trim();
			if (c.getCellType() == CellType.NUMERIC) {
				// possível CNPJ numérico (sem zeros à esquerda)
				double val = c.getNumericCellValue();
				long asLong = (long) val;
				return String.valueOf(asLong);
			}
			if (c.getCellType() == CellType.BOOLEAN)
				return String.valueOf(c.getBooleanCellValue());
			if (c.getCellType() == CellType.FORMULA)
				return c.getStringCellValue();
		} catch (Exception e) {
		}
		return null;
	}

	private static Integer getCellInteger(Cell c) {
		if (c == null)
			return null;
		try {
			if (c.getCellType() == CellType.NUMERIC)
				return (int) c.getNumericCellValue();
			if (c.getCellType() == CellType.STRING) {
				String s = c.getStringCellValue().replace(".", "").replace(",", ".").trim();
				if (s.isEmpty())
					return null;
				return Integer.parseInt(s);
			}
		} catch (Exception e) {
			return null;
		}
		return null;
	}

	@SuppressWarnings("deprecation")
	private static BigDecimal getCellDecimal(Cell c) {
		if (c == null)
			return null;
		try {
			if (c.getCellType() == CellType.NUMERIC) {
				return BigDecimal.valueOf(c.getNumericCellValue()).setScale(2, BigDecimal.ROUND_HALF_UP);
			}
			if (c.getCellType() == CellType.STRING) {
				String s = c.getStringCellValue().replace(".", "").replace(",", ".").trim();
				if (s.isEmpty())
					return null;
				return new BigDecimal(s).setScale(2, BigDecimal.ROUND_HALF_UP);
			}
		} catch (Exception e) {
			return null;
		}
		return null;
	}

	private static Integer safeInt(Integer i) {
		return i == null ? 0 : i;
	}
}