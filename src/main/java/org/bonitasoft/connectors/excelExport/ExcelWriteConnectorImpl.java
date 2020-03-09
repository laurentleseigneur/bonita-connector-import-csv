package org.bonitasoft.connectors.excelExport;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bonitasoft.engine.connector.ConnectorException;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;


/**
 * The connector execution will follow the steps
 * 1 - setInputParameters() --> the connector receives input parameters values
 * 2 - validateInputParameters() --> the connector can validate input parameters values
 * 3 - connect() --> the connector can establish a connection to a remote server (if necessary)
 * 4 - executeBusinessLogic() --> execute the connector
 * 5 - getOutputParameters() --> output are retrieved from connector
 * 6 - disconnect() --> the connector can close connection to remote server (if any)
 */
public class ExcelWriteConnectorImpl extends AbstractExcelWriteConnectorImpl {

	@Override
	protected void executeBusinessLogic() throws ConnectorException {
		//Get access to the connector input parameters
		List<String> headers = getHeader();
		List<List<Object>> data = getData();

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet(getSheetName());
		AtomicInteger rowCount = new AtomicInteger();
		Row row = sheet.createRow(rowCount.getAndIncrement());
		AtomicInteger columnCount = new AtomicInteger();
		for (String header1 : headers) {
			Cell cell1 = row.createCell(columnCount.getAndIncrement());
			cell1.setCellValue(header1);
		}

		for (List<Object> dataLine : data) {
			row = sheet.createRow(rowCount.getAndIncrement());
			columnCount = new AtomicInteger();
			for (Object cellData : dataLine) {
				int countAndIncrement = columnCount.getAndIncrement();
				if (cellData != null) {
					Cell cell = row.createCell(countAndIncrement);
					String cellClass = cellData.getClass().getSimpleName();
					switch (cellClass) {
						case "String":
							cell.setCellValue((String) cellData);
							break;
						case "Double":
							cell.setCellValue((Double) cellData);
							break;
						default:
							throw new ConnectorException("cell with type " + cellClass + " is not yet supported." +
									" Consider using a String");
					}
				}


			}
		}

		try {
			File excelFile = null;
			excelFile = File.createTempFile("bonita", "");
			try (FileOutputStream outputStream = new FileOutputStream(excelFile)) {
				workbook.write(outputStream);
				workbook.close();
				setExcelFile(excelFile);
			}
		} catch (IOException e) {
			throw new ConnectorException(e);
		}
	}

	@Override
	public void connect() throws ConnectorException {
		//[Optional] Open a connection to remote server

	}

	@Override
	public void disconnect() throws ConnectorException {
		//[Optional] Close connection to remote server

	}

}
