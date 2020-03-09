package org.bonitasoft.connectors.excelExport;

import org.bonitasoft.engine.connector.ConnectorException;
import org.bonitasoft.engine.connector.ConnectorValidationException;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import static org.assertj.core.api.Assertions.*;
import static org.bonitasoft.connectors.excelExport.AbstractExcelWriteConnectorImpl.*;

public class ExcelWriteConnectorImplTest {

    @Test(expected = ConnectorValidationException.class)
    public void shouldRaiseErrorWhenSheetNameIsNotSet() throws ConnectorException, IOException, ConnectorValidationException {
        //given
        ExcelWriteConnectorImpl excelWriteConnector = new ExcelWriteConnectorImpl();
        Map<String, Object> params = new HashMap<>();
        excelWriteConnector.setInputParameters(params);

        //when
        excelWriteConnector.validateInputParameters();

        //then exception

    }

    @Test(expected = ConnectorValidationException.class)
    public void shouldRaiseErrorWhenHeaderIsNotSet() throws ConnectorException, IOException, ConnectorValidationException {
        //given
        ExcelWriteConnectorImpl excelWriteConnector = new ExcelWriteConnectorImpl();
        Map<String, Object> params = new HashMap<>();
        params.put(SHEET_NAME_INPUT_PARAMETER, "sheet");
        params.put(HEADER_INPUT_PARAMETER, null);
        params.put(DATA_INPUT_PARAMETER, Arrays.asList("a"));


        excelWriteConnector.setInputParameters(params);

        //when
        excelWriteConnector.validateInputParameters();

        //then exception

    }

    @Test(expected = ConnectorValidationException.class)
    public void shouldRaiseErrorWhenDataIsNotSet() throws ConnectorException, IOException, ConnectorValidationException {
        //given
        ExcelWriteConnectorImpl excelWriteConnector = new ExcelWriteConnectorImpl();
        Map<String, Object> params = new HashMap<>();
        params.put(SHEET_NAME_INPUT_PARAMETER, "sheet");
        params.put(HEADER_INPUT_PARAMETER, Arrays.asList("a"));

        excelWriteConnector.setInputParameters(params);

        //when
        excelWriteConnector.validateInputParameters();

        //then exception

    }

    @Test
    public void shouldGenerateExcelFileWithStringData() throws ConnectorException, IOException, ConnectorValidationException {
        //given
        ExcelWriteConnectorImpl excelWriteConnector = new ExcelWriteConnectorImpl();
        Map<String, Object> params = new HashMap<>();
        params.put("sheetName", "sheetName");
        params.put("header", Arrays.asList("a", "b"));
        params.put("data", Arrays.asList(Arrays.asList("1", "2"), Arrays.asList("3", "4")));
        excelWriteConnector.setInputParameters(params);
        File outFile = File.createTempFile("test", ".xlsx" +
                "");
        excelWriteConnector.setExcelFile(outFile);

        //when
        excelWriteConnector.validateInputParameters();
        excelWriteConnector.executeBusinessLogic();

        //then
        assertThat(outFile.exists());
    }

    @Test
    public void shouldHandleIntegerData() throws ConnectorException, IOException, ConnectorValidationException {
        //given
        ExcelWriteConnectorImpl excelWriteConnector = new ExcelWriteConnectorImpl();
        Map<String, Object> params = new HashMap<>();
        params.put("sheetName", "sheetName");
        params.put("header", Arrays.asList("a", "b"));
        params.put("data", Arrays.asList(Arrays.asList(2)));
        excelWriteConnector.setInputParameters(params);
        File outFile = File.createTempFile("test", ".xlsx" +
                "");
        excelWriteConnector.setExcelFile(outFile);

        //when
        excelWriteConnector.validateInputParameters();
        excelWriteConnector.executeBusinessLogic();

        //then
        assertThat(outFile.exists());
    }

    @Test
    public void shouldHandleLongData() throws ConnectorException, IOException, ConnectorValidationException {
        //given
        ExcelWriteConnectorImpl excelWriteConnector = new ExcelWriteConnectorImpl();
        Map<String, Object> params = new HashMap<>();
        params.put("sheetName", "sheetName");
        params.put("header", Arrays.asList("a", "b"));
        params.put("data", Arrays.asList(Arrays.asList(2L)));
        excelWriteConnector.setInputParameters(params);
        File outFile = File.createTempFile("test", ".xlsx" +
                "");
        excelWriteConnector.setExcelFile(outFile);

        //when
        excelWriteConnector.validateInputParameters();
        excelWriteConnector.executeBusinessLogic();

        //then
        assertThat(outFile.exists());
    }

    @Test
    public void shouldHandleDateData() throws ConnectorException, IOException, ConnectorValidationException {
        //given
        ExcelWriteConnectorImpl excelWriteConnector = new ExcelWriteConnectorImpl();
        Map<String, Object> params = new HashMap<>();
        params.put("sheetName", "sheetName");
        params.put("header", Arrays.asList("a", "b"));
        params.put("data", Arrays.asList(Arrays.asList(new Date())));
        excelWriteConnector.setInputParameters(params);
        File outFile = File.createTempFile("test", ".xlsx" +
                "");
        excelWriteConnector.setExcelFile(outFile);

        //when
        excelWriteConnector.validateInputParameters();
        excelWriteConnector.executeBusinessLogic();

        //then
        assertThat(outFile.exists());
    }

    @Test
    public void shouldHandleDoubleData() throws ConnectorException, IOException, ConnectorValidationException {
        //given
        ExcelWriteConnectorImpl excelWriteConnector = new ExcelWriteConnectorImpl();
        Map<String, Object> params = new HashMap<>();
        params.put("sheetName", "sheetName");
        params.put("header", Arrays.asList("a", "b"));
        params.put("data", Arrays.asList(Arrays.asList(2D)));
        excelWriteConnector.setInputParameters(params);
        File outFile = File.createTempFile("test", ".xlsx" +
                "");
        excelWriteConnector.setExcelFile(outFile);

        //when
        excelWriteConnector.validateInputParameters();
        excelWriteConnector.executeBusinessLogic();

        //then
        assertThat(outFile.exists());
    }

    @Test
    public void shouldHandleFloatData() throws ConnectorException, IOException, ConnectorValidationException {
        //given
        ExcelWriteConnectorImpl excelWriteConnector = new ExcelWriteConnectorImpl();
        Map<String, Object> params = new HashMap<>();
        params.put("sheetName", "sheetName");
        params.put("header", Arrays.asList("a", "b"));
        params.put("data", Arrays.asList(Arrays.asList(2)));
        excelWriteConnector.setInputParameters(params);
        File outFile = File.createTempFile("test", ".xlsx" +
                "");
        excelWriteConnector.setExcelFile(outFile);

        //when
        excelWriteConnector.validateInputParameters();
        excelWriteConnector.executeBusinessLogic();

        //then
        assertThat(outFile.exists());
    }

    @Test
    public void shouldHandleLocalDateData() throws ConnectorException, IOException, ConnectorValidationException {
        //given
        ExcelWriteConnectorImpl excelWriteConnector = new ExcelWriteConnectorImpl();
        Map<String, Object> params = new HashMap<>();
        params.put("sheetName", "sheetName");
        params.put("header", Arrays.asList("a", "b"));
        params.put("data", Arrays.asList(Arrays.asList(LocalDate.now())));
        excelWriteConnector.setInputParameters(params);
        File outFile = File.createTempFile("test", ".xlsx" +
                "");
        excelWriteConnector.setExcelFile(outFile);

        //when
        excelWriteConnector.validateInputParameters();
        excelWriteConnector.executeBusinessLogic();

        //then
        assertThat(outFile.exists());
    }

    @Test
    public void shouldHandleLocalDateTimeData() throws ConnectorException, IOException, ConnectorValidationException {
        //given
        ExcelWriteConnectorImpl excelWriteConnector = new ExcelWriteConnectorImpl();
        Map<String, Object> params = new HashMap<>();
        params.put("sheetName", "sheetName");
        params.put("header", Arrays.asList("a", "b"));
        params.put("data", Arrays.asList(Arrays.asList(LocalDateTime.now())));
        excelWriteConnector.setInputParameters(params);
        File outFile = File.createTempFile("test", ".xlsx" +
                "");
        excelWriteConnector.setExcelFile(outFile);

        //when
        excelWriteConnector.validateInputParameters();
        excelWriteConnector.executeBusinessLogic();

        //then
        assertThat(outFile.exists());
    }

    @Test
    public void shouldHandleOffsetDateTimeData() throws ConnectorException, IOException, ConnectorValidationException {
        //given
        ExcelWriteConnectorImpl excelWriteConnector = new ExcelWriteConnectorImpl();
        Map<String, Object> params = new HashMap<>();
        params.put("sheetName", "sheetName");
        params.put("header", Arrays.asList("a", "b"));
        params.put("data", Arrays.asList(Arrays.asList(OffsetDateTime.now())));
        excelWriteConnector.setInputParameters(params);
        File outFile = File.createTempFile("test", ".xlsx" +
                "");
        excelWriteConnector.setExcelFile(outFile);

        //when
        excelWriteConnector.validateInputParameters();
        excelWriteConnector.executeBusinessLogic();

        //then
        assertThat(outFile.exists());
    }

    @Test(expected = ConnectorException.class)
    public void shouldRaiseExceptionForUnsupportedDataType() throws ConnectorException, IOException, ConnectorValidationException {
        //given
        ExcelWriteConnectorImpl excelWriteConnector = new ExcelWriteConnectorImpl();
        Map<String, Object> params = new HashMap<>();
        params.put("sheetName", "sheetName");
        params.put("header", Arrays.asList("a", "b"));

        params.put("data", Arrays.asList(Arrays.asList(new Exotic())));
        excelWriteConnector.setInputParameters(params);
        File outFile = File.createTempFile("test", ".xlsx" +
                "");
        excelWriteConnector.setExcelFile(outFile);

        //when
        excelWriteConnector.validateInputParameters();
        excelWriteConnector.executeBusinessLogic();

        //then
        assertThat(outFile.exists());
    }

    class Exotic {

    }
}