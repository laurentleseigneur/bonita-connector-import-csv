package org.bonitasoft.connectors.excelExport;

import org.bonitasoft.engine.connector.AbstractConnector;
import org.bonitasoft.engine.connector.ConnectorValidationException;

import java.util.List;
import java.util.Optional;

public abstract class AbstractExcelWriteConnectorImpl extends AbstractConnector {

    protected final static String HEADER_INPUT_PARAMETER = "header";
    protected final static String DATA_INPUT_PARAMETER = "data";
    protected final static String SHEET_NAME_INPUT_PARAMETER = "sheetName";
    protected final String EXCELFILE_OUTPUT_PARAMETER = "excelFile";

    protected final List getHeader() {
        return (List) getInputParameter(HEADER_INPUT_PARAMETER);
    }

    protected final String getSheetName() {
        return (String) getInputParameter(SHEET_NAME_INPUT_PARAMETER);
    }

    protected final List getData() {
        return (List) getInputParameter(DATA_INPUT_PARAMETER);
    }

    protected final void setExcelFile(java.io.File excelFile) {
        setOutputParameter(EXCELFILE_OUTPUT_PARAMETER, excelFile);
    }

    @Override
    public void validateInputParameters() throws ConnectorValidationException {
        Optional<String> sheetName = Optional.ofNullable(getSheetName());
        Optional<List> header = Optional.ofNullable(getHeader());
        Optional<List> data = Optional.ofNullable(getData());

        if (!sheetName.isPresent()) {
            throw new ConnectorValidationException("sheetName is required");
        }
        if (!header.isPresent()) {
            throw new ConnectorValidationException("header is required");
        }
        if (!data.isPresent()) {
            throw new ConnectorValidationException("data is required");
        }

    }

}
