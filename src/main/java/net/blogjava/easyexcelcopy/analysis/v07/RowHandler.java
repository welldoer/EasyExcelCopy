package net.blogjava.easyexcelcopy.analysis.v07;

import java.util.Arrays;
import java.util.List;

import org.apache.poi.xssf.model.SharedStringsTable;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import net.blogjava.easyexcelcopy.annotation.FieldType;
import net.blogjava.easyexcelcopy.constant.ExcelXmlConstants;
import net.blogjava.easyexcelcopy.context.IAnalysisContext;
import net.blogjava.easyexcelcopy.event.IAnalysisEventRegisterCenter;
import net.blogjava.easyexcelcopy.event.OneRowAnalysisFinishEvent;
import net.blogjava.easyexcelcopy.util.PositionUtils;

import static net.blogjava.easyexcelcopy.constant.ExcelXmlConstants.*;

public class RowHandler extends DefaultHandler {

    private String currentCellIndex;

    private FieldType currentCellType;

    private int curRow;

    private int curCol;

    private String[] curRowContent = new String[20];

    private String currentCellValue;

    private SharedStringsTable sst;

    private IAnalysisContext analysisContext;

    private IAnalysisEventRegisterCenter registerCenter;

    private List<String> sharedStringList;

    public RowHandler(IAnalysisEventRegisterCenter registerCenter, SharedStringsTable sst,
                      IAnalysisContext analysisContext, List<String> sharedStringList) {
        this.registerCenter = registerCenter;
        this.analysisContext = analysisContext;
        this.sst = sst;
        this.sharedStringList = sharedStringList;

    }

    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {

        setTotalRowCount(name, attributes);

        startCell(name, attributes);

        startCellValue(name);

    }

    private void startCellValue(String name) {
        if (name.equals(CELL_VALUE_TAG) || name.equals(CELL_VALUE_TAG_1)) {
            // initialize current cell value
            currentCellValue = "";
        }
    }

    private void startCell(String name, Attributes attributes) {
        if (ExcelXmlConstants.CELL_TAG.equals(name)) {
            currentCellIndex = attributes.getValue(ExcelXmlConstants.POSITION);
            int nextRow = PositionUtils.getRow(currentCellIndex);
            if (nextRow > curRow) {
                curRow = nextRow;
                // endRow(ROW_TAG);
            }
            analysisContext.setCurrentRowNum(curRow);
            curCol = PositionUtils.getCol(currentCellIndex);

            String cellType = attributes.getValue("t");
            currentCellType = FieldType.EMPTY;
            if (cellType != null && cellType.equals("s")) {
                currentCellType = FieldType.STRING;
            }
            //if ("6".equals(attributes.getValue("s"))) {
            //    // date
            //    currentCellType = FieldType.DATE;
            //}

        }
    }

    private void endCellValue(String name) throws SAXException {
        // ensure size
        if (curCol >= curRowContent.length) {
            curRowContent = Arrays.copyOf(curRowContent, (int)(curCol * 1.5));
        }
        if (CELL_VALUE_TAG.equals(name)) {

            switch (currentCellType) {
                case STRING:
                    int idx = Integer.parseInt(currentCellValue);
                    currentCellValue = sharedStringList.get(idx);
                    currentCellType = FieldType.EMPTY;
                    break;
                //case DATE:
                //    Date dateVal = HSSFDateUtil.getJavaDate(Double.parseDouble(currentCellValue),
                //        analysisContext.use1904WindowDate());
                //    currentCellValue = TypeUtil.getDefaultDateString(dateVal);
                //    currentCellType = FieldType.EMPTY;
                //    break;
            }
            curRowContent[curCol] = currentCellValue;
        } else if (CELL_VALUE_TAG_1.equals(name)) {
            curRowContent[curCol] = currentCellValue;
        }
    }

    @Override
    public void endElement(String uri, String localName, String name) throws SAXException {

        endRow(name);
        endCellValue(name);
    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {

        currentCellValue += new String(ch, start, length);

    }

    private void setTotalRowCount(String name, Attributes attributes) {
        if (DIMENSION.equals(name)) {
            String d = attributes.getValue(DIMENSION_REF);
            String totalStr = d.substring(d.indexOf(":") + 1, d.length());
            String c = totalStr.toUpperCase().replaceAll("[A-Z]", "");
            analysisContext.setTotalCount(Integer.parseInt(c));
        }

    }

    private void endRow(String name) {
        if (name.equals(ROW_TAG)) {
            registerCenter.notifyListeners(new OneRowAnalysisFinishEvent(Arrays.asList(curRowContent)));
            curRowContent = new String[20];
        }
    }

}
