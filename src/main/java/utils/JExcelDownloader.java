package utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import resource.data.ExcelData;
import resource.data.ExcelDataFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.List;

public class JExcelDownloader {
    private int ROW_INDEX = 0;

    private Workbook workbook;
    private Sheet sheet;
    private ExcelData excelData;

    public JExcelDownloader() {
        this.workbook = new XSSFWorkbook();
        this.sheet = this.workbook.createSheet("Sheet");
    }

    public void excelDownload(List<?> data, Class<?> clazzType, HttpServletResponse response, String fileName, boolean useSeq) throws Exception {
        ExcelDataFactory excelDataFactory = new ExcelDataFactory();
        this.excelData = excelDataFactory.getExcelData(clazzType, this.workbook);

        setHttpHeader(response, fileName);

        renderHeader(clazzType, useSeq);
        renderBody(data, clazzType, useSeq);

        write(response.getOutputStream());
    }

    private void setHttpHeader(HttpServletResponse response, String fileName) {
        response.setHeader("Content-Disposition", "attachment; filename=" + fileName + ".xlsx");
        response.setContentType("application/download;charset=utf-8");
        response.setHeader("Content-Transfer-Encoding", "binary");
    }

    private void renderHeader(Class<?> clazzType, boolean useSeq) throws Exception {
        Row row = this.sheet.createRow(ROW_INDEX++);

        int col = 0;
        for(Field field : clazzType.getDeclaredFields()) {
            if(field.isAnnotationPresent(annotation.Cell.class)) {
                if (col == 0 && useSeq) {
                    Cell cell = row.createCell(col++);
                    setCellStyle(cell, excelData.getDefaultHeaderStyle());
                    setCellValue(cell, "No");
                }
                String headerName = excelData.getHeaderMap().get(field.getName());

                Cell cell = row.createCell(col++);
                setCellStyle(cell, excelData.getHeaderStyleMap().get(field.getName()));
                setCellValue(cell, headerName);
            }
        }
    }

    private void renderBody(List<?> data, Class<?> clazzType, boolean useSeq) throws Exception {
        for(Object object : data) {
            Row row = this.sheet.createRow(ROW_INDEX++);

            int col = 0;
            for(String fieldName : excelData.getFieldName()) {
                if(col==0 && useSeq) {
                    Cell cell = row.createCell(col++);
                    setCellStyle(cell, excelData.getDefaultBodyStyle());
                    setCellValue(cell, Integer.toString(ROW_INDEX-1));
                }
                Field field = getField(clazzType, fieldName);
                field.setAccessible(true);
                Object cellValue = field.get(object);

                Cell cell = row.createCell(col++);
                setCellStyle(cell, excelData.getBodyStyleMap().get(field.getName()));
                setCellValue(cell, (String) cellValue);
            }
        }
    }

    private void setCellStyle(Cell cell, CellStyle cellStyle) {
        cell.setCellStyle(cellStyle);
    }

    private void setCellValue(Cell cell, String value) {
        cell.setCellValue(value);
    }

    private void write(OutputStream outputStream) throws IOException {
        this.workbook.write(outputStream);
        outputStream.close();
    }

    private Field getField(Class<?> clazz, String fieldName) {
        for(Field field : clazz.getDeclaredFields()) {
            if(field.getName().equals(fieldName)) {
                return field;
            }
        }
        return null;
    }
}
