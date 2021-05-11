package resource.data;

import annotation.AllCell;
import annotation.Cell;
import annotation.CustomCellStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.xmlbeans.impl.xb.xsdschema.All;
import resource.style.ExcelCellStyle;
import resource.style.NoExcelCellStyle;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelDataFactory {
    private final int HEADER = 0;
    private final int BODY = 1;

    public ExcelData getExcelData(Class<?> clazz, Workbook wb) {
        List<String> fieldName = new ArrayList<>();
        Map<String, String> headerMap = new HashMap<>();
        Map<String, CellStyle> headerStyleMap = new HashMap<>();
        Map<String, CellStyle> bodyStyleMap = new HashMap<>();
        CellStyle defaultHeaderStyle = getDefaultCellStyle(clazz, wb, HEADER);
        CellStyle defaultBodyStyle = getDefaultCellStyle(clazz, wb, BODY);

        for(Field field : clazz.getDeclaredFields()) {
            if(field.isAnnotationPresent(Cell.class)) {
                Cell cell = field.getAnnotation(Cell.class);

                fieldName.add(field.getName());
                headerMap.put(field.getName(), cell.headerName());

                CustomCellStyle headerCellStyle = cell.headerStyle();
                CustomCellStyle bodyCellStyle = cell.bodyStyle();

                headerStyleMap.put(field.getName(), getCellStyle(clazz, headerCellStyle, wb, HEADER));
                bodyStyleMap.put(field.getName(), getCellStyle(clazz, bodyCellStyle, wb, BODY));
            }
        }

        return new ExcelData(fieldName, headerMap, headerStyleMap, bodyStyleMap, defaultHeaderStyle, defaultBodyStyle);
    }

    private CellStyle getCellStyle(Class<?> clazz, CustomCellStyle customCellStyle, Workbook wb, int location) {
        CellStyle cellStyle = wb.createCellStyle();
        try{
            if(isCustom(customCellStyle)) {
                ExcelCellStyle excelCellStyle = customCellStyle.excelCellStyle().newInstance();
                excelCellStyle.styleApply(cellStyle);
            } else {
                AllCell allCell = clazz.getAnnotation(AllCell.class);
                ExcelCellStyle excelCellStyle;

                if(location == HEADER) excelCellStyle = allCell.headerStyle().excelCellStyle().newInstance();
                else excelCellStyle = allCell.bodyStyle().excelCellStyle().newInstance();

                excelCellStyle.styleApply(cellStyle);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return cellStyle;
    }

    private CellStyle getDefaultCellStyle(Class<?> clazz, Workbook wb, int location) {
        CellStyle cellStyle = wb.createCellStyle();

        try {
            AllCell allCell = clazz.getAnnotation(AllCell.class);
            ExcelCellStyle excelCellStyle;

            if (location == HEADER) excelCellStyle = allCell.headerStyle().excelCellStyle().newInstance();
            else excelCellStyle = allCell.bodyStyle().excelCellStyle().newInstance();

            excelCellStyle.styleApply(cellStyle);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return cellStyle;
    }

    private boolean isCustom(CustomCellStyle customCellStyle) {
        if(customCellStyle.excelCellStyle().equals(NoExcelCellStyle.class)) {
            return false;
        }
        return true;
    }
}
