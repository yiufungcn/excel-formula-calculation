package com.wsbxd.excel.formula.calculation.common.prop;

import com.wsbxd.excel.formula.calculation.common.field.annotation.ExcelField;
import com.wsbxd.excel.formula.calculation.common.field.enums.ExcelFieldTypeEnum;
import com.wsbxd.excel.formula.calculation.common.field.enums.ExcelIdTypeEnum;
import com.wsbxd.excel.formula.calculation.common.prop.enums.ExcelCalculateTypeEnum;
import com.wsbxd.excel.formula.calculation.common.util.ExcelStrUtil;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * description: Excel 数据属性
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/27 11:11
 */
public class ExcelDataProperties {

    private final static Pattern CELL_NUMBER_PATTERN = Pattern.compile("[A-Z]+\\d+");

    private final static Pattern CELL_UUID_PATTERN = Pattern.compile("[A-Z]+[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}");

    public final static Pattern SHEET_CELL_NUMBER_PATTERN = Pattern.compile("('[^\\\\/?*\\[\\]]+?'![A-Z]+\\d+|[^\\\\/?*\\[\\]():,+-]+?![A-Z]+\\d+|[A-Z]+\\d+)");

    public final static Pattern SHEET_CELL_UUID_PATTERN = Pattern.compile("('[^\\\\/?*\\[\\]]+?'![A-Z]+[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}|[^\\\\/?*\\[\\]():,+-]+?![A-Z]+[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}|[A-Z]+[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12})");


    private final static Pattern COLUMN_PATTERN = Pattern.compile("(?![A-Z]+\\()([A-Z]+)");

    private final static Pattern RETURN_CELL_NUMBER_PATTERN = Pattern.compile("^[A-Z]+\\d+=");

    private final static Pattern RETURN_CELL_UUID_PATTERN = Pattern.compile("^[A-Z]+[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}=");

    public final static Pattern RETURN_SHEET_CELL_NUMBER_PATTERN = Pattern.compile("^('[^\\\\/?*\\[\\]]+?'![A-Z]+\\d+|[^\\\\/?*\\[\\]():,+-]+?![A-Z]+\\d+|[A-Z]+\\d+)=");

    public final static Pattern RETURN_SHEET_CELL_UUID_PATTERN = Pattern.compile("^('[^\\\\/?*\\[\\]]+?'![A-Z]+[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}|[^\\\\/?*\\[\\]():,+-]+?![A-Z]+[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}|[A-Z]+[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12})=");

    private final static Pattern RETURN_COLUMN_PATTERN = Pattern.compile("^(?![A-Z]+\\()([A-Z]+)=");

    private Field idField;

    private ExcelIdTypeEnum excelIdTypeEnum;

    private Field cellTypesField;

    private Field sheetField;

    private Field sortField;

    private List<Field> columnFieldList;

    private ExcelCalculateTypeEnum calculateType;

    /**
     * 不用注解时
     *
     * @param clazz               excel
     * @param idFieldName         id字段名称
     * @param excelIdTypeEnum     id类型
     * @param sheetFieldName      工作表字段名称
     * @param sortFieldName       排序字段名称
     * @param columnFieldNameList 列字段名称集合
     */
    public ExcelDataProperties(
            ExcelCalculateTypeEnum calculateType, Class<?> clazz, String idFieldName,
            ExcelIdTypeEnum excelIdTypeEnum, String sheetFieldName, String sortFieldName, List<String> columnFieldNameList
    ) {
        try {
            this.calculateType = calculateType;
            this.idField = clazz.getDeclaredField(idFieldName);
            this.excelIdTypeEnum = excelIdTypeEnum;
            this.sortField = clazz.getDeclaredField(sortFieldName);
            if (ExcelStrUtil.isNotBlank(sheetFieldName)) {
                //如果不是工作簿计算是用不到工作表名称的
                this.sheetField = clazz.getDeclaredField(sheetFieldName);
            }
            this.columnFieldList = new ArrayList<>();
            for (String column : columnFieldNameList) {
                this.columnFieldList.add(clazz.getDeclaredField(column));
            }
        } catch (NoSuchFieldException e) {
            e.printStackTrace();
        }
    }

    /**
     * 用注释时
     *
     * @param clazz excel
     */
    public ExcelDataProperties(ExcelCalculateTypeEnum calculateType, Class<?> clazz) {
        this.calculateType = calculateType;
        for (Field field : clazz.getDeclaredFields()) {
            field.setAccessible(true);
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            if (null != excelField) {
                if (ExcelFieldTypeEnum.CELL.name().equals(excelField.value().name())) {
                    if (null == this.columnFieldList) {
                        this.columnFieldList = new ArrayList<>();
                        this.columnFieldList.add(field);
                    } else {
                        this.columnFieldList.add(field);
                    }
                } else if (ExcelFieldTypeEnum.ID.name().equals(excelField.value().name())) {
                    this.idField = field;
                    this.excelIdTypeEnum = excelField.idType();
                } else if (ExcelFieldTypeEnum.CELL_TYPES.name().equals(excelField.value().name())) {
                    this.cellTypesField = field;
                } else if (ExcelFieldTypeEnum.SORT.name().equals(excelField.value().name())) {
                    this.sortField = field;
                } else if (ExcelFieldTypeEnum.SHEET.name().equals(excelField.value().name())) {
                    this.sheetField = field;
                }
            }
        }
    }

    public Pattern getReturnCellPattern() {
        if (ExcelCalculateTypeEnum.BOOK.equals(calculateType)) {
            if (ExcelIdTypeEnum.NUMBER.equals(excelIdTypeEnum)) {
                return RETURN_SHEET_CELL_NUMBER_PATTERN;
            } else if (ExcelIdTypeEnum.UUID.equals(excelIdTypeEnum)) {
                return RETURN_SHEET_CELL_UUID_PATTERN;
            }
        } else if (ExcelCalculateTypeEnum.SHEET.equals(calculateType)) {
            if (ExcelIdTypeEnum.NUMBER.equals(excelIdTypeEnum)) {
                return RETURN_SHEET_CELL_NUMBER_PATTERN;
            } else if (ExcelIdTypeEnum.UUID.equals(excelIdTypeEnum)) {
                return RETURN_SHEET_CELL_UUID_PATTERN;
            }
        } else if (ExcelCalculateTypeEnum.ROW.equals(calculateType)) {
            return RETURN_COLUMN_PATTERN;
        }
        return null;
    }

    public List<String> getCellStrListByFormula(String formula) {
        List<String> cellStrList = new ArrayList<>();
        Matcher matcher = null;
        if (ExcelCalculateTypeEnum.BOOK.equals(calculateType)) {
            if (ExcelIdTypeEnum.NUMBER.equals(excelIdTypeEnum)) {
                matcher = SHEET_CELL_NUMBER_PATTERN.matcher(formula);
            } else if (ExcelIdTypeEnum.UUID.equals(excelIdTypeEnum)) {
                matcher = SHEET_CELL_UUID_PATTERN.matcher(formula);
            }
        } else if (ExcelCalculateTypeEnum.SHEET.equals(calculateType)) {
            if (ExcelIdTypeEnum.NUMBER.equals(excelIdTypeEnum)) {
                matcher = CELL_NUMBER_PATTERN.matcher(formula);
            } else if (ExcelIdTypeEnum.UUID.equals(excelIdTypeEnum)) {
                matcher = CELL_UUID_PATTERN.matcher(formula);
            }
        } else if (ExcelCalculateTypeEnum.ROW.equals(calculateType)) {
            matcher = COLUMN_PATTERN.matcher(formula);
        }
        while (null != matcher && matcher.find()) {
            cellStrList.add(matcher.group());
        }
        return cellStrList;
    }

    public Field getIdField() {
        return idField;
    }

    public void setIdField(Field idField) {
        this.idField = idField;
    }

    public ExcelIdTypeEnum getExcelIdTypeEnum() {
        return excelIdTypeEnum;
    }

    public void setExcelIdTypeEnum(ExcelIdTypeEnum excelIdTypeEnum) {
        this.excelIdTypeEnum = excelIdTypeEnum;
    }

    public Field getCellTypesField() {
        return cellTypesField;
    }

    public void setCellTypesField(Field cellTypesField) {
        this.cellTypesField = cellTypesField;
    }

    public Field getSheetField() {
        return sheetField;
    }

    public void setSheetField(Field sheetField) {
        this.sheetField = sheetField;
    }

    public Field getSortField() {
        return sortField;
    }

    public void setSortField(Field sortField) {
        this.sortField = sortField;
    }

    public List<Field> getColumnFieldList() {
        return columnFieldList;
    }

    public void setColumnFieldList(List<Field> columnFieldList) {
        this.columnFieldList = columnFieldList;
    }

    public ExcelCalculateTypeEnum getCalculateType() {
        return calculateType;
    }

    public void setCalculateType(ExcelCalculateTypeEnum calculateType) {
        this.calculateType = calculateType;
    }

    public ExcelDataProperties(Field idField, ExcelIdTypeEnum excelIdTypeEnum, Field cellTypesField, Field sheetField, Field sortField, List<Field> columnFieldList, ExcelCalculateTypeEnum calculateType) {
        this.idField = idField;
        this.excelIdTypeEnum = excelIdTypeEnum;
        this.cellTypesField = cellTypesField;
        this.sheetField = sheetField;
        this.sortField = sortField;
        this.columnFieldList = columnFieldList;
        this.calculateType = calculateType;
    }

    @Override
    public String toString() {
        return "ExcelDataProperties{" +
                "idField=" + idField +
                ", excelIdTypeEnum=" + excelIdTypeEnum +
                ", cellTypesField=" + cellTypesField +
                ", sheetField=" + sheetField +
                ", sortField=" + sortField +
                ", columnFieldList=" + columnFieldList +
                ", calculateType=" + calculateType +
                '}';
    }
}
