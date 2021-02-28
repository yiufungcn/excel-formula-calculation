package com.wsbxd.excel.formula.calculation.module.book;

import com.wsbxd.excel.formula.calculation.common.prop.ExcelDataProperties;
import com.wsbxd.excel.formula.calculation.common.prop.enums.ExcelCalculateTypeEnum;
import com.wsbxd.excel.formula.calculation.module.book.entity.ExcelBook;
import com.wsbxd.excel.formula.calculation.module.interfaces.abstracts.AbstractsCalculate;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

/**
 * description: Excel 工作簿 计算器
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/27 11:11
 */
public class BookCalculate<T> extends AbstractsCalculate {

    private final static Logger logger = LoggerFactory.getLogger(BookCalculate.class);

    /**
     * excel book
     */
    private final ExcelBook<T> excelBook;

    /**
     * excel data properties
     */
    private final ExcelDataProperties properties;

    public void integrationResult() {
        this.excelBook.integrationResult();
    }

    @Override
    public String calculateNotChangeValue(String formula) {
        return calculateNotChangeValue(null, formula);
    }

    public String calculateNotChangeValue(String currentSheet, String formula) {
        BookFormula<T> bookFormula = new BookFormula<>(currentSheet, formula, this.properties);
        return bookFormula.calculate(currentSheet, this.excelBook);
    }

    @Override
    public String calculateChangeValue(String formula) {
        return calculateChangeValue(null, formula);
    }

    public String calculateChangeValue(String currentSheet, String formula) {
        BookFormula<T> bookFormula = new BookFormula<>(currentSheet, formula, this.properties);
        String value = bookFormula.calculate(currentSheet, this.excelBook);
        if (null != bookFormula.getReturnCell()) {
            this.excelBook.updateExcelCellValue(bookFormula.getReturnCell());
        }
        return value;
    }

    public BookCalculate(List<T> excelList, Class<T> tClass) {
        this.properties = new ExcelDataProperties(ExcelCalculateTypeEnum.BOOK, tClass);
        this.excelBook = new ExcelBook<>(excelList, this.properties);
    }

}
