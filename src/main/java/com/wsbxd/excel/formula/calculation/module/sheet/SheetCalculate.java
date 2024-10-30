package com.wsbxd.excel.formula.calculation.module.sheet;

import com.wsbxd.excel.formula.calculation.common.calculation.formula.ExcelFormula;
import com.wsbxd.excel.formula.calculation.common.interfaces.IExcelCalculate;
import com.wsbxd.excel.formula.calculation.common.config.ExcelCalculateConfig;
import com.wsbxd.excel.formula.calculation.common.config.enums.ExcelCalculateTypeEnum;
import com.wsbxd.excel.formula.calculation.common.util.ExcelStrUtil;
import com.wsbxd.excel.formula.calculation.module.sheet.entity.ExcelSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

/**
 * description: Sheet Calculate
 *
 * @author chenhaoxuan
 * @date 2019/8/28
 */
public class SheetCalculate<T> implements IExcelCalculate {

    private final static Logger logger = LoggerFactory.getLogger(SheetCalculate.class);

    /**
     * excel cells data
     */
    private final ExcelSheet<T> excelSheet;

    /**
     * excel data properties
     */
    private final ExcelCalculateConfig excelCalculateConfig;

    @Override
    public String calculate(String formula) {
        ExcelFormula<T> sheetFormula = new ExcelFormula<>(null, formula, this.excelCalculateConfig);
        String value = sheetFormula.calculate(null, this.excelSheet);
        if (null != sheetFormula.getReturnCell()) {
            value = this.excelSheet.updateExcelCellValue(sheetFormula.getReturnCell());
        }
        return value;
    }

    @Override
    public void integrationResult() {
        this.excelSheet.integrationResult();
    }

    public SheetCalculate(List<T> excelList, ExcelCalculateConfig excelCalculateConfig) {
        this.excelCalculateConfig = excelCalculateConfig;
        this.excelCalculateConfig.setCalculateType(ExcelCalculateTypeEnum.SHEET);
        this.excelSheet = new ExcelSheet<>(excelList, this.excelCalculateConfig);
        // 未完成公式集合
        Set<String> undoFormula = new HashSet<>();
        // 遍历所有单元格，找到含有公式的单元格
        this.excelSheet.getIdAndCellListMap().forEach((rowNum, row) -> {
            row.forEach((columnNum, cell) -> {
                if (ExcelStrUtil.isFormula(cell.getBaseValue())) {
                    // 保存公式，格式为 "坐标 + 公式"
                    String formula = cell.getCoordinate().getCell() + cell.getBaseValue();
                    undoFormula.add(formula);
                }
            });
        });
        // 用于记录最后的异常信息
        String exceptionMsg = null;
        // 最多执行 30 次以处理公式，避免无限循环
        for (int count = 0; count < 30 && !undoFormula.isEmpty(); count++) {
            Iterator<String> iterator = undoFormula.iterator();
            while (iterator.hasNext()) {
                String formula = iterator.next();
                try {
                    // 计算公式
                    calculate(formula);
                    // 计算成功，移除公式
                    iterator.remove();
                } catch (Exception e) {
                    // 捕获异常并记录错误信息
                    exceptionMsg = "Error in formula: " + formula + " - " + e.getMessage();
                }
            }
        }
        // 如果还有未计算完的公式，抛出异常
        if (!undoFormula.isEmpty()) {
            throw new RuntimeException("计算单元格公式异常：" + exceptionMsg);
        }
    }

}
