package com.wsbxd.excel.formula.calculation;

import com.wsbxd.excel.formula.calculation.common.config.ExcelCalculateConfig;
import com.wsbxd.excel.formula.calculation.module.book.BookCalculate;
import com.wsbxd.excel.formula.calculation.module.row.RowCalculate;
import com.wsbxd.excel.formula.calculation.module.sheet.SheetCalculate;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.ArrayList;
import java.util.List;

/**
 * description: Excel 计算器 测试类
 *
 * @author chenhaoxuan
 * @version 1.0
 * @date 2021/2/27 11:11
 */
@SpringBootTest
class ExcelFormulaCalculationTests {

    /**
     * 唯一标识为数字类型的计算配置
     */
    private final static ExcelCalculateConfig numberExcelCalculateConfig = new ExcelCalculateConfig(ExcelNumber.class);

    /**
     * 唯一标识为UUID类型的计算配置
     */
    private final static ExcelCalculateConfig uuidExcelCalculateConfig = new ExcelCalculateConfig(ExcelUUID.class);

    /**
     * 多页签计算(数字id)
     */
    @Test
    public void bookNum() {
        List<ExcelNumber> excelList = new ArrayList<>();
        excelList.add(new ExcelNumber("1", "200", null, "100000000%", "210", "3.14", "2019/11/15", "5", "44", null, null, "47", "99", "0.05", "0.2", "C-3;F-5", "现金", 1));
        excelList.add(new ExcelNumber("2", "4545", null, "74182", null, "0.0111", null, null, "61", null, "10086", null, "1010%", null, null, "L-3", "现金", 2));
        excelList.add(new ExcelNumber("3", null, "12323", null, null, "1212", null, null, null, "2011/11/01", "0.4", null, null, "0.2", null, "I-5", "交易性-股票", 1));
        excelList.add(new ExcelNumber("4", null, null, "45%", null, null, null, "85296%", null, null, null, "7474", "2565", null, "44", "C-3;G-3", "银行存款", 1));
        excelList.add(new ExcelNumber("5", "2000/01/05", "5%", null, "741", null, "2000/02/05", null, null, "2200/11/05", null, "1900-01-01", null, "777", null, "A-5;B-3;F-5;I-5;K-4", "银行存款", 2));
        System.out.println("-------------------------------------------------");
        long start = System.currentTimeMillis();
        BookCalculate<ExcelNumber> bookCalculate = new BookCalculate<>(excelList, numberExcelCalculateConfig);
        String result = bookCalculate.calculate("交易性-股票", "N3=现金!A1*银行存款!C1+ABS(MIN(O2+2,银行存款!B3:银行存款!E5)*2-MAX('交易性-股票'!A1:'交易性-股票'!D3))+MAX(现金!A2:现金!O3)");
        bookCalculate.integrationResult();
        System.out.println(System.currentTimeMillis() - start + "ms");
        System.out.println(result);
    }

    /**
     * 多页签计算(uuid)
     */
    @Test
    public void bookUUID() {
        List<ExcelUUID> excelList = new ArrayList<>();
        excelList.add(new ExcelUUID("1bbd420f-3ce3-11ea-a676-00163e048143", "200", null, "100000000%", "210", "3.14", "2019/11/15", "5", "44", null, null, "47", "99", "0.05", "0.2", "C-3;F-5", "现金", 1));
        excelList.add(new ExcelUUID("2bbd420f-3ce3-11ea-a676-00163e048143", "4545", null, "74182", null, "0.0111", null, null, "61", null, "10086", null, "1010%", null, null, "L-3", "现金", 2));
        excelList.add(new ExcelUUID("3bbd420f-3ce3-11ea-a676-00163e048143", null, "12323", null, null, "1212", null, null, null, "2011/11/01", "0.4", null, null, "0.2", null, "I-5", "交易性-股票", 1));
        excelList.add(new ExcelUUID("4bbd420f-3ce3-11ea-a676-00163e048143", null, null, "45%", null, null, null, "85296%", null, null, null, "7474", "2565", null, "44", "C-3;G-3", "银行存款", 1));
        excelList.add(new ExcelUUID("5bbd420f-3ce3-11ea-a676-00163e048143", "2000/01/05", "5%", null, "741", null, "2000/02/05", null, null, "2200/11/05", null, "1900-01-01", null, "777", null, "A-5;B-3;F-5;I-5;K-4", "银行存款", 2));
        System.out.println("-------------------------------------------------");
        long start = System.currentTimeMillis();
        BookCalculate<ExcelUUID> bookCalculate = new BookCalculate<>(excelList, uuidExcelCalculateConfig);
        String result = bookCalculate.calculate("交易性-股票", "N3bbd420f-3ce3-11ea-a676-00163e048143=现金!A1bbd420f-3ce3-11ea-a676-00163e048143*银行存款!C1bbd420f-3ce3-11ea-a676-00163e048143+ABS(MIN(O2bbd420f-3ce3-11ea-a676-00163e048143+2,银行存款!B3bbd420f-3ce3-11ea-a676-00163e048143:银行存款!E5bbd420f-3ce3-11ea-a676-00163e048143)*2-MAX('交易性-股票'!A1bbd420f-3ce3-11ea-a676-00163e048143:'交易性-股票'!D3bbd420f-3ce3-11ea-a676-00163e048143))+MAX(现金!A2bbd420f-3ce3-11ea-a676-00163e048143:现金!O3bbd420f-3ce3-11ea-a676-00163e048143)");
        bookCalculate.integrationResult();
        System.out.println(System.currentTimeMillis() - start + "ms");
        System.out.println(result);
    }

    /**
     * 单页签计算(数字id)
     */
    @Test
    public void sheetNum() {
        List<ExcelNumber> excelList = new ArrayList<>();
        excelList.add(new ExcelNumber("1", "200", null, "100000000%", "210", "3.14", "2019/11/15", "5", "44", null, null, "47", "99", "0.05", "0.2", "C-3;F-5", null, 1));
        excelList.add(new ExcelNumber("2", "4545", null, "74182", null, "0.01", null, null, "6341", null, null, null, "1010%", null, null, "L-3", null, 2));
        excelList.add(new ExcelNumber("3", null, "12323", null, null, "1212", null, null, null, "2011/11/01", "0.4", null, null, "0.2", null, "I-5", null, 3));
        excelList.add(new ExcelNumber("4", null, null, "45%", null, null, null, "85296%", null, null, null, "7474", "2565", null, "44", "C-3;G-3", null, 4));
        excelList.add(new ExcelNumber("5", "2000/01/05", "5%", null, "741", null, "2000/02/05", null, null, "2200/11/05", null, null, null, "777", null, "A-5;B-3;F-5;I-5;K-4;N-3", null, 5));
        System.out.println("-------------------------------------------------");
        long start = System.currentTimeMillis();
        SheetCalculate<ExcelNumber> sheetCalculate = new SheetCalculate<>(excelList, numberExcelCalculateConfig);
        String result1 = sheetCalculate.calculate("K5=A1+ABS(MIN(C5+2,A1:I5)*2-MAX(L1:N5))+MAX(A2:J2)");
        String result2 = sheetCalculate.calculate("N5=ROUND(2*K5+5,0)");
        sheetCalculate.integrationResult();
        System.out.println(System.currentTimeMillis() - start + "ms");
        System.out.println(result1);
        System.out.println(result2);
    }

    /**
     * 单页签计算(uuid)
     */
    @Test
    public void sheetUUID() {
        List<ExcelUUID> excelList = new ArrayList<>();
        excelList.add(new ExcelUUID("1bbd420f-3ce3-11ea-a676-00163e048143", "200", null, "100000000%", "210", "3.14", "2019/11/15", "5", "44", null, null, "47", "99", "0.05", "0.2", "C-3;F-5", null, 1));
        excelList.add(new ExcelUUID("2bbd420f-3ce3-11ea-a676-00163e048143", "4545", null, "74182", null, "0.01", null, null, "6341", null, null, null, "1010%", null, null, "L-3", null, 2));
        excelList.add(new ExcelUUID("3bbd420f-3ce3-11ea-a676-00163e048143", null, "12323", null, null, "1212", null, null, null, "2011/11/01", "0.4", null, null, "0.2", null, "I-5", null, 3));
        excelList.add(new ExcelUUID("4bbd420f-3ce3-11ea-a676-00163e048143", null, null, "45%", null, null, null, "85296%", null, null, null, "7474", "2565", null, "44", "C-3;G-3", null, 4));
        excelList.add(new ExcelUUID("5bbd420f-3ce3-11ea-a676-00163e048143", "2000/01/05", "5%", null, "741", null, "2000/02/05", null, null, "2200/11/05", null, null, null, "777", null, "A-5;B-3;F-5;I-5;K-4;N-3", null, 5));
        System.out.println("-------------------------------------------------");
        long start = System.currentTimeMillis();
        SheetCalculate<ExcelUUID> sheetCalculate = new SheetCalculate<>(excelList, uuidExcelCalculateConfig);
        String result1 = sheetCalculate.calculate("K5bbd420f-3ce3-11ea-a676-00163e048143=A1bbd420f-3ce3-11ea-a676-00163e048143+ABS(MIN(C5bbd420f-3ce3-11ea-a676-00163e048143+2,A1bbd420f-3ce3-11ea-a676-00163e048143:I5bbd420f-3ce3-11ea-a676-00163e048143)*2-MAX(L1bbd420f-3ce3-11ea-a676-00163e048143:N5bbd420f-3ce3-11ea-a676-00163e048143))+MAX(A2bbd420f-3ce3-11ea-a676-00163e048143:J2bbd420f-3ce3-11ea-a676-00163e048143)");
        String result2 = sheetCalculate.calculate("N5bbd420f-3ce3-11ea-a676-00163e048143=ROUND(2*K5bbd420f-3ce3-11ea-a676-00163e048143+5,0)");
        sheetCalculate.integrationResult();
        System.out.println(System.currentTimeMillis() - start + "ms");
        System.out.println(result1);
        System.out.println(result2);
    }

    /**
     * 单行计算(数字id)
     */
    @Test
    public void rowNum() {
        ExcelNumber excel = new ExcelNumber(null, "200", null, "100000000%", "210", "3.14", "2019/11/15", "5", "44", null, null, "47", "99", "0.05", "0.2", "C-3;F-5", null, null);
        long start = System.currentTimeMillis();
        RowCalculate<ExcelNumber> calculateHelper = new RowCalculate<>(excel, numberExcelCalculateConfig);
        String result1 = calculateHelper.calculate("I=A+ABS(MIN(O+2,B)*2-MAX(A:H))+MAX(K:N)");
        String result2 = calculateHelper.calculate("J=ROUND(2*I+5,0)");
        calculateHelper.integrationResult();
        System.out.println("-------------------------------------------------");
        System.out.println(System.currentTimeMillis() - start + "ms");
        System.out.println(result1);
        System.out.println(result2);
    }

    /**
     * 单行计算(uuid)
     */
    @Test
    public void rowUUID() {
        ExcelUUID excel = new ExcelUUID(null, "200", null, "100000000%", "210", "3.14", "2019/11/15", "5", "44", null, null, "47", "99", "0.05", "0.2", "C-3;F-5", null, null);
        long start = System.currentTimeMillis();
        RowCalculate<ExcelUUID> calculateHelper = new RowCalculate<>(excel, uuidExcelCalculateConfig);
        String result1 = calculateHelper.calculate("I=A+ABS(MIN(O+2,B)*2-MAX(A:H))+MAX(K:N)");
        String result2 = calculateHelper.calculate("J=ROUND(2*I+5,0)");
        calculateHelper.integrationResult();
        System.out.println("-------------------------------------------------");
        System.out.println(System.currentTimeMillis() - start + "ms");
        System.out.println(result1);
        System.out.println(result2);
    }

    /**
     * 测试单页复杂公式以及嵌套
     */
    @Test
    public void sheetNumNestedFormulaTest() {
        List<ExcelNumber> excelList = new ArrayList<>();
        excelList.add(new ExcelNumber("1", "11", "22", "33", "11", "=SUM(A1:D1)", null, null, null, null, null, null, null, null, null, null, null, 1));
        excelList.add(new ExcelNumber("2", "111", "222", "333", "111", "=SUM(A2+B2+C2+D2)", null, null, null, null, null, null, null, null, null, null, null, 2));
        excelList.add(new ExcelNumber("3", "1.111", "2.222", "3.333", "9.999", "=SUM(A3:D3)", "=SUM(E2-E1)", null, null, null, null, null, null, null, null, null, null, 3));
        excelList.add(new ExcelNumber("4", "-1111111111111111111111.111", "2222222222222222222222.222", "3333333333333333333333.333", "1111111111111111111111.111", "=SUM(A4+B4+C4+D4)", "=ABS(A4)", null, null, null, null, null, null, null, null, null, null, 4));
        excelList.add(new ExcelNumber("5", "=SUM(A1:A4)", "=MAX(B1:B4)", "=IF(SUM(C1:C4) > SUM(D1:D4), SUM(D1:D4), SUM(C1:C4))", "=SUM(D1:D4)", "=ROUND(E4, 2)", "", null, null, null, null, null, null, null, null, null, null, 5));
        System.out.println("-------------------------------------------------");
        long start = System.currentTimeMillis();
        SheetCalculate<ExcelNumber> sheetCalculate = new SheetCalculate<>(excelList, numberExcelCalculateConfig);

        String A5 = sheetCalculate.calculate("A5");
        String B5 = sheetCalculate.calculate("B5");
        String C5 = sheetCalculate.calculate("C5");
        String D5 = sheetCalculate.calculate("D5");
        String E1 = sheetCalculate.calculate("E1");
        String E2 = sheetCalculate.calculate("E2");
        String E3 = sheetCalculate.calculate("E3");
        String E4 = sheetCalculate.calculate("E4");
        String E5 = sheetCalculate.calculate("E5");
        String F3 = sheetCalculate.calculate("F3");
        String F4 = sheetCalculate.calculate("F4");
        String F5 = sheetCalculate.calculate("F5");


        String F1 = sheetCalculate.calculate("F1=A1+ABS(MIN(C5,A1:C4)*2-MAX(A1:A4))+MAX(A1:D4)");

        sheetCalculate.integrationResult();
        System.out.println("耗时：" + (System.currentTimeMillis() - start) + "ms");
        System.out.println("A5【=SUM(A1:A4)】:" + A5);
        System.out.println("B5【=MAX(B1:B4)】:" + B5);
        System.out.println("C5【=IF(SUM(C1:C4) > SUM(D1:D4), SUM(D1:D4), SUM(C1:C4))】:" + C5);
        System.out.println("D5【=ROUND(SUM(E1:E4), 2)】:" + D5);
        System.out.println("E1【=SUM(A1:D1)】:" + E1);
        System.out.println("E2【=SUM(A2+B2+C2+D2)】:" + E2);
        System.out.println("E3【=SUM(A3:D3)】:" + E3);
        System.out.println("E4【=SUM(A4+B4+C4+D4)】:" + E4);
        System.out.println("E5【=ROUND(E4, 2)】:" + E5);
        System.out.println("F3【=SUM(E2-E1)】:" + F3);
        System.out.println("F4【=ABS(A4)】:" + F4);
        System.out.println("F5【=MIN(C5,A1:C4)】:" + F5);
        System.out.println("F1【F1=A1+ABS(MIN(C5,A1:C4)*2-MAX(A1:A4))+MAX(A1:D4)】:" + F1);

    }

}
