import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.CellData;
import com.alibaba.excel.read.builder.ExcelReaderSheetBuilder;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFFormulaEvaluator;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

public class test1 {

    /**
     * 最简单的填充
     *
     * @since 2.1.1
     */
    @Test
    public void simpleFill() throws Exception {
        // 此处为模版路径，模板注意 用{} 来表示你要用的变量 如果本来就有"{","}" 特殊字符 用"\{","\}"代替
        String templateFileName = "C:\\Users\\Administrator\\IdeaProjects\\TestEasyExcel\\src\\main\\resources\\book" + ".xlsx";

       //此处为生成文件后的路径
        String fileName = "C:\\Users\\Administrator\\IdeaProjects\\TestEasyExcel\\src\\main\\resources\\book2.xlsx";

        String fileName2 = "C:\\Users\\Administrator\\IdeaProjects\\TestEasyExcel\\src\\main\\resources\\book3.xlsx";

        // 这里 会填充到第一个sheet， 然后文件流会自动关闭
        // 方案1 根据对象填充
//        FillData fillData = new FillData();
//        fillData.setName("张三");
//        fillData.setNumber(5.2);
//        EasyExcel.write(fileName).withTemplate(templateFileName).sheet().doFill(fillData);

        //使用Map进行填充，可以不使用FillData类去创建字段名，感觉不错
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("name", "张三");
        map.put("number", 5);
        map.put("name2", "张4");
        EasyExcel.write(fileName).withTemplate(templateFileName).sheet().doFill(map);

       Workbook workbook =new XSSFWorkbook(new FileInputStream(new File(fileName)));
       workbook.setForceFormulaRecalculation(true);

//        Sheet sheet = workbook.getSheetAt(0);
//
//        Row row = sheet.getRow(1);
//        Cell cell = row.getCell(2);
//
//        FormulaEvaluator formulaEvaluator = new SXSSFFormulaEvaluator((SXSSFWorkbook) workbook);
//        String cellFormula = cell.getCellFormula();
//        cell.setCellValue(cellFormula);

        workbook.write(new FileOutputStream(new File(fileName)));





    }


}
