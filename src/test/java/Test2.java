import com.alibaba.excel.EasyExcel;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

public class Test2 {

    /**
     * 最简单的填充
     *
     * @since 2.1.1
     */
    @Test
    public void simpleFill() throws Exception {
        // 此处为模版路径，模板注意 用{} 来表示你要用的变量 如果本来就有"{","}" 特殊字符 用"\{","\}"代替
        String templateFileNamePath = "C:\\Users\\Administrator\\IdeaProjects\\TestEasyExcel\\src\\main\\resources\\book" + ".xlsx";
        //此处为生成文件后的路径
        String fileNamePath = "C:\\Users\\Administrator\\IdeaProjects\\TestEasyExcel\\src\\main\\resources\\book2.xlsx";
        // 这里 会填充到第一个sheet， 然后文件流会自动关闭
        //使用Map进行填充，可以不使用FillData类去创建字段名，感觉不错
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("name", "张三");
        map.put("number", 5);
        map.put("name2", "张4");
        //填充完数据后，模版里已有的公式不会重新计算一次
        EasyExcel.write(fileNamePath).withTemplate(templateFileNamePath).sheet().doFill(map);

        //使模版里的公式计算一次，但是要重新生成一次文件，有没有性能更好的方法？
        Workbook workbook =new XSSFWorkbook(new FileInputStream(new File(fileNamePath)));
        workbook.setForceFormulaRecalculation(true);
        workbook.write(new FileOutputStream(new File(fileNamePath)));
    }


}
