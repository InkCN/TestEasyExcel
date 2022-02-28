


import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.NumberFormat;
import com.alibaba.excel.converters.doubleconverter.DoubleStringConverter;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.Setter;


@Setter
@Getter
@EqualsAndHashCode
public class FillData {
    private String name;
    @NumberFormat("#")
    @ExcelProperty(converter = DoubleStringConverter.class)
    private Double number;
    private String empty;
}