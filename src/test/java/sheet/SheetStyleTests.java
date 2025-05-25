package sheet;

import com.dwl.excel.exporter.ExcelExporter;
import com.dwl.excel.exporter.SXSSFExcelExporter;
import com.dwl.excel.style.CellStyleApplier;
import com.dwl.excel.style.font.FontStyleApplier;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.MethodOrderer;
import org.junit.jupiter.api.Order;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestMethodOrder;

@TestMethodOrder(MethodOrderer.OrderAnnotation.class)
public class SheetStyleTests {


  @Test
  @Order(1)
  @DisplayName("시트 스타일 생성 확인 테스트")
  public void sheetStyleCreateTests(){
    //given
    int styleCount = 10;
    CellStyleApplier cellStyler = CellStyleApplier.builder()
        .build();

    ExcelExporter exporter = SXSSFExcelExporter.builder()
        .build();

    //when
    for(int i = 0 ; i < styleCount ; i++){
      exporter.createCellStyle("STYLE_"+(i+1), cellStyler);
    }

    //then
    int expectedCount = styleCount+1;
    Assertions.assertEquals(expectedCount, exporter.getCellStyleCount());
  }

  @Test
  @Order(2)
  @DisplayName("시트 폰트 생성 확인 테스트")
  public void sheetFontCreateTests(){
    //given
    int count = 10;
    FontStyleApplier fontStyler = FontStyleApplier.builder()
        .build();

    ExcelExporter exporter = SXSSFExcelExporter.builder()
        .build();

    //when
    for(int i = 0 ; i < count ; i++){
      exporter.createFont("FONT_"+(i+1), fontStyler);
    }

    //then
    int expectedCount = count+1;
    Assertions.assertEquals(expectedCount, exporter.getFontCount());
  }
}
