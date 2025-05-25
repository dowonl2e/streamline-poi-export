package sheet;

import com.dwl.excel.exporter.ExcelExporter;
import com.dwl.excel.exporter.SXSSFExcelExporter;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.MethodOrderer;
import org.junit.jupiter.api.Order;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestMethodOrder;

@TestMethodOrder(MethodOrderer.OrderAnnotation.class)
public class SheetTests {

  @Test
  @Order(1)
  @DisplayName("시트 생성 확인 테스트")
  public void sheetCreateTests(){
    //given
    int sheetCount = 5;
    int flushCount = 100;
    String[] sheetNames = new String[]{"회원","여행","지역","일정","계획"};
    ExcelExporter exporter = SXSSFExcelExporter.builder()
        .flushCount(flushCount)
        .build();

    //when
    for(int i = 0 ; i < sheetCount ; i++){
      exporter.createSheet(sheetNames[i]);
    }
    String[] resultSheetNames = new String[sheetCount];
    for(int i = 0 ; i < exporter.getSheetCount() ; i++){
      resultSheetNames[i] = exporter.getSheetName(i);
    }

    //then
    int expectedCount = 5;
    String[] expectedSheetNames = new String[]{"회원","여행","지역","일정","계획"};

    Assertions.assertEquals(expectedCount, exporter.getSheetCount());
    Assertions.assertArrayEquals(expectedSheetNames, resultSheetNames);
  }
}
