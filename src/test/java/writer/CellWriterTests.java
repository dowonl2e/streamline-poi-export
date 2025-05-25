package writer;

import com.dwl.excel.exporter.ExcelExporter;
import com.dwl.excel.exporter.SXSSFExcelExporter;
import com.dwl.excel.exporter.writer.CellWriter;
import dto.CustomExportableDto;
import dto.CustomExportableDto2;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.MethodOrderer;
import org.junit.jupiter.api.Order;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestMethodOrder;

import java.security.SecureRandom;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

@TestMethodOrder(MethodOrderer.OrderAnnotation.class)
public class CellWriterTests {

  @Test
  @Order(1)
  @DisplayName("셀 행/열 인덱스 테스트")
  public void cellRowColIndexTest() throws Exception {
    //given
    int count = 10;
    String[] headers = new String[]{"헤더1","헤더2","헤더3","헤더4","헤더5"};
    ExcelExporter exporter = SXSSFExcelExporter.builder()
        .build();
    exporter.createSheet("시트1");

    //when
    exporter.createHeader().writeCells(headers);
    CellWriter writer = null;
    for(int i = 0 ; i < count ; i++){
      writer = exporter.createRow()
          .writeCell((i+1))
          .writeCell("데이터_" + (i+1))
          .writeCell((i%2 == 0))
          .writeCell(LocalDate.now())
          .writeCell(LocalDateTime.now());
    }

    //then
    int resultRow = 11;
    int resultCol = 5;
    Assertions.assertEquals(resultRow, exporter.getCurrentRow());
    Assertions.assertEquals(resultCol, writer.getCurrentCol());
  }

  @Test
  @Order(2)
  @DisplayName("셀 행/열(배열) 인덱스 테스트")
  public void cellArrayRowColIndexTest() throws Exception {
    //given
    int count = 10;
    String[] headers = new String[]{"헤더1","헤더2","헤더3","헤더4","헤더5"};
    String[] data = new String[]{"데이터1","데이터2","데이터3","데이터4","데이터5"};
    ExcelExporter exporter = SXSSFExcelExporter.builder()
        .build();
    exporter.createSheet("시트1");

    //when
    exporter.createHeader().writeCells(headers);
    CellWriter writer = null;
    for(int i = 0 ; i < count ; i++){
      writer = exporter.createRow().writeCells(data);
    }

    //then
    int resultRow = 11;
    int resultCol = 5;
    Assertions.assertEquals(resultRow, exporter.getCurrentRow());
    Assertions.assertEquals(resultCol, writer.getCurrentCol());
  }

  @Test
  @Order(3)
  @DisplayName("셀 커스텀 객체 헤더/바디 테스트")
  public void cellCustomHeaderBodyTest() throws Exception {
    //given
    int count = 10;
    Class<CustomExportableDto> clazz = CustomExportableDto.class;
    List<CustomExportableDto> objects = new ArrayList<>();
    for(int i = 0 ; i < count ; i++){
      int withdrawValue = new SecureRandom().nextInt(10);
      CustomExportableDto obj = new CustomExportableDto();
      obj.setId((long)(i+1));
      obj.setName("회원명_"+(i+1));
      obj.setEmail("test"+(i+1)+"@gmail.com");
      obj.setPhone("01000000000");
      boolean withdraw = withdrawValue == 4;
      obj.setWithdraw(withdraw);
      if(withdraw){
        LocalDateTime withdrawAt = LocalDateTime.now();
        obj.setWithdrawAt(withdrawAt);
      }
      objects.add(obj);
    }

    ExcelExporter exporter = SXSSFExcelExporter.builder()
        .build();
    exporter.createSheet("시트1");

    //when
    exporter.createHeader().writeTargetHeaders(clazz);
    CellWriter writer = null;
    for(CustomExportableDto dto : objects){
      writer = exporter.createRow().writeTargetCells(dto);
    }

    //then
    int resultRow = 11;
    int resultCol = 6;
    Assertions.assertEquals(resultRow, exporter.getCurrentRow());
    Assertions.assertEquals(resultCol, writer.getCurrentCol());
  }

  @Test
  @Order(4)
  @DisplayName("시트 커스텀 객체 Boolean(Getter) Exception 테스트")
  public void cellCustomGetterExceptionTest() throws Exception {
    //given
    CustomExportableDto2 obj = new CustomExportableDto2();
    obj.setId(1L);
    obj.setName("회원명_1");
    obj.setEmail("test1@gmail.com");
    obj.setPhone("01000000000");
    obj.setWithdraw(true);
    obj.setWithdrawAt(LocalDateTime.now());

    ExcelExporter exporter = SXSSFExcelExporter.builder()
        .build();
    exporter.createSheet("시트1");

    //when
    RuntimeException thrown = Assertions.assertThrows(RuntimeException.class, () -> {
      exporter.createRow().writeTargetCells(obj);
    });

    //then
    Class<?> expectedException = NoSuchMethodException.class;

    Assertions.assertEquals(expectedException, thrown.getCause().getClass());
  }

}
