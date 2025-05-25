package sheet;

import com.dwl.excel.exporter.ExcelExporter;
import com.dwl.excel.exporter.SXSSFExcelExporter;
import com.dwl.excel.exporter.writer.CellWriter;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.MethodOrderer;
import org.junit.jupiter.api.Order;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestMethodOrder;

@TestMethodOrder(MethodOrderer.OrderAnnotation.class)
public class SheetMergeTest {

  @Test
  @Order(1)
  @DisplayName("헤더 병합 후 시작열 확인 테스트")
  public void headerStartColumnByMergeRegion() throws Exception {
    //given
    String[] firstHeader1 = new String[]{"헤더0", "헤더1", "헤더2", "헤더(3~5)", "헤더6", "헤더7"};
    String[] firstHeader2 = new String[]{"헤더0", "헤더1", "헤더2", "헤더(3~5)", "헤더6", "헤더7"};
    ExcelExporter exporter1 = SXSSFExcelExporter.builder()
        .build();
    ExcelExporter exporter2 = SXSSFExcelExporter.builder()
        .build();

    //when
    exporter1.createSheet();

    exporter1.mergeCell(0, 1, 1, 1)
        .mergeCell(0, 1, 2, 2)
        .mergeCell(0, 0, 3, 5)
        .mergeCell(0, 1, 6, 6)
        .mergeCell(0, 1, 7, 7);

    exporter1.createHeader().writeCells(firstHeader1);
    int currentCol1 = exporter1.createHeader().getCurrentCol();

    exporter2.createSheet();
    exporter2.mergeCell(0, 1, 0, 0)
        .mergeCell(0, 1, 1, 1)
        .mergeCell(0, 1, 2, 2)
        .mergeCell(0, 0, 3, 5)
        .mergeCell(0, 1, 6, 6)
        .mergeCell(0, 1, 7, 7);

    exporter2.createHeader().writeCells(firstHeader2);
    int currentCol2 = exporter2.createHeader().getCurrentCol();

    //then
    int expectedColumn1 = 0;
    int expectedColumn2 = 3;

    Assertions.assertEquals(expectedColumn1, currentCol1);
    Assertions.assertEquals(expectedColumn2, currentCol2);
  }

  @Test
  @Order(2)
  @DisplayName("바디 병합 후 시작열 확인 테스트")
  public void bodyStartColumnByMergeRegion() throws Exception {
    //given
    String[] row1 = new String[]{"행1-열1","행1-열2","행1-열3","행1-열4","행1-열5"};
    String[] row2 = new String[]{"행2-열2","행1-열3","행1-열4","행1-열5"};
    String[] row3 = new String[]{"행3-열3","행1-열4","행1-열5"};
    String[] row4 = new String[]{"행4-열4","행1-열5"};
    String[] row5 = new String[]{"행5-열5"};
    ExcelExporter exporter = SXSSFExcelExporter.builder()
        .build();

    //when
    exporter.createSheet();
    exporter.mergeCell(0, 4, 0, 0)
        .mergeCell(1, 4, 1, 1)
        .mergeCell(2, 4, 2, 2)
        .mergeCell(3, 4, 3, 3);

    CellWriter writer = exporter.createRow();
    int currentCol1 = writer.getCurrentCol();
    writer.writeCells(row1);

    writer = exporter.createRow();
    int currentCol2 = writer.getCurrentCol();
    writer.writeCells(row2);

    writer = exporter.createRow();
    int currentCol3 = writer.getCurrentCol();
    writer.writeCells(row3);

    writer = exporter.createRow();
    int currentCol4 = writer.getCurrentCol();
    writer.writeCells(row4);

    writer = exporter.createRow();
    int currentCol5 = writer.getCurrentCol();
    writer.writeCells(row5);

    //then
    int expectedColumn1 = 0;
    int expectedColumn2 = 1;
    int expectedColumn3 = 2;
    int expectedColumn4 = 3;
    int expectedColumn5 = 4;

    Assertions.assertEquals(expectedColumn1, currentCol1);
    Assertions.assertEquals(expectedColumn2, currentCol2);
    Assertions.assertEquals(expectedColumn3, currentCol3);
    Assertions.assertEquals(expectedColumn4, currentCol4);
    Assertions.assertEquals(expectedColumn5, currentCol5);
  }
}
