package dto;

import com.dwl.excel.annotation.ExportTargetField;
import com.dwl.excel.exporter.Exportable;

import java.time.LocalDateTime;

public class TravelDto implements Exportable {

  @ExportTargetField(header = "여행명", order = 1)
  private String travelName;
  @ExportTargetField(header = "여행 시작일", order = 2)
  private LocalDateTime startDate;
  @ExportTargetField(header = "여행 종료일", order = 3)
  private LocalDateTime endDate;
  @ExportTargetField(header = "등록일", order = 4)
  private LocalDateTime createDate;


  public String getTravelName() {
    return travelName;
  }

  public void setTravelName(String travelName) {
    this.travelName = travelName;
  }

  public LocalDateTime getStartDate() {
    return startDate;
  }

  public void setStartDate(LocalDateTime startDate) {
    this.startDate = startDate;
  }

  public LocalDateTime getEndDate() {
    return endDate;
  }

  public void setEndDate(LocalDateTime endDate) {
    this.endDate = endDate;
  }

  public LocalDateTime getCreateDate() {
    return createDate;
  }

  public void setCreateDate(LocalDateTime createDate) {
    this.createDate = createDate;
  }

}
