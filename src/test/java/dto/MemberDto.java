package dto;

import com.dwl.excel.annotation.ExportTargetField;
import com.dwl.excel.exporter.Exportable;

import java.util.List;

public class MemberDto implements Exportable {

  @ExportTargetField(header = "번호", order = 1)
  private Long id;
  @ExportTargetField(header = "이름", order = 2)
  private String name;
  @ExportTargetField(header = "이메일", order = 3)
  private String email;

  private List<TravelDto> travels;

  public Long getId() {
    return id;
  }

  public void setId(Long id) {
    this.id = id;
  }

  public String getName() {
    return name;
  }

  public void setName(String name) {
    this.name = name;
  }

  public String getEmail() {
    return email;
  }

  public void setEmail(String email) {
    this.email = email;
  }

  public List<TravelDto> getTravels() {
    return travels;
  }

  public void setTravels(List<TravelDto> travels) {
    this.travels = travels;
  }
}
