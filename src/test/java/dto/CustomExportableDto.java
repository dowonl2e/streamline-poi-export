package dto;

import com.dwl.excel.annotation.ExportTargetField;
import com.dwl.excel.exporter.Exportable;

import java.time.LocalDateTime;

public class CustomExportableDto implements Exportable {

  @ExportTargetField(header = "번호", order = 1)
  private Long id;
  @ExportTargetField(header = "이름", order = 2)
  private String name;
  @ExportTargetField(header = "이메일", order = 3)
  private String email;
  @ExportTargetField(header = "휴대폰번호", order = 4)
  private String phone;
  @ExportTargetField(header = "탈퇴여부", order = 5)
  private Boolean withdraw;
  @ExportTargetField(header = "탈퇴일시", order = 6)
  private LocalDateTime withdrawAt;

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

  public String getPhone() {
    return phone;
  }

  public void setPhone(String phone) {
    this.phone = phone;
  }

  public Boolean isWithdraw() {
    return withdraw;
  }

  public void setWithdraw(Boolean withdraw) {
    this.withdraw = withdraw;
  }

  public LocalDateTime getWithdrawAt() {
    return withdrawAt;
  }

  public void setWithdrawAt(LocalDateTime withdrawAt) {
    this.withdrawAt = withdrawAt;
  }
}
