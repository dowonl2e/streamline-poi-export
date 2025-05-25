# Streamline POI Export

이 모듈은 Apache POI를 사용하여 Excel 다운로드 구현을 보다 간소화합니다.

## 지원 버전

| Component       | Supported Version(s) |
|-----------------|----------------------|
| Java            | 11, 17, 21, 24       |
| Apache POI      | 5.2.3+               |

## 필수 적용 사항

1. Apache POI Dependencies 추가

   `[Supported Apache POI Version]`는 **지원 버전**에서 Apache POI 버전을 입력해야 합니다

   - Gradle 
        ```
        dependencies {
            implementation 'org.apache.poi:poi:[Supported Apache POI Version]'
            implementation 'org.apache.poi:poi-ooxml:[Supported Apache POI Version]'
        }
        ``` 
   - Maven
      ```
        <dependencies>
          <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi</artifactId>
            <version>[Supported Apache POI Version]</version>
          </dependency>
          <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi-ooxml</artifactId>
            <version>[Supported Apache POI Version]</version>
          </dependency>
      </dependencies>
      ```
     
## 사용 예시

### 1. 셀 스타일 적용을 위한 폰트 및 스타일 적용자 구현 예시

#### 폰트 스타일 적용 객체 생성

```java
FontStyleApplier fontStyleApplier = FontStyleApplier.builder()
    .font(c -> 
        c.bold().sizeInPoints((short) 12)  // 폰트 굵기 및 사이즈 설정
    )
    .build();
```

#### 셀 스타일 적용자 생성

```java
//스타일 적용자 생성
CellStyleApplier customCellStyleApplier = CellStyleApplier.builder()
    .wrap(c -> c.enable())  // 텍스트 Wrapping
    .alignment(c -> 
        c.horizontal(HorizontalAlignmentValues.CENTER)   // 텍스트 수평 정렬
            .vertical(VerticalAlignmentValues.CENTER)   // 텍스트 수직 정렬
    )
    .border(c ->
        c.styleAll(BorderStyleValues.THIN)  // 테두리 스타일 설정
            .colorAll(IndexedColorValues.BLACK))  // 테두리 색 설정
    .foreground(c -> c.color(IndexedColorValues.PALE_BLUE))  // 셀 배경색 설정
    .fontStyleApplier(fontStyleApplier) // 폰트 스타일 적용자 설정
    .build();

// 스타일 일부 구성 변경
customCellStyleApplier
    .alignment(c ->
        c.horizontal(HorizontalAlignmentValues.RIGHT)
            .vertical(VerticalAlignmentValues.BOTTOM)
    )
    .border(c ->
        c.styleAll(BorderStyleValues.THICK)
            .colorAll(IndexedColorValues.RED)
    );
```

### 2.엑셀 Workbook 생성 설정 예시

#### Apache POI XSSF 

```java
ExcelExporter exporter = XSSFExcelExporter.builder()
    .defaultHeaderCellStyleApplier(defaultHeaderStyleApplier) // 전체 시트 기본 헤더 스타일 지정(선택)
    .defaultBodyCellStyleApplier(defaultBodyStyleApplier) // 전체 시트 기본 헤더 스타일 지정(선택)
    .build();
```

#### Apache POI SXSSF

```java
ExcelExporter exporter = SXSSFExcelExporter.builder()
    .defaultHeaderCellStyleApplier(defaultHeaderStyleApplier) // 전체 시트 기본 헤더 스타일 지정(선택)
    .defaultBodyCellStyleApplier(defaultBodyStyleApplier) // 전체 시트 기본 헤더 스타일 지정(선택)
    .rowAccessWindowSize(100)   // Streaming 방식의 메모리에 유지할 수 있는 행의 개수 지정(선택)
    .flushCount(200) // Streaming 방식의 flush Rows 설정 (default: 100)
    .build();
```

### 3. 커스텀 셀 스타일 및 폰트 생성 예시

```java
exporter
    .createCellStyle("ISSUE_STYLE", issueStyleApplier) // 커스텀 스타일 Key, CellStyleApplier로 셀 스타일 생성
    .createFont("BLUE_FONT", blueFontStyleApplier); // 커스텀 폰트 Key, FontStyleApplier로 폰트 생성
```

### 4. 시트 생성 및 구성 예시

```java
exporter.createSheet("Sheet")
    .displayZeros(true); //값이 0인 셀에 0을 표시할지 말지 설정
    .rowBreaks(1000) //특정 행에서 강제 인쇄 페이지 나눔 설정
    .marginTop(20).marginLeft(10).marginBottom(20).marginRight(10); //인쇄 시 페이지 여백(Top, Left, Bottom, Right) 설정
```

### 5. 헤더 구성 예시

#### 단일 데이터를 이용한 셀 헤더 구성

```java
exporter.createHeader()
    .writeCell("번호")
    .writeCell("이름")
    .writeCell("이메일")
    .writeCell("탈퇴여부")
    .writeCell("탈퇴일시");
```

#### 단일 데이터를 이용한 셀 스타일 및 헤더 구성

```java
exporter.createHeader()
    .writeCell("번호")
    .writeCell("ISSUE_STYLE", "이름") // 커스텀 스타일 Key 값
    .writeCell("ISSUE_STYLE", "이메일") // 커스텀 스타일 Key 값
    .writeCell("탈퇴여부")
    .writeCell("탈퇴일시");
```

#### 배열를 이용한 셀 헤더 구성

```java
exporter.createHeader()
    .writeCells(new String[]{"번호", "이름", "이메일", "탈퇴여부", "탈퇴일시"});
```

#### 배열를 이용한 셀 스타일 및 헤더 구성

```java
exporter.createHeader()
    .writeCells("ISSUE_STYLE", new String[]{"번호", "이름", "이메일", "탈퇴여부", "탈퇴일시"}); // 커스텀 스타일 Key 값
```

#### 커스텀 객체(DTO)를 이용한 헤더 구성

```java
exporter.createHeader()
    .writeTargetHeaders({{CustomDto implements Exportable Class}});
```

#### 커스텀 객체(DTO)를 이용한 스타일 및 헤더 구성

```java
exporter.createHeader()
    .writeTargetHeaders("ISSUE_STYLE", {{CustomDto implements Exportable Class}});
```


### 6. 바디 구성 예시

#### 단일 데이터를 이용한 셀 바디 구성

```java
exporter.createRow()
    .writeCell(member.getId())
    .writeCell(member.getName())
    .writeCell(member.getEmail())
    .writeCell(member.isWithdraw())
    .writeCell(member.getWithdrawAt());
```

#### 단일 데이터를 이용한 셀 스타일 및 바디 구성

```java
exporter.createRow()
    .writeCell(member.getId())
    .writeCell("ISSUE_STYLE", member.getName()) // 커스텀 스타일 Key 값
    .writeCell("ISSUE_STYLE", member.getEmail()) // 커스텀 스타일 Key 값
    .writeCell(member.isWithdraw())
    .writeCell(member.getWithdrawAt());
```

#### 배열을 이용한 셀 바디 구성

```java
exporter.createRow()
    .writeCells(new String[]{ ... });
```

#### 배열을 이용한 셀 스타일 및 바디 구성

```java
exporter.createRow()
    .writeCells("ISSUE_STYLE", new String[]{ ... }); // 커스텀 스타일 Key 값
```

#### 커스텀 객체(DTO)를 이용한 바디 구성

```java
exporter.createRow()
    .writeTargetCells({{CustomDto implements Exportable}});
```

#### 커스텀 객체(DTO)를 이용한 스타일 및 바디 구성

```java
exporter.createRow()
    .writeTargetCells("ISSUE_STYLE", {{CustomDto implements Exportable}});
```