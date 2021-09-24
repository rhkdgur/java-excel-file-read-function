# java-excel-file-read-function-part1

1.우선 엑셀 파일 처리 기능 part1<br>
->저용량 데이터 처리 시 사용하는 코드 입니다.

우선 엑셀 파일을 읽기 위해서는 라이브러리가 필요
-> poi.jar 파일이 필요
jsp/servlet 같은 경우 library에 선언해주면되고 Spring 같은 경우에는 pom.xml에 선언 해주면 된다.

```C
FileInputStream file = new FileInputStream("d:\\excelread.xlsx");
XSSFWorkbook workbook = new XSSFWorkbook(file);
XSSFSheet sheet = null;


for(int i=0 ; i<workbook.length ; i++){ //시트 개수
        sheet = workbook.getSheetAt(i);
        int rows = sheet.getPhysicalNumberOfRows();


        for(int j=0 ; j<rows ;j++){ //행
                XSSFRow row = sheet.getRow(j);
                int cells = row.getPhysicalNumberOfCells();


                for(int k=0 ; k<cells ;j++){ //열
                        XSSFCell cell = row.getCell(k);


                        switch (cell.getCellType()){
                        case XSSFCell.CELL_TYPE_FORMULA:
                                value=cell.getCellFormula();
                                break;
                        case XSSFCell.CELL_TYPE_NUMERIC:
                                value=cell.getNumericCellValue()+"";
                                break;
                        case XSSFCell.CELL_TYPE_STRING:
                                value=cell.getStringCellValue()+"";
                                break;
                        case XSSFCell.CELL_TYPE_BLANK:
                                value=cell.getBooleanCellValue()+"";
                                break;
                        case XSSFCell.CELL_TYPE_ERROR:
                                value=cell.getErrorCellValue()+"";
                                break;
                        }
                        System.out.println(i + "번 시트 : " + j + "행의 " + k + "열 = " + value);
                }
        }
}
```

|이름|설명|
|:---|---:|
|XSSF|엑셀 2007부터 지원하는 OOXML 파일 모맷인 '.xlsx' 파일을 읽고 쓰는 컴포넌트입니다.|
|XSSFWorkbook|하나의 파일에 대한 모든 정보를 저장하고 있습니다.|
|XSSFSheet|하나의 파일의 시트에 대한 정보를 가지고 있습니다.|
|XSSFRow|하나의 파일 시트에 해당하는 행에 대한 정보를 가지고 있습니다.|
|XSSFCell|한 행의 셀(열)에 해당하는 정보를 가지고 있습니다.|
