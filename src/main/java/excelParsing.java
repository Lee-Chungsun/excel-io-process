//엑셀 파일 물론 절대 경로 까지 (local에 있다면)
FileInputStream fis = new FileInputStream("excel.xls");
//엑셀파일 읽이 위한 시작
HSSFWorkbook workbook = new HSSFWorkbook(fis);
//0번쨰 시트 
//여러개의 시트가 있다면 workbook.getNumberOfSheets() 로 갯수 구해서 for문
HSSFSheet sheet = workbook.getSheetAt(0);
//로우의 갯수
int rows = sheet.getPhysicalNumberOfRows();
//내가 필요해서 넣은 변수
int valueee = 1;
for (int rowIndex = 1; rowIndex < rows; rowIndex++) {
  //row 불러오기
  HSSFRow row = sheet.getRow(rowIndex);
  String value = "";
  if (row != null) {
    int cells = row.getPhysicalNumberOfCells();
    for (int columnIndex = 0; columnIndex < cells; columnIndex++) {
      //셀 정보 가져오기
      HSSFCell cell = row.getCell(columnIndex); 
      //해당 셀 데이터의 타입을 체크 및 해당 타입에 맞게 가져오기
      switch (cell.getCellType()) { 
        case HSSFCell.CELL_TYPE_NUMERIC:
             value += cell.getNumericCellValue() + "";
             break;
        case HSSFCell.CELL_TYPE_STRING:
             value += cell.getStringCellValue() + "";
             break;
        case HSSFCell.CELL_TYPE_BLANK:
             value += cell.getBooleanCellValue() + "";
             break;
        case HSSFCell.CELL_TYPE_ERROR:
             value += cell.getErrorCellValue() + "";
             break;
      }
    }
  }  
} 
