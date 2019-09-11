//Excel Write 엑셀 쓰기 시작
HSSFWorkbook ediExcel = new HSSFWorkbook();
//list는 조회 결과가 들어있음 List<Map>
for (Map<String, Object> item : list) {
  //sheet name 지정
  String sheetName = "Sheet";

  //sheet 만들어진 sheet인지 아닌지 체크 없으면 -1 반환 
  int value = ediExcel.getSheetIndex(sheetName);
  int rowNum = 0;

  //sheet 생성 
  HSSFSheet sheet = null;
  if(value == -1)
    sheet = ediExcel.createSheet(sheetName);
  if(value != -1){
    //이전에 생성된 같은이름의 sheet라면 정보를 불러옴
	sheet = ediExcel.getSheetAt(value);
    //마지막 row 인덱스를 가져옴
	rowNum = sheet.getLastRowNum();
  }

  //1ROW COLUMN NAME
  Set<String> colId = item.keySet();

  //FONT, ROW HEIGHT
  HSSFFont font = ediExcel.createFont(); 
  font.setFontName("맑은 고딕");
  font.setFontHeight((short)222);

  //CELL STYLE 
  HSSFCellStyle titlestyle = ediExcel.createCellStyle(); 
  titlestyle.setAlignment(HSSFCellStyle.ALIGN_LEFT);
  titlestyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
  titlestyle.setFont(font);

  HSSFRow row = null;
  HSSFCell cell = null;
  int colcount = 0;

  //1ROW 생성 ( COLUMN ROW )
  if(rowNum == 0){
	//Row 생성 
	row = sheet.createRow((short)rowNum);
				
	//행 높이 지정
	row.setHeight((short)330);
				
	//COLUMN NAME P_ & GRP TAG CONTINUE
	for(String colum : colId){
	  cell = row.createCell((short)colcount);
      cell.setCellValue(colum);
	  cell.setCellStyle(titlestyle); 
	  colcount++;
	}
  }
		
 cell = null;
			
 //Data Cell 생성
 colcount = 0;
 row = sheet.createRow((short)(rowNum+1));
 for(String colum : colId){
	cell = row.createCell((short)colcount);
	if(item.get(colum) != null)
		cell.setCellValue(item.get(colum).toString());
	cell.setCellStyle(titlestyle);
	colcount++;
				
  }
			
  //cell 너비 자동 맞춤 (data 길이에 따라 적용)
  for(int i=0;i<colId.size();i++){
	 sheet.autoSizeColumn(i);
	 sheet.setColumnWidth(i, (sheet.getColumnWidth(i))+512 ); 
  }
			
  //행 높이 지정
  row.setHeight((short)330);
