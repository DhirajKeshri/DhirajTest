
public class One {

	
	String FilePath = Helper.Filepath("Container") ;                  //Excel file call in constructor
	FileInputStream	fs = new FileInputStream(FilePath);  // file Path 
	XSSFWorkbook work = new XSSFWorkbook(fs);
	XSSFSheet containerSheet = work.getSheet("Container");
	
	String ContainerNumber = null;
	for(int i = 0;i<containerSheet.getLastRowNum();i++)
	{
		try {
			if((containerSheet.getRow(i).getCell(1).getStringCellValue()).equalsIgnoreCase(""))
			{
				ContainerNumber =  containerSheet.getRow(i).getCell(0).getStringCellValue().trim();
				//Row row = S.createRow(i);
				
			}
					
		}
		catch (Exception e) 
		{
			Row row = containerSheet.getRow(i);
			row.createCell(1).setCellValue("Used");
			FileOutputStream fos = new FileOutputStream(FilePath);
            work.write(fos);
            fos.close();
           break;
		}
	}	
	work.close();
	return ContainerNumber;

}

}
