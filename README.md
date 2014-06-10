Java-Excel-COM-Bridge (1.0)
=====================


A library that properly communicates with an excel com object to load, open, modify, get values from, and save any Excel acceptable file. Uses [jacob], which is already implemented.


 How to use This Bridge:

 - Setting Up, Library is in ~/distribution

               1. Use Java-Excel-COM-Bridge.jar as library either in project or classpath. 
               2. import main.bridge.ExcelWorker;

 - Initialize

                    ExcelWorker worker = new ExcelWorker();
                    
 - Load Workbook (Use any excel supported file here)

                    worker.load(new File("C:\\Workbook.xlsx"));
                    //Load with excel window visible 
                    worker.load(new File("C:\\Workbook.xlsx"), true);
                    
 - Set Sheet     

                    // You must select a work sheet to operate on
                    worker.setSheet(3); //Third sheet
                    worker.setSheet("Sheet Sample"); //Sheet Titled "Sheet Sample"
 - Use the sheet

                    worker.setCellValue("This is A3", "A3");
                    Object cellVal = worker.getCellValue("D5");
                    
 - Save the Sheet

                    worker.saveAs(new File("C:\\newWorkbook.xlsx");
                    
 - Destroy the memory

                    worker.safeRelease();
                    //Jacob has to be destoryed manually or else there are lingering
                    //Excel processes which leak memory

![alt text](http://i.imgur.com/b3FuvgU.png, "Data Diagram")

- Version 1.0
![GitHub Logo](http://www.brotherlykicks.com/pics/profile.png)


- Version 2.0

[jacob]:http://danadler.com/jacob/
