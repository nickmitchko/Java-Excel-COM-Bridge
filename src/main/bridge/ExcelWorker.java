/**
 * How to use This CLASS:
 * 1. Initialize    : ExcelWorker worker = new ExcelWorker();
 * 2. Load Workbook : worker.load(new File("C:\\Workbook"));
 * 3. Set Sheet     : worker.setSheet(3); //Third sheet (excel is base 1 unlike java)
 * 4. Use the sheet : worker.setCellValue("This is A3", "A3");
 *                  : Object cellVal = worker.getCellValue("D5");
 */

package main.bridge;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.LibraryLoader;
import com.jacob.com.Variant;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * @date June 5, 2014
 * @name ExcelWorker.java
 * @author nicholai.mitchkoax
 */
public class ExcelWorker {
    
    private ActiveXComponent ComBridge;
    private Dispatch activeWorkbook;
    private Dispatch sheetsDispatch;
    private Dispatch activeSheet;
    private boolean isLoaded = false;
    
    public ExcelWorker() throws IOException{
        initComBridge();
    }
    
    public List getSheets(){
        if(!this.isLoaded){
            return null;
        }
        @SuppressWarnings("UnusedAssignment")
        ArrayList<String> sheets = new ArrayList<>();
        this.sheetsDispatch = Dispatch.get(activeWorkbook, "Sheets").toDispatch();
        int sheetCount = Dispatch.get(sheetsDispatch, "Count").getInt();
        sheets = new ArrayList<>();
        for (int i = 1; i <= sheetCount; i++) {
            Dispatch onesheet = Dispatch.invoke(sheetsDispatch, "Item", Dispatch.Get, new Object[]{i}, new int[0]).getDispatch();
            sheets.add(Dispatch.get(onesheet, "Name").toString());
        }
        return sheets;
    }
    
    public void setSheet(int i){
        if(!this.isLoaded){
            return;
        }
        activeSheet = Dispatch.invoke(sheetsDispatch, "Item", Dispatch.Get, new Object[]{i}, new int[0]).getDispatch();
    }
    
    public void setSheet(String str){
        if(!this.isLoaded){
            return;
        }
        activeSheet = Dispatch.invoke(sheetsDispatch, "Item", Dispatch.Get, new Object[]{str}, new int[0]).getDispatch();
    }
    
    public void setCellValue(Object val, String cellRef){
        if(!this.isLoaded){
            return;
        }
        Dispatch cell = Dispatch.invoke(activeSheet, "Range", Dispatch.Get, new Object[]{cellRef}, new int[1]).toDispatch();
        Dispatch.put(cell, "Value", val);
    }
    
    public void setCellValue(Object val, int row, int col){
        if(!this.isLoaded){
            return;
        }
        Dispatch cellRef = Dispatch.invoke(activeSheet, "Cells", Dispatch.Get, new Object[]{++row, ++col}, new int[1]).toDispatch();
        //Dispatch cell = Dispatch.invoke(activeSheet, "Range", Dispatch.Get, new Object[]{cellRef}, new int[1]).toDispatch();
        Dispatch.put(cellRef, "Value", val);
    }
    
    public void setCellValue(Date val, int row, int col){
        if(!this.isLoaded){
            return;
        }
        Dispatch cellRef = Dispatch.invoke(activeSheet, "Cells", Dispatch.Get, new Object[]{++row, ++col}, new int[1]).toDispatch();
        //Dispatch cell = Dispatch.invoke(activeSheet, "Range", Dispatch.Get, new Object[]{cellRef}, new int[1]).toDispatch();
        Dispatch.put(cellRef, "Value", val);
    }
    
    public void setCellValue(String val, int row, int col){
        if(!this.isLoaded){
            return;
        }
        Dispatch cellRef = Dispatch.invoke(activeSheet, "Cells", Dispatch.Get, new Object[]{++row, ++col}, new int[1]).toDispatch();
        //Dispatch cell = Dispatch.invoke(activeSheet, "Range", Dispatch.Get, new Object[]{cellRef}, new int[1]).toDispatch();
        Dispatch.put(cellRef, "Value", val);
    }
    
    public void setCellValue(int val, int row, int col){
        if(!this.isLoaded){
            return;
        }
        Dispatch cellRef = Dispatch.invoke(activeSheet, "Cells", Dispatch.Get, new Object[]{++row, ++col}, new int[1]).toDispatch();
        //Dispatch cell = Dispatch.invoke(activeSheet, "Range", Dispatch.Get, new Object[]{cellRef}, new int[1]).toDispatch();
        Dispatch.put(cellRef, "Value", val);
    }
    
    public void setCellValue(double val, int row, int col){
        if(!this.isLoaded){
            return;
        }
        Dispatch cellRef = Dispatch.invoke(activeSheet, "Cells", Dispatch.Get, new Object[]{++row, ++col}, new int[1]).toDispatch();
        //Dispatch cell = Dispatch.invoke(activeSheet, "Range", Dispatch.Get, new Object[]{cellRef}, new int[1]).toDispatch();
        Dispatch.put(cellRef, "Value", val);
    }
    
    public Object getCellValue(String cellRef){
        if(!this.isLoaded){
            return null;
        }
        Dispatch cell = Dispatch.invoke(activeSheet, "Range", Dispatch.Get, new Object[] {cellRef}, new int[1]).toDispatch();
        return Dispatch.get(cell, "Value");
    }
    
    public void recalculate(){
        if(!this.isLoaded){
            return;
        }
        this.ComBridge.invoke("CalculateFull");
    }
    
    public void load(File fileIn){
        Dispatch.put(ComBridge, "Visible", new Variant(false));
        Dispatch workbook = ComBridge.getProperty("Workbooks").toDispatch();
        Dispatch.call(workbook, "Open", new Variant(fileIn.getAbsolutePath()));
        this.activeWorkbook = ComBridge.getProperty("ActiveWorkbook").toDispatch();
        isLoaded = true;
        this.getSheets();
    }
    
    public void load(File fileIn, boolean visible){
        Dispatch.put(ComBridge, "Visible", new Variant(visible));
        Dispatch workbook = ComBridge.getProperty("Workbooks").toDispatch();
        Dispatch.call(workbook, "Open", new Variant(fileIn.getAbsolutePath()));
        this.activeWorkbook = ComBridge.getProperty("ActiveWorkbook").toDispatch();
        isLoaded = true;
        this.getSheets();
    }
    
    public void saveAs(File s){
        Dispatch.call(this.activeWorkbook, "SaveAs", s.getAbsolutePath());
    }
    
    public void save(){
        Dispatch.call(this.activeWorkbook, "Save");
    }
    
    private void initComBridge() throws IOException{
        //lib file name
        String libFile = System.getProperty("os.arch").equals("amd64") ? "jacob-1.18-M2-x64.dll" : "jacob-1.18-M2-x86.dll";
        /* Read DLL file*/
        InputStream inputStream = ExcelWorker.class.getResourceAsStream(libFile);
        File temporaryDll = File.createTempFile("jacob", ".dll");
        /* Write dll to a tempFile */
        try (FileOutputStream outputStream = new FileOutputStream(temporaryDll)) {
            byte[] array = new byte[8192];
            for (int i = inputStream.read(array); i != -1; i = inputStream.read(array)) {
                outputStream.write(array, 0, i);
            }
        }
        System.setProperty(LibraryLoader.JACOB_DLL_PATH, temporaryDll.getAbsolutePath());
        LibraryLoader.loadJacobLibrary();
        this.ComBridge = new ActiveXComponent("Excel.Application");
        temporaryDll.deleteOnExit();
    }
    
    public void safeRelease(){
        Dispatch.call(this.activeWorkbook, "Close");
        this.ComBridge.invoke("Quit");
        this.ComBridge.safeRelease();
        ComThread.Release();
        ComThread.quitMainSTA();
    }
    
    public static void main(String args[]){
        try {
            ExcelWorker ew = new ExcelWorker();
            System.out.println();
        } catch (IOException ex) {
            Logger.getLogger(ExcelWorker.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
}
