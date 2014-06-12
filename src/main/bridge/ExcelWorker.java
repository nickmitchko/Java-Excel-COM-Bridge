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
    
    public ArrayList<String> getSheets(){
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
    
    public Object getCellValue(int row, int col){
        if(!this.isLoaded){
            return null;
        }
        Dispatch cell = Dispatch.invoke(activeSheet, "Cells", Dispatch.Get, new Object[] {row+1, col+1}, new int[1]).toDispatch();
        return Dispatch.get(cell, "Value");
    }
    
    public void recalculate(){
        if(!this.isLoaded){
            return;
        }
        this.ComBridge.invoke("CalculateFull");
    }
    
    public ArrayList<String> load(File fileIn){
        return this.load(fileIn, false);
    }
    
    public ArrayList<String> load(File fileIn, boolean visible){
        Dispatch.put(ComBridge, "Visible", new Variant(visible));
        Dispatch workbook = ComBridge.getProperty("Workbooks").toDispatch();
        Dispatch.call(workbook, "Open", new Variant(fileIn.getAbsolutePath()));
        this.activeWorkbook = ComBridge.getProperty("ActiveWorkbook").toDispatch();
        isLoaded = true;
        return this.getSheets();
    }
    
    public void setCellFormula(String formula, int row, int col){
        Dispatch cellRef = Dispatch.invoke(activeSheet, "Cells", Dispatch.Get, new Object[] {row+1, col+1}, new int[1]).toDispatch();
        Dispatch.put(cellRef, "Formula", formula);
    }
    
    public void saveAs(File s){
        Dispatch.call(this.activeWorkbook, "SaveAs", s.getAbsolutePath());
    }
    
    public void save(){
        Dispatch.call(this.activeWorkbook, "Save");
    }
    
    public void enableCalculation(){
        Dispatch.put(this.ComBridge, "Calculation", -4105);
    }
    
    public void disableCalculation(){
        Dispatch.put(this.ComBridge, "Calculation", -4135);
    }
    
    public void close(){
        Dispatch.call(this.activeWorkbook, "Close");
    }
    
    public void quit(){
        this.ComBridge.invoke("Quit");
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
        this.ComBridge.safeRelease();
    }
    
    public void endProcess(){
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
