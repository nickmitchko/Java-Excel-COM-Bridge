/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package main.bridge;

import com.jacob.com.ComThread;
import com.jacob.com.LibraryLoader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * @file    Loader.java
 * @date    Jun 15, 2014
 * @author  Nicholai
 */
public class Loader {
    
    public Loader() throws Exception{
        try {
            loadLibrary();
            startThread();
        } catch (IOException ex) {
            Logger.getLogger(Loader.class.getName()).log(Level.SEVERE, null, ex);
            endThread();
        }
    }
    
    public void shutdown() throws Exception{
        endThread();
    }

    private void endThread() throws Exception {
        ComThread.quitMainSTA();
        ComThread.Release();
    }

    private void startThread() {
        ComThread.InitSTA(true);
    }

    private void loadLibrary() throws IOException {
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
        temporaryDll.deleteOnExit();
    }

}
