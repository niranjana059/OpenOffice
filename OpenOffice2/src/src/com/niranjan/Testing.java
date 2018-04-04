package src.com.niranjan;

import java.io.File;
import java.net.ConnectException;

import org.jodconverter.OfficeDocumentConverter;
import org.jodconverter.office.DefaultOfficeManagerBuilder;
import org.jodconverter.office.OfficeException;
import org.jodconverter.office.OfficeManager;

public class Testing { public static void main(String[] args) {
	 
    // 1) Start LibreOffice in headless mode.
    OfficeManager officeManager = null;
    try {
        officeManager = new DefaultOfficeManagerBuilder()
                .setOfficeHome(new File("C:/Program Files (x86)/OpenOffice 4")).build();
        try {
			officeManager.start();
		} catch (OfficeException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

        // 2) Create JODConverter converter
        OfficeDocumentConverter converter = new OfficeDocumentConverter(
                officeManager);

        // 3) Create PDF
        createPDF(converter);

    } finally {
        // 4) Stop LibreOffice in headless mode.
        if (officeManager != null) {
            try {
				officeManager.stop();
			} catch (OfficeException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
        }
    }
}

private static void createPDF(OfficeDocumentConverter converter) {
    try {
        long start = System.currentTimeMillis();
        converter.convert(new File("input/test.xlsx"), new File(
                "output/test.xlsx.pdf"));
        System.out.println("output/Performance_Out.pdf with "
                + (System.currentTimeMillis() - start) + "ms");
    } catch (Throwable e) {
        e.printStackTrace();
    }
}}
