package excel;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;

public class Excel_read implements direction {
	
	public static void main (String [] args) throws EncryptedDocumentException, IOException {
		Excel_read r = new Excel_read();
		r.input();
		
		r.nextfoermate();
		r.squareformat();
		r.reverceSquare();
		r.output();
	
		}

}
