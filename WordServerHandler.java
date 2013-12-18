import java.io.File;
import java.io.Reader;
import java.io.BufferedReader;
import java.io.OutputStreamWriter;
import java.io.BufferedInputStream;
import java.io.FileReader;
import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.InputStream;
import java.io.IOException;
import java.util.concurrent.Executors;
import java.util.concurrent.ExecutorService;
import org.apache.mina.common.IdleStatus;
import org.apache.mina.handler.StreamIoHandler;
import org.apache.mina.common.IoSession;
import org.apache.mina.util.SessionLog;

import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.*;
import org.apache.poi.xwpf.extractor.*;
import org.apache.xmlbeans.*;
import org.apache.poi.openxml4j.exceptions.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.extractor.XSSFExcelExtractor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import org.apache.pdfbox.pdmodel.*;
import org.apache.pdfbox.util.*;

import de.innosystec.unrar.Archive;
import de.innosystec.unrar.rarfile.FileHeader;

public class WordServerHandler extends StreamIoHandler 
{
    private ExecutorService pool = Executors.newCachedThreadPool();
     
    private static class Worker extends Thread 
    {
        private static final byte OP_WORD = 1;
        private static final byte OP_EXCEL = 2;
	private static final byte OP_WORDX = 3;
	private static final byte OP_EXCELX = 4;
	private static final byte OP_PDF = 5;
	private static final byte OP_RAR = 6;
	private static final byte OP_ZIP = 7;

        private static final byte OP_OK = 0;
        private static final byte OP_FAIL = -1;

	private static final String CHARSET = "utf-8";

        private final InputStream in;
        private final OutputStream out;
        private final IoSession session;

        public Worker(IoSession session,InputStream in,OutputStream out) 
        {
            setDaemon(true);
            this.in = in;
            this.out = out;
            this.session = session;
        }

	public static String getCharset(String fileName)
	{
	    String charset = ""; 
	    try 
	    {
	        BufferedInputStream bin = new BufferedInputStream(new FileInputStream(
	            new File(fileName)));
	        int p = (bin.read() << 8) + bin.read();
		bin.close();
	        switch (p) 
	        {
	            case 0xefbb: charset = "UTF-8"; break;
	            case 0xfffe: charset = "Unicode"; break;
	            case 0xfeff: charset = "UTF-16BE"; break;
	            default: charset = "gbk"; break;
	        }
	    }
	    catch(IOException e) {}
	    return charset;
	}

	public static void deleteFile(String sPath)
	{
		File file = new File(sPath);
		if (file.isFile() && file.exists())
		{
			file.delete();	
		}
	}

        public void run() 
        {
            int repWordLen = 0;
            byte repWordOpCode = OP_FAIL;

	    int reqWordLen = 0;
	    byte reqWordOpCode = 0;

            try
	    {
                byte[] hl = new byte[4];
                byte[] ho = new byte[1];
                System.out.println("session extrator started.");
                if (in.read(hl) != -1) 
                {
                    if (in.read(ho) != -1) 
                    {
                        reqWordLen = Convertor.bytesToInt(hl);
                        reqWordOpCode = ho[0];
                        System.out.println(reqWordLen);
			String wfname = String.valueOf(Thread
				.currentThread().getId());

			if (reqWordLen > 0)
			{
			    FileOutputStream fos = new FileOutputStream(
			    	new File(wfname));
			    byte[] word_bytes = new byte[4096];
			    int offset = 0;
			    int left = reqWordLen;
			    while (left > 0) 
			    {
			        if (left < 4096)
			        {
			    	byte[] remainings = new byte[left];
			    	if ((offset = in.read(remainings)) != -1)
			    	{
			    	    left -= offset;
			    	    fos.write(remainings, 0, offset);
			    	} 
			    	else
			    	{
			    	    throw new IOException();
			    	}
			        } 
			        else
			        {
			    	if ((offset = in.read(word_bytes)) != -1) 
			    	{
			    	    left -= offset;
			    	    fos.write(word_bytes, 0, offset);
			    	}
			    	else 
			    	{
			    	    throw new IOException();
			    	}
			        }
			    }

			    fos.flush();
			    fos.close();
			}

                        if (reqWordLen > 0 && reqWordOpCode == OP_WORD) 
                        {
                            System.out.println("WordExtractor start.");

                            FileInputStream fin = new FileInputStream(wfname);
                            WordExtractor extractor = new WordExtractor(fin);
                            String repWordData = extractor.getText();
                            fin.close();

                            SessionLog.warn(session, "extrator success");
                            repWordLen = repWordData.getBytes(CHARSET).length;
                            SessionLog.warn(session, repWordLen + "");
                            out.write(Convertor.intToByte(repWordLen));

                            repWordOpCode = OP_OK;
                            ho[0] = repWordOpCode;
                            out.write(ho);

                            out.write(repWordData.getBytes(CHARSET));
                            System.out.println("WordExtractor end.");
                        } 
                        else if (reqWordLen > 0 && reqWordOpCode == OP_WORDX) 
                        {
                            System.out.println("WordxExtractor start.");

			    POITextExtractor extractor = null;
			    OPCPackage opcPackage = POIXMLDocument.openPackage(wfname);
   		    	    extractor = new XWPFWordExtractor(opcPackage);
                            String repWordData = extractor.getText();
			    opcPackage.close();

                            SessionLog.warn(session, "extrator success");
                            repWordLen = repWordData.getBytes(CHARSET).length;
                            SessionLog.warn(session, repWordLen + "");
                            out.write(Convertor.intToByte(repWordLen));

                            repWordOpCode = OP_OK;
                            ho[0] = repWordOpCode;
                            out.write(ho);

                            out.write(repWordData.getBytes(CHARSET));
                            System.out.println("WordxExtractor end.");
                        }
                        else if (reqWordLen > 0 && reqWordOpCode == OP_EXCEL) 
                        {
                            System.out.println("ExcelExtractor start.");

                            InputStream inp = new FileInputStream(wfname);
                            HSSFWorkbook wb = new HSSFWorkbook(
                                    new POIFSFileSystem(inp));
                            ExcelExtractor extractor = new ExcelExtractor(wb);

                            extractor.setFormulasNotResults(true);
                            extractor.setIncludeSheetNames(false);
                            String repExcelData = extractor.getText();
                            inp.close();

                            SessionLog.warn(session, "extrator success");
                            repWordLen = repExcelData.getBytes(CHARSET).length;
                            SessionLog.warn(session, repWordLen + "");
                            out.write(Convertor.intToByte(repWordLen));

                            repWordOpCode = OP_OK;
                            ho[0] = repWordOpCode;
                            out.write(ho);

                            out.write(repExcelData.getBytes(CHARSET));
                            System.out.println("ExcelExtractor end.");

                        }
                        else if (reqWordLen > 0 && reqWordOpCode == OP_EXCELX) 
                        {
                            System.out.println("ExcelxExtractor start.");

                            InputStream inp = new FileInputStream(wfname);
			    Workbook wb = new XSSFWorkbook(inp);
			    XSSFExcelExtractor extractor = new XSSFExcelExtractor((XSSFWorkbook)wb);
			    extractor.setFormulasNotResults(false);
			    extractor.setIncludeSheetNames(false);
			    extractor.setIncludeCellComments(false);
                            String repExcelData = extractor.getText();
                            inp.close();

                            SessionLog.warn(session, "extrator success");
                            repWordLen = repExcelData.getBytes(CHARSET).length;
                            SessionLog.warn(session, repWordLen + "");
                            out.write(Convertor.intToByte(repWordLen));

                            repWordOpCode = OP_OK;
                            ho[0] = repWordOpCode;
                            out.write(ho);

                            out.write(repExcelData.getBytes(CHARSET));
                            System.out.println("ExcelxExtractor end.");

                        }
                        else if (reqWordLen > 0 && reqWordOpCode == OP_PDF) 
                        {
                            System.out.println("PdfxExtractor start.");

                            InputStream inp = new FileInputStream(wfname);
			    PDDocument document = PDDocument.load(inp);
			    PDFTextStripper stripper = new PDFTextStripper("utf-8");
			    stripper.setSortByPosition(false);
			    stripper.setStartPage(1);
			    stripper.setEndPage(2);
			    OutputStreamWriter output = 
				new OutputStreamWriter(new FileOutputStream(wfname + "-txt"),"utf-8"); 
			    stripper.writeText(document, output);
			    output.close();
			    String repPdfData = stripper.getText(document);
			    document.close();
			    inp.close();
				
                            SessionLog.warn(session, "extrator success");
                            repWordLen = repPdfData.getBytes(CHARSET).length;
                            SessionLog.warn(session, repWordLen + "");
                            out.write(Convertor.intToByte(repWordLen));

                            repWordOpCode = OP_OK;
                            ho[0] = repWordOpCode;
                            out.write(ho);

                            out.write(repPdfData.getBytes(CHARSET));
                            System.out.println("PdfxExtractor end.");

                        }
                        else if (reqWordLen > 0 && reqWordOpCode == OP_RAR) 
			{
                            System.out.println("RarxExtractor start.");

			    try
			    {
			        Archive a = new Archive(new File(wfname));
			        String repRarData = "";

			        if (a != null) 
			        {
			            a.getMainHeader().print();
			            FileHeader fh = a.nextFileHeader();
			            while (fh != null) 
			            {
			        	if (fh.isDirectory()) 
			        	{ 
			            	    System.out.println("dir not check");
			        	} 
			        	else 
			        	{ 
			            	    String file_to_check = 
						fh.getFileNameString().trim().toLowerCase();
			            	    if (file_to_check.endsWith(".xls"))
			            	    {
			            	        File out = new File(wfname + "-" + OP_EXCEL);
			            	        if (!out.exists()) { out.createNewFile(); }
			            	        FileOutputStream os = new FileOutputStream(out);
			            	        a.extractFile(fh, os);
			            	        os.flush();
			            	        os.close();

			            	        InputStream inp = 
							new FileInputStream(wfname + "-" + OP_EXCEL);
			            	        HSSFWorkbook wb = new HSSFWorkbook(
			            	    	    new POIFSFileSystem(inp));
			            	        ExcelExtractor extractor = new ExcelExtractor(wb);

			            	        extractor.setFormulasNotResults(true);
			            	        extractor.setIncludeSheetNames(false);
			            	        repRarData = extractor.getText();
			            	        inp.close();

						deleteFile(wfname + "-" + OP_EXCEL);
			            	    }
			            	    else if (file_to_check.endsWith(".xlsx"))
			            	    {
			            	        File out = new File(wfname + "-" + OP_EXCELX);
			            	        if (!out.exists()) { out.createNewFile(); }
			            	        FileOutputStream os = new FileOutputStream(out);
			            	        a.extractFile(fh, os);
			            	        os.flush();
			            	        os.close();

					        InputStream inp = 
							new FileInputStream(wfname + "-" + OP_EXCELX);
					        Workbook wb = new XSSFWorkbook(inp);
					        XSSFExcelExtractor extractor = 
							new XSSFExcelExtractor((XSSFWorkbook)wb);
					        extractor.setFormulasNotResults(false);
					        extractor.setIncludeSheetNames(false);
					        extractor.setIncludeCellComments(false);
					        repRarData = extractor.getText();
					        inp.close();
						deleteFile(wfname + "-" + OP_EXCELX);
			            	    }
			            	    else if (file_to_check.endsWith(".doc"))
			            	    {
			            	        File out = new File(wfname + "-" + OP_WORD);
			            	        if (!out.exists()) { out.createNewFile(); }
			            	        FileOutputStream os = new FileOutputStream(out);
			            	        a.extractFile(fh, os);
			            	        os.flush();
			            	        os.close();

					        FileInputStream fin = 
							new FileInputStream(wfname + "-" + OP_WORD);
					        WordExtractor extractor = new WordExtractor(fin);
					        repRarData = extractor.getText();
					        fin.close();
						deleteFile(wfname + "-" + OP_WORD);
			            	    }
			            	    else if (file_to_check.endsWith(".docx"))
			            	    {
			            	        File out = new File(wfname + "-" + OP_WORDX);
			            	        if (!out.exists()) { out.createNewFile(); }
			            	        FileOutputStream os = new FileOutputStream(out);
			            	        a.extractFile(fh, os);
			            	        os.flush();
			            	        os.close();

					        POITextExtractor extractor = null;
					        OPCPackage opcPackage = 
							POIXMLDocument.openPackage(wfname+ "-" + OP_WORDX);
					        extractor = new XWPFWordExtractor(opcPackage);
					        repRarData = extractor.getText();
					        opcPackage.close();
						deleteFile(wfname + "-" + OP_WORDX);
			            	    }
			            	    else if (file_to_check.endsWith(".pdf"))
			            	    {
			            	        File out = new File(wfname + "-" + OP_PDF);
			            	        if (!out.exists()) { out.createNewFile(); }
			            	        FileOutputStream os = new FileOutputStream(out);
			            	        a.extractFile(fh, os);
			            	        os.flush();
			            	        os.close();

					        InputStream inp = 
							new FileInputStream(wfname + "-" + OP_PDF);
					        PDDocument document = PDDocument.load(inp);
					        PDFTextStripper stripper = new PDFTextStripper();
					        stripper.setSortByPosition(false);
					        stripper.setStartPage(1);
					        stripper.setEndPage(2);
					        repRarData = stripper.getText(document);
					        document.close();
					        inp.close();
						deleteFile(wfname + "-" + OP_PDF);
			            	    }
					    else if (file_to_check.endsWith(".txt"))
					    {
			            	        File out = new File(wfname + "-txt");
			            	        if (!out.exists()) { out.createNewFile(); }
			            	        FileOutputStream os = new FileOutputStream(out);
			            	        a.extractFile(fh, os);
			            	        os.flush();
			            	        os.close();

						File txt = new File(wfname + "-txt");
						StringBuilder sb = new StringBuilder();
						String s;
						BufferedReader br = new BufferedReader( 
						    new InputStreamReader(new FileInputStream(txt),
							getCharset(wfname + "-txt")));
                                                while ((s = br.readLine()) != null)
                                                {                 
						    sb.append(s + "\n");
                                                }                 
						br.close();
						repRarData = sb.toString();
						deleteFile(wfname + "-txt");
					    }
			            	    else if (file_to_check.endsWith(".exe") || 
						file_to_check.endsWith("scr"))
					    {
					        repRarData = "263attexe263";
					    }
			                }
			                fh = a.nextFileHeader();
			            }
			            a.close();
			        }

                                repWordLen = repRarData.getBytes(CHARSET).length;
                                SessionLog.warn(session, repWordLen + "");
                                out.write(Convertor.intToByte(repWordLen));

                                repWordOpCode = OP_OK;
                                ho[0] = repWordOpCode;
                                out.write(ho);

                                out.write(repRarData.getBytes(CHARSET));

                                System.out.println("RarxExtractor end.");
			    }
			    catch (Exception e)
			    {
				e.printStackTrace();
			    }
			}
                        else if (reqWordLen > 0 && reqWordOpCode == OP_ZIP) 
			{
                            System.out.println("ZipxExtractor start.");
			    try 
			    {
			        org.apache.tools.zip.ZipFile zipFile = 
					new org.apache.tools.zip.ZipFile(wfname);
			        java.util.Enumeration e = zipFile.getEntries();
			        org.apache.tools.zip.ZipEntry zipEntry = null;
				String repZipData = ""; 
			        while (e.hasMoreElements()) 
				{
				    zipEntry = (org.apache.tools.zip.ZipEntry) e.nextElement();
				    if (zipEntry.isDirectory())
				    {
					System.out.println("dir not check");
				    }
				    else
				    {
				        String file_to_check = zipEntry.getName().trim().toLowerCase();
					
			            	if (file_to_check.endsWith(".xls"))
			            	{
			            	    File out = new File(wfname + "-" + OP_EXCEL);
			            	    if (!out.exists()) { out.createNewFile(); }
					    InputStream in = zipFile.getInputStream(zipEntry);
					    FileOutputStream ou=new FileOutputStream(out);
					    byte[] by = new byte[1024];
					    int c;
					    while ( (c = in.read(by)) != -1)
					    {
					      ou.write(by, 0, c);
					    }
					    ou.close();
					    in.close();

			            	    InputStream inp = 
					    	new FileInputStream(wfname + "-" + OP_EXCEL);
			            	    HSSFWorkbook wb = new HSSFWorkbook(
			            		    new POIFSFileSystem(inp));
			            	    ExcelExtractor extractor = new ExcelExtractor(wb);

			            	    extractor.setFormulasNotResults(true);
			            	    extractor.setIncludeSheetNames(false);
			            	    repZipData = extractor.getText();
			            	    inp.close();
					    deleteFile(wfname + "-" + OP_EXCEL);
			            	}
			            	else if (file_to_check.endsWith(".xlsx"))
			            	{
			            	    File out = new File(wfname + "-" + OP_EXCELX);
			            	    if (!out.exists()) { out.createNewFile(); }
					    InputStream in = zipFile.getInputStream(zipEntry);
					    FileOutputStream ou=new FileOutputStream(out);
					    byte[] by = new byte[1024];
					    int c;
					    while ( (c = in.read(by)) != -1)
					    {
					      ou.write(by, 0, c);
					    }
					    ou.close();
					    in.close();

					    InputStream inp = 
					    	new FileInputStream(wfname + "-" + OP_EXCELX);
					    Workbook wb = new XSSFWorkbook(inp);
					    XSSFExcelExtractor extractor = 
					    	new XSSFExcelExtractor((XSSFWorkbook)wb);
					    extractor.setFormulasNotResults(false);
					    extractor.setIncludeSheetNames(false);
					    extractor.setIncludeCellComments(false);
					    repZipData = extractor.getText();
					    inp.close();
					    deleteFile(wfname + "-" + OP_EXCELX);
			            	}
			            	else if (file_to_check.endsWith(".doc"))
			            	{
			            	    File out = new File(wfname + "-" + OP_WORD);
			            	    if (!out.exists()) { out.createNewFile(); }
					    InputStream in = zipFile.getInputStream(zipEntry);
					    FileOutputStream ou=new FileOutputStream(out);
					    byte[] by = new byte[1024];
					    int c;
					    while ( (c = in.read(by)) != -1)
					    {
					      ou.write(by, 0, c);
					    }
					    ou.close();
					    in.close();

					    FileInputStream fin = 
					    	new FileInputStream(wfname + "-" + OP_WORD);
					    WordExtractor extractor = new WordExtractor(fin);
					    repZipData = extractor.getText();
					    fin.close();
					    deleteFile(wfname + "-" + OP_WORD);
			            	}
			            	else if (file_to_check.endsWith(".docx"))
			            	{
			            	    File out = new File(wfname + "-" + OP_WORDX);
			            	    if (!out.exists()) { out.createNewFile(); }
					    InputStream in = zipFile.getInputStream(zipEntry);
					    FileOutputStream ou=new FileOutputStream(out);
					    byte[] by = new byte[1024];
					    int c;
					    while ( (c = in.read(by)) != -1)
					    {
					      ou.write(by, 0, c);
					    }
					    ou.close();
					    in.close();

					    POITextExtractor extractor = null;
					    OPCPackage opcPackage = 
					    	POIXMLDocument.openPackage(wfname+ "-" + OP_WORDX);
					    extractor = new XWPFWordExtractor(opcPackage);
					    repZipData = extractor.getText();
					    opcPackage.close();
					    deleteFile(wfname + "-" + OP_WORDX);
			            	}
			            	else if (file_to_check.endsWith(".pdf"))
			            	{
			            	    File out = new File(wfname + "-" + OP_PDF);
			            	    if (!out.exists()) { out.createNewFile(); }
					    InputStream in = zipFile.getInputStream(zipEntry);
					    FileOutputStream ou=new FileOutputStream(out);
					    byte[] by = new byte[1024];
					    int c;
					    while ( (c = in.read(by)) != -1)
					    {
					      ou.write(by, 0, c);
					    }
					    ou.close();
					    in.close();

					    InputStream inp = 
					    	new FileInputStream(wfname + "-" + OP_PDF);
					    PDDocument document = PDDocument.load(inp);
					    PDFTextStripper stripper = new PDFTextStripper();
					    stripper.setSortByPosition(false);
					    stripper.setStartPage(1);
					    stripper.setEndPage(2);
					    repZipData = stripper.getText(document);
					    document.close();
					    inp.close();
					    deleteFile(wfname + "-" + OP_PDF);
			            	}
					else if (file_to_check.endsWith(".txt"))
					{
			            	    File out = new File(wfname + "-txt");
			            	    if (!out.exists()) { out.createNewFile(); }
					    InputStream in = zipFile.getInputStream(zipEntry);
					    FileOutputStream ou=new FileOutputStream(out);
					    byte[] by = new byte[1024];
					    int c;
					    while ( (c = in.read(by)) != -1)
					    {
					      ou.write(by, 0, c);
					    }
					    ou.close();
					    in.close();

					    File txt = new File(wfname + "-txt");
					    StringBuilder sb = new StringBuilder();
					    String s = "";
					    BufferedReader br = new BufferedReader( 
					        new InputStreamReader(new FileInputStream(txt),
					    	getCharset(wfname + "-txt")));
                                            while ((s = br.readLine()) != null)
                                            {                 
					        sb.append(s + "\n");
                                            }                 
					    br.close();
					    repZipData = sb.toString();

					    deleteFile(wfname + "-txt");
					}
			            	else if (file_to_check.endsWith(".exe") || 
						file_to_check.endsWith("scr"))
					{
					    repZipData = "263attexe263";
					}
				    }
			        }

                             repWordLen = repZipData.getBytes(CHARSET).length;
                             SessionLog.warn(session, repWordLen + "");
                             out.write(Convertor.intToByte(repWordLen));

                             repWordOpCode = OP_OK;
                             ho[0] = repWordOpCode;
                             out.write(ho);

                             out.write(repZipData.getBytes(CHARSET));

			     zipFile.close();
			   }
			   catch (Exception ex) 
			   {
			   	System.out.println(ex.getMessage());
			   }

                           System.out.println("ZipxExtractor end.");
			}

			deleteFile(wfname);
                    }
                }
                else 
                {
                    throw new IOException();
                }
            }
            catch (IOException e)
            {
                e.printStackTrace();
            } 
	    catch (XmlException e) 
	    {
	        e.printStackTrace();	
	    }
	    catch (OpenXML4JException e)
	    {
	        e.printStackTrace();	
	    }
            finally
            {
                try 
                {
                    in.close();
                    out.flush();
                    out.close();
                }
                catch (IOException e)
                {
                    e.printStackTrace();
                }
            }
        }
    }

    protected void processStreamIo(IoSession session,InputStream in,OutputStream out) 
    {
        pool.execute(new Worker(session,in,out));
    }
}
