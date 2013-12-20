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

        public void run() 
        {
            try
	    	{
				String wfname = String.valueOf(Thread.currentThread().getId());

				Result ret = saveFile(wfname);
				int reqWordLen = ret.getLen();
				byte reqWordOpCode = ret.getType();

				if (reqWordLen <= 0) 
				{
					deleteFile(wfname);
					return;
				}
				
				switch (reqWordOpCode)
				{
					case OP_WORD:
						processWord(wfname); break;
					case OP_WORDX:
						processWordx(wfname); break;
					case OP_EXCEL:
						processExcel(wfname); break;
					case OP_EXCELX:
						processExcelx(wfname); break;
					case OP_PDF:
						processPDF(wfname); break;
					case OP_RAR:
						processRar(wfname); break;
					case OP_ZIP:
						processZip(wfname); break;
					default:
						System.out.println("type not support.");
				}

				deleteFile(wfname);
			}
            catch (Exception e)
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
		
		public void processZip(String wfname)
		{
             System.out.println("ZipxExtractor start.");
			 org.apache.tools.zip.ZipFile zipFile = null;
			 try 
			 {
			 	zipFile = new org.apache.tools.zip.ZipFile(wfname);
			 	java.util.Enumeration e = zipFile.getEntries();
			 	org.apache.tools.zip.ZipEntry zipEntry = null;
			 	String repZipData = ""; 
			 	while (e.hasMoreElements()) 
			 	{
			 		zipEntry = (org.apache.tools.zip.ZipEntry) e.nextElement();
			 		if (zipEntry.isDirectory())
			 		{
			 			System.out.println("dir not check");
						continue;
			 		}
			 		else
			 		{
			 	    	String file_to_check = zipEntry.getName().trim().toLowerCase();
			 	
			 	    	if (file_to_check.endsWith(".xls"))
			 	    	{
			 				saveZipFile(zipFile,zipEntry,wfname + "-" + OP_EXCEL);
			 				repZipData = parseExcel(wfname + "-" + OP_EXCEL);
			 				deleteFile(wfname + "-" + OP_EXCEL);
							break;
			 			}
			 	    	else if (file_to_check.endsWith(".xlsx"))
			 	    	{
			 				saveZipFile(zipFile,zipEntry,wfname + "-" + OP_EXCELX);
			 				repZipData = parseExcelx(wfname + "-" + OP_EXCELX);
			 				deleteFile(wfname + "-" + OP_EXCELX);
							break;
			 	    	}
			 	    	else if (file_to_check.endsWith(".doc"))
			 	    	{
			 				saveZipFile(zipFile,zipEntry,wfname + "-" + OP_WORD);
			 				repZipData = parseWord(wfname + "-" + OP_WORD);
			 				deleteFile(wfname + "-" + OP_WORD);
							break;
			 	    	}
			 	    	else if (file_to_check.endsWith(".docx"))
			 	    	{
			 				saveZipFile(zipFile,zipEntry,wfname + "-" + OP_WORDX);
			 				repZipData = parseWordx(wfname + "-" + OP_WORDX);
			 				deleteFile(wfname + "-" + OP_WORDX);
							break;
			 	    	}
			 	    	else if (file_to_check.endsWith(".pdf"))
			 	    	{
			 				saveZipFile(zipFile,zipEntry,wfname + "-" + OP_PDF);
			 				repZipData = parsePDF(wfname + "-" + OP_PDF);
			 				deleteFile(wfname + "-" + OP_PDF);
							break;
			 	    	}
			 			else if (file_to_check.endsWith(".txt"))
			 			{
			 				saveZipFile(zipFile,zipEntry,wfname + "-txt");
			 				repZipData = getTxt(wfname + "-txt");
			 				deleteFile(wfname + "-txt");
							break;
			 			}
			 	    	else if (file_to_check.endsWith(".exe") || 
			 				file_to_check.endsWith("scr"))
			 			{
			 	    		repZipData = "263attexe263";
							break;
			 			}
			 		}
			 	}

			 	sessionResponse(repZipData); 
             	System.out.println("ZipxExtractor end.");
			 }
			 catch (Exception e) 
			 {
			 	System.out.println(e.getMessage());
			 }
			 finally
			 {
				try
				{
					zipFile.close();
				}
				catch (Exception e)
				{
					e.printStackTrace();
				}
			 }
             System.out.println("ZipxExtractor end.");
		}

		public void processRar(String fileName)
		{
             System.out.println("RarxExtractor start.");
			 Archive a = null;
			 try
			 {
			 	 File file = new File(fileName);
			     a = new Archive(file);
			     String repRarData = "";
			     if (a != null) 
			     {
			         FileHeader fh = a.nextFileHeader();
			         while (fh != null) 
			         {
			 			if (fh.isDirectory()) 
			 			{ 
			 				System.out.println("dir not check");
			 			} 
			 			else 
			 			{ 
			 				String file_to_check = fh.getFileNameString().trim().toLowerCase();
			 				if (file_to_check.endsWith(".xls"))
			 				{
			 					saveRarFile(a,fh,fileName + "-" + OP_EXCEL);
			 					repRarData = parseExcel(fileName + "-" + OP_EXCEL);
			 					deleteFile(fileName + "-" + OP_EXCEL);
								break;
			 				}
			 				else if (file_to_check.endsWith(".xlsx"))
			 				{
			 					saveRarFile(a,fh,fileName + "-" + OP_EXCELX);
			 					repRarData = parseExcelx(fileName + "-" + OP_EXCELX);
			 					deleteFile(fileName + "-" + OP_EXCELX);
								break;
			 				}
			 				else if (file_to_check.endsWith(".doc"))
			 				{
			 					saveRarFile(a,fh,fileName + "-" + OP_WORD);
			 					repRarData = parseWord(fileName + "-" + OP_WORD);
			 					deleteFile(fileName + "-" + OP_WORD);
								break;
			 				}
			 				else if (file_to_check.endsWith(".docx"))
			 				{
			 					saveRarFile(a,fh,fileName + "-" + OP_WORDX);
			 					repRarData = parseWordx(fileName + "-" + OP_WORDX);
			 					deleteFile(fileName + "-" + OP_WORDX);
								break;
			 				}
			 				else if (file_to_check.endsWith(".pdf"))
			 				{
			 					saveRarFile(a,fh,fileName + "-" + OP_PDF);
			 					repRarData = parsePDF(fileName + "-" + OP_PDF);
			 					deleteFile(fileName + "-" + OP_PDF);
								break;
			 				}
			 				else if (file_to_check.endsWith(".txt"))
			 				{
			 					saveRarFile(a,fh,fileName + "-txt");
			 					repRarData = getTxt(fileName + "-txt");
			 					deleteFile(fileName + "-txt");
								break;
			 				}
			 				else if (file_to_check.endsWith(".exe") || 
			 					file_to_check.endsWith("scr"))
			 				{
			 					repRarData = "263attexe263";
								break;
			 				}
			 			}

			 			fh = a.nextFileHeader();
			        }
			    }

			    sessionResponse(repRarData);
                System.out.println("RarxExtractor end.");
			}
			catch (Exception e)
			{
				e.printStackTrace();
			}
			finally 
			{
				try
				{
					a.close();
				}
				catch (Exception e)
				{
					e.printStackTrace();
				}
			}
			
		}

		public void processWord(String fileName)
		{
            System.out.println("WordExtractor start.");
			String repWordData = parseWord(fileName);
			sessionResponse(repWordData);
            System.out.println("WordExtractor end.");
		}
		
		public void processWordx(String fileName)
		{
            System.out.println("WordxExtractor start.");
            String repWordData = parseWordx(fileName);
			sessionResponse(repWordData);
            System.out.println("WordxExtractor end.");
		}

		public void processExcel(String fileName)
		{
        	 System.out.println("ExcelExtractor start.");
			 String repExcelData = parseExcel(fileName);
			 sessionResponse(repExcelData);
        	 System.out.println("ExcelExtractor end.");
		}
		
		public void processExcelx(String fileName)
		{
             System.out.println("ExcelxExtractor start.");
			 String repExcelData = parseExcelx(fileName);
			 sessionResponse(repExcelData);
             System.out.println("ExcelxExtractor end.");
		}

		public void processPDF(String fileName)
		{
             System.out.println("PdfxExtractor start.");
			 String repPdfData = parsePDF(fileName);
			 sessionResponse(repPdfData);
             System.out.println("PdfxExtractor end.");
		}


		public void sessionResponse(String response)
		{
			try
			{ 
				byte[] ho = new byte[1];
				int responseLen = response.getBytes(CHARSET).length;
				SessionLog.warn(session, responseLen + "");
				out.write(Convertor.intToByte(responseLen));
				byte responseOpCode = OP_OK;
				ho[0] = responseOpCode;
				out.write(ho);
				out.write(response.getBytes(CHARSET));
			}
			catch (Exception e)
			{
				e.printStackTrace();
			}
			finally
			{
				try
				{
					out.flush();
					out.close();
				}
				catch (Exception e)
				{
					e.printStackTrace();
				}
			}
		}
	

		public Result saveFile(String fileName)
		{
			byte[] hl = new byte[4];
			byte[] ho = new byte[1];

	    	byte reqWordOpCode = 0;
			int reqWordLen = 0;

			try 
			{
				if (in.read(hl) != -1) 
				{
					if (in.read(ho) != -1) 
					{
						reqWordLen = Convertor.bytesToInt(hl);
						reqWordOpCode = ho[0];
						System.out.println(reqWordLen);

						if (reqWordLen > 0)
						{
							File file = new File(fileName);	    
							FileOutputStream fos = new FileOutputStream(file);
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
										fos.close();
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
										fos.close();
										throw new IOException();
									}
								}
							}
							fos.flush();
							fos.close();
						}
					}
				}
			} 
			catch (Exception e)
			{
				e.printStackTrace();
			}

			Result ret = new Result();	
			ret.setValue(reqWordOpCode,reqWordLen);
			return ret;
		}

		public void saveRarFile(Archive a,FileHeader fh,String fileName)
		{
			File ou = new File(fileName);
			if (!ou.exists()) 
			{ 
				try
				{
					ou.createNewFile();
				}
				catch (Exception e)
				{
					e.printStackTrace();
					return;
				}
			}

			FileOutputStream os = null;
			try
			{
				os = new FileOutputStream(ou);
				a.extractFile(fh, os);
			} 
			catch (Exception e) 
			{
				e.printStackTrace();
			}	
			finally
			{
				try
				{
					os.flush();
					os.close();
				}
				catch (Exception e)
				{
					e.printStackTrace();
				}
			}
		}

		public void saveZipFile(org.apache.tools.zip.ZipFile zipFile,
				org.apache.tools.zip.ZipEntry zipEntry,String fileName)
		{
			File outt = new File(fileName);
			if (!outt.exists()) 
			{ 
				try
				{
					outt.createNewFile(); 
				}
				catch (Exception e)
				{
					e.printStackTrace();
					return;
				}
			}

			InputStream in = null;
			FileOutputStream ou = null;
			try
			{
				in = zipFile.getInputStream(zipEntry);
				ou = new FileOutputStream(outt);
				byte[] by = new byte[1024];
				int c;
				while ( (c = in.read(by)) != -1)
				{
				  ou.write(by, 0, c);
				}
			}
			catch (Exception e)
			{
				e.printStackTrace();
			}
			finally
			{
				try
				{
					ou.close();
					in.close();
				}
				catch (Exception e)
				{
					e.printStackTrace();
				}
			}
		}

		public String getTxt(String fileName)
		{
			File txt = new File(fileName);
			StringBuilder sb = new StringBuilder();
			String s;
			BufferedReader br = null;
			String repRarData = "";

			try
			{
				br = new BufferedReader( 
					new InputStreamReader(new FileInputStream(txt),
					getCharset(fileName)));
				while ((s = br.readLine()) != null)
				{                 
					sb.append(s + "\n");
				}                 
				repRarData = sb.toString();
			}
			catch (Exception e)
			{
				e.printStackTrace();
			}
			finally
			{
				try
				{
					br.close();
				}
				catch (Exception e)
				{
					e.printStackTrace();
				}
			}

			return  repRarData;
		}
    }

    protected void processStreamIo(IoSession session,InputStream in,OutputStream out) 
    {
        pool.execute(new Worker(session,in,out));
    }

    
    public static String getCharset(String fileName)
    {
        String charset = ""; 
        InputStream inp = null;
        BufferedInputStream bin = null;
        try 
        {
    		File file = new File(fileName);
    		inp = new FileInputStream(file);
            bin = new BufferedInputStream(inp);
            int p = (bin.read() << 8) + bin.read();
            switch (p) 
            {
                case 0xefbb: charset = "UTF-8"; break;
                case 0xfffe: charset = "Unicode"; break;
                case 0xfeff: charset = "UTF-16BE"; break;
                default: charset = "gbk"; break;
            }
        } 
		catch (Exception e) 
		{
    		e.printStackTrace();
        } 
		finally 
		{
            try 
            {
    	    	inp.close();	
    	    	bin.close();
            }
            catch (IOException e)
            {
                e.printStackTrace();
            }
        }
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

    public static String parseWord(String fileName)
    {
        FileInputStream fin = null;
    	String repWordData = "";
    	try {
    		fin = new FileInputStream(fileName);
            WordExtractor extractor = new WordExtractor(fin);
            repWordData = extractor.getText();
    	} 
		catch (Exception e) 
		{
    	    e.printStackTrace();
    	}
		finally 
		{
    	    try 
			{
    	    	fin.close();
    	    } 
			catch (IOException e)
			{
    			e.printStackTrace();
    	    }
    	}
    
    	return repWordData;
    }

    public static String parseWordx(String fileName)
    {
		OPCPackage opcPackage = null;
		String repWordData = "";
		try
		{
			POITextExtractor extractor = null;
			opcPackage = POIXMLDocument.openPackage(fileName);
			extractor = new XWPFWordExtractor(opcPackage);
			repWordData = extractor.getText();
		} 
		catch(Exception e) 
		{
	    	e.printStackTrace();
		} 
		finally
		{
	    	try
			{
	    		opcPackage.close();
	    	} 
			catch (Exception e) 
			{
				e.printStackTrace();
	    	}
		}
		return repWordData;	
    }

    public static String parseExcel(String fileName)
    {
        InputStream inp = null;
		String repExcelData = "";
		try 
		{
			inp = new FileInputStream(fileName);
        	HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(inp));
        	ExcelExtractor extractor = new ExcelExtractor(wb);
        	extractor.setFormulasNotResults(true);
        	extractor.setIncludeSheetNames(false);
        	repExcelData = extractor.getText();
		} 
		catch (Exception e) 
		{
	    	e.printStackTrace();
		} 
		finally 
		{
	    	try 
			{
	    		inp.close();
	    	}
			catch (Exception e) 
			{
				e.printStackTrace();
	    	}
		}

		return repExcelData;
    }

    public static String parseExcelx(String fileName)
    {
        InputStream inp = null;
		String repExcelData = "";
		try 
		{
			inp = new FileInputStream(fileName);
			Workbook wb = new XSSFWorkbook(inp);
			XSSFExcelExtractor extractor = new XSSFExcelExtractor((XSSFWorkbook)wb);
			extractor.setFormulasNotResults(false);
			extractor.setIncludeSheetNames(false);
			extractor.setIncludeCellComments(false);
    	    repExcelData = extractor.getText();
		}
		catch (Exception e)
		{
		    e.printStackTrace();
		}
		finally
		{
		    try
			{
		    	inp.close();
		    } 
			catch (Exception e) 
			{
				e.printStackTrace();
		    }
		}

		return repExcelData;
    }

    public static String parsePDF(String fileName)
    {
		String repPdfData = "";
        InputStream inp = null;
		OutputStream outp = null;
		OutputStreamWriter output = null;
		try 
		{
			inp = new FileInputStream(fileName);
			PDDocument document = PDDocument.load(inp);
			PDFTextStripper stripper = new PDFTextStripper("utf-8");
			stripper.setSortByPosition(false);
			stripper.setStartPage(1);
			stripper.setEndPage(2);
			outp = new FileOutputStream(fileName + "-txt");
			output = new OutputStreamWriter(outp,"utf-8"); 
			stripper.writeText(document, output);
			repPdfData = stripper.getText(document);
			document.close();
		}
		catch (Exception e)
		{
	    	e.printStackTrace();
		}
		finally
		{
			deleteFile(fileName + "-txt");
	    	try
			{
				output.close();
				outp.close();
				inp.close();
	    	}
			catch (Exception e)
			{ 
				e.printStackTrace();
	    	}
		}
		return repPdfData;
    }
}
