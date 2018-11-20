import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.logging.Level;
import java.util.logging.Logger;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Image;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.PdfArray;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfDictionary;
import com.itextpdf.text.pdf.PdfImportedPage;
import com.itextpdf.text.pdf.PdfName;
import com.itextpdf.text.pdf.PdfNumber;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;
import com.itextpdf.text.pdf.PdfTemplate;
import com.itextpdf.text.pdf.PdfWriter;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;


public class Ppt2pdf {
	private static Logger loger;
	private static File directory;
	public static void main(String[] args) {
		
		loger = Logger.getLogger("This is log for developer!");
		// convert ppt/pptx to pdf by wps or ms office
		directory = new File("D://PPT/PDF");
		if (!directory.exists()) {
			directory.mkdirs();
		}
		//step 1:
		//get all file in the directory
		ArrayList<String> al = new ArrayList<String>();
		String pathName="D://PPT";
		File dirFile = new File(pathName);
		//判断该文件或目录是否存在，不存在时在控制台输出提醒
		if (!dirFile.exists()) {
			System.out.println("do not exit");
			return ;
		}
		String[] fileList = dirFile.list();
		for (int i = 0; i < fileList.length; i++) {
			String string = fileList[i];
			//File("documentName","fileName")是File的另一个构造器
			File file = new File(dirFile.getPath(),string);
			String name = file.getName();
			String path = file.getAbsolutePath();
			System.out.println(path);
			if(path.endsWith("ppt") || path.endsWith("PPT") || path.endsWith("pptx") || path.endsWith("PPTX")){
			
				
				ExecutorService exec = Executors.newCachedThreadPool();
				Future<Boolean> future = exec.submit(new Ppt2pdf().new PDFUtilCallable(path, path+".pdf"));
				String pdfPath  = directory.getAbsolutePath() + File.separatorChar + "New"+name+".pdf";
				al.add(pdfPath);
				try {
					Boolean backCode = future.get();
					if(backCode){
						zoomMerge(path + ".pdf","baheyi");
						
						//zoomMerge(path + ".pdf","liuheyi");
					}
					else{
						loger.log(Level.SEVERE,"转换失败！");
					}
					
				} catch (InterruptedException e) {
					e.printStackTrace();
				} catch (ExecutionException e) {
					e.printStackTrace();
				}
			}
		}
		
		/**
		 *  合成一张pdf打印，有待开发！
		 */
		/*
		if(compose(al)){
			loger.log(Level.INFO, "合并完成！");
		}
		*/
		
		
	}
	
	
	static String  zoomMerge(String src, String zoom) {
		loger.log(Level.INFO,"开始合成");

		File fisrtfile = new File(src);
		String zoomPath = directory.getAbsolutePath() + File.separatorChar + "Zoom" + fisrtfile.getName();
		String targetFile = directory.getAbsolutePath() + File.separatorChar + "New" + fisrtfile.getName();
		try {
			PdfReader reader1 = new PdfReader(src);
			System.out.println("Before Zoom Orientation：" + reader1.getPageRotation(1));
			Rectangle rectangle = reader1.getPageSizeWithRotation(1);
			// rectangle.rotate();
			// System.out.println("rectangle
			// rotation开始:"+rectangle.getRotation());
			// rectangle.setRotation(0);

			// System.out.println("rectangle
			// rotation结尾:"+rectangle.getRotation());
			float height = rectangle.getHeight();// height 841
			float width = rectangle.getWidth();// width 595

			loger.log(Level.INFO, "Zoom before Height:" + height + "\nZoom before  Width:" + width);
			PdfStamper stamper1 = new PdfStamper(reader1, new FileOutputStream(zoomPath));
			int n = reader1.getNumberOfPages();

			PdfDictionary page1;
			PdfArray crop1;
			// PdfArray media;
			for (int p = 1; p <= n; p++) {
				page1 = reader1.getPageN(p);
				// page1.remove(PdfName.ROTATE);//去除方向
				/*
				 * media = page1.getAsArray(PdfName.CROPBOX); if (media ==
				 * null) { media = page1.getAsArray(PdfName.MEDIABOX); }
				 */
				PdfNumber rotate = page1.getAsNumber(PdfName.ROTATE);
				// System.out.println("Check the rotation：" + rotate);
				loger.log(Level.INFO, "Check the rotation：" + rotate);
				// 正常情况下这个是0或者null，如果不正常则把视觉方向与机器方向统一。
				if (rotate != null) {
					if (rotate.intValue() != 0) {
						page1.put(PdfName.ROTATE, new PdfNumber(360 - rotate.intValue()));
						rotate = page1.getAsNumber(PdfName.ROTATE);
						// System.out.println("after change ：" + rotate);
						loger.log(Level.INFO, "after change ：" + rotate);
					} else
						;
				} else
					;
				crop1 = new PdfArray();
				crop1.add(new PdfNumber(0));
				crop1.add(new PdfNumber(0));

				if (zoom.equals("liuheyi")) {

					float percentage = 280.0f / width;
					crop1.add(new PdfNumber(280.0f));// width
					crop1.add(new PdfNumber(height * percentage));// height
					page1.put(PdfName.MEDIABOX, crop1);
					page1.put(PdfName.CROPBOX, crop1);

					stamper1.getUnderContent(p)
							.setLiteral(String.format("\nq %s 0 0 %s %s %s cm\nq\n", percentage, percentage, 0, 0));

				} else if (zoom.equals("baheyi")) {
					float percentage = 260.0f / width;
					crop1.add(new PdfNumber(width * percentage));
					crop1.add(new PdfNumber(195.0));
					page1.put(PdfName.MEDIABOX, crop1);
					page1.put(PdfName.CROPBOX, crop1);
					stamper1.getUnderContent(p)
							.setLiteral(String.format("\nq %s 0 0 %s %s %s cm\nq\n", percentage, percentage, 0, 0));

				}

				stamper1.getOverContent(p).setLiteral("\nQ\nQ\n");
			}

			stamper1.close();
			reader1.close();
		} catch (DocumentException e1) {
			e1.printStackTrace();
		} catch (IOException e1) {
			e1.printStackTrace();
		}

		PdfReader reader = null;
		try {
			reader = new PdfReader(zoomPath);

			Document document = new Document(PageSize.A4);// A4 paper Size

			PdfWriter pdfWriter = PdfWriter.getInstance(document, new FileOutputStream(targetFile));
			document.open();
			Image page5 = null;
			for (int i = 1; i <= reader.getNumberOfPages(); i++) {

				page5 = Image.getInstance(pdfWriter.getImportedPage(reader, i));
				if (zoom.equals("liuheyi")) {

					float x = 20;
					float y = 570;
					float pageHeight = 210f;

					float pageWidth = 280.0f;
					float x1 = 15;
					float y1 = 40;

					switch (i % 6) {
					case 1:
						// x=25;
						// y=580;
						break;

					case 2:
						x = x + pageWidth - x1;
						// y=580;
						break;
					case 3:
						// x=25;
						y = y - pageHeight - y1;
						break;
					case 4:
						x = x + pageWidth - x1;
						y = y - pageHeight - y1;
						break;
					case 5:
						// x=25;
						y = y - 2 * (pageHeight + y1);
						break;
					case 0:
						x = x + pageWidth - x1;
						y = y - 2 * (pageHeight + y1);

						break;

					}
					page5.setAbsolutePosition(x, y);
					// page5.scalePercent(scale * 100);
					document.add(page5);

					if (i % 6 == 0) {
						document.newPage();
					}

				} else if (zoom.equals("baheyi")) {

					float x = 30;
					float y = 625;
					float pageHeight = 200f;
					float pageWidth = 280.0f;
					float x1 = 1;
					float y1 = 1;

					switch (i % 8) {
					case 1:
						// x=25;
						// y=580;
						break;

					case 2:
						x = x + pageWidth - x1;
						// y=580;
						break;
					case 3:
						// x=25;
						y = y - pageHeight - y1;
						break;
					case 4:
						x = x + pageWidth - x1;
						y = y - pageHeight - y1;
						break;
					case 5:
						// x=25;
						y = y - 2 * (pageHeight + y1);
						break;
					case 6:

						x = x + pageWidth - x1;
						y = y - 2 * (pageHeight + y1);
						break;
					case 7:

						y = y - 3 * (pageHeight + y1);
						break;
					case 0:
						x = x + pageWidth - x1;
						y = y - 3 * (pageHeight + y1);

						break;

					}

					page5.setAbsolutePosition(x, y);
					// page5.scalePercent(scale * 100);
					document.add(page5);

					if (i % 8 == 0) {
						// System.out.println("the " + i + " pages，need new
						// a blank paper");
						document.newPage();
					}

					/*
					 * canvas.addTemplate(page, x, y); try {
					 * stamper.getWriter().freeReader(r); } catch
					 * (IOException e) { loger.log(Level.SEVERE,
					 * e.getMessage()); }
					 * 
					 * canvas = stamper.getUnderContent(i / 8 + 1);
					 */
				}

			}
			document.close();
			pdfWriter.close();
			reader.close();
			// stamper.close();
		} catch (DocumentException e) {
			loger.log(Level.SEVERE, e.getMessage());
		} catch (IOException e) {
			loger.log(Level.SEVERE, e.getMessage());
		}
		File srcFile = new File(src);
		if (srcFile.delete()) {
			// System.out.println("The Zoom file delete rightly");
			loger.log(Level.INFO, "The srcpdf file delete rightly");
		} else
			;
		// delete the zoom file
		File zoomDelete = new File(zoomPath);
		if (zoomDelete.delete()) {
			// System.out.println("The Zoom file delete rightly");
			loger.log(Level.INFO, "The Zoom file delete rightly");
		} else
			;
		return targetFile;

	}
	
	
	static boolean compose(ArrayList<String> al){
		PdfReader reader=null;
		String dest =directory.getAbsolutePath()+"最后的pdf.pdf";
		try {
			reader = new PdfReader(al.get(0));
		    Rectangle pageSize = reader.getPageSize(1);
		    
	        //遍历每一个pdf
		    for(int i=0;i<al.size();i++){
		    	reader = new PdfReader(al.get(i));
		    	int pagesNumber = reader.getNumberOfPages();
		    	//遍历pdf的吗，每一页
		    	for(int j = 1 ;j<=pagesNumber;j++){
		    		Document document = new Document(pageSize);
				    PdfWriter writer  = PdfWriter.getInstance(document, new FileOutputStream(dest));
			        document.open();
			        PdfImportedPage page = writer.getImportedPage(reader, j);
			        PdfContentByte cb = writer.getDirectContent();
			        PdfTemplate template1 = cb.createTemplate(pageSize.getWidth(), pageSize.getHeight());
			        template1.rectangle(0, 0, pageSize.getWidth(), pageSize.getHeight());
			        template1.eoClip();
			        template1.newPath();
			        template1.addTemplate(page, 0, 0);
			        cb.addTemplate(template1, 0, 0);

				    document.close();
		    	}
		    	
		    	  reader.close();
			}
		  
			return true;
		} catch (IOException | DocumentException e1) {
			e1.printStackTrace();
			return false;
		}
	
	}

class PDFUtilCallable implements Callable<Boolean> {
	
	private static final int wdFormatPDF = 17;
	private static final int xlTypePDF = 0;
	private static final int ppSaveAsPDF = 32;
	String inputFile;
	String pdfFile;

	PDFUtilCallable(String inputFile, String pdfFile) {
		this.inputFile = inputFile;
		this.pdfFile = pdfFile;
	}

	@Override
	public Boolean call() throws Exception {

		String suffix = getFileSufix(inputFile);
		File file = new File(inputFile);
		if (!file.exists()) {
			return false;
		}
		if (suffix.equals("pdf")) {
			return false;
		}
		if (suffix.equals("doc") || suffix.equals("docx") || suffix.equals("txt")) {
			return word2PDF(inputFile, pdfFile);
		} else if (suffix.equals("ppt") || suffix.equals("pptx")) {
			return ppt2PDF(inputFile, pdfFile);
		} else if (suffix.equals("xls") || suffix.equals("xlsx")) {
			return excel2PDF(inputFile, pdfFile);
		} else {
			return false;
		}
	}

	public String getFileSufix(String fileName) {
		int splitIndex = fileName.lastIndexOf(".");
		return fileName.substring(splitIndex + 1);
	}

	// word转换为pdf
	public boolean word2PDF(String inputFile, String pdfFile) {
		loger.log(Level.INFO, "正在调用函数jacob word2PDF()");
		loger.log(Level.INFO, "inputFile:" + inputFile + "\tpdfFile:" + pdfFile);

		File file = new File(inputFile);
		loger.log(Level.INFO, "判断文件是否存在:" + file.exists());
		Dispatch doc = null;
		Dispatch docs = null;
		try {

			loger.log(Level.INFO, "启动Word.Application...");
			long start = System.currentTimeMillis();
			ActiveXComponent app = null;

			try {
				app = new ActiveXComponent("Word.Application");
				app.setProperty("Visible", new Variant(false));
				docs = app.getProperty("Documents").toDispatch();
				// doc = Dispatch.call(docs, "Open" ,
				// sourceFile).toDispatch();
				doc = Dispatch
						.invoke(docs, "Open", Dispatch.Method,
								new Object[] { inputFile, new Variant(false), new Variant(true) }, new int[1])
						.toDispatch();
				loger.log(Level.INFO, "转化之前：" + inputFile + "\t转换文档到PDF:" + pdfFile);
				File tofile = new File(pdfFile);
				if (tofile.exists()) {
					tofile.delete();
				}
				// Dispatch.call(doc, "SaveAs", destFile, 17);
				Dispatch.invoke(doc, "SaveAs", Dispatch.Method, new Object[] { pdfFile, new Variant(17) },
						new int[1]);
				long end = System.currentTimeMillis();
				loger.log(Level.INFO, "转换完成..用时" + (end - start) + "ms.");
			} catch (Exception e) {
				e.printStackTrace();
				loger.log(Level.SEVERE, "========Error:文档转换失败" + e.getMessage());
				// 尝试第二方案转化

				// 打开word应用程序
				ActiveXComponent app1 = new ActiveXComponent("Word.Application"); // 设置word不可见
				app1.setProperty("Visible", false); // 获得word中所有打开的文档,返回Documents对象
													// Dispatch docs =
				app1.getProperty("Documents").toDispatch(); // 调用Documents对象中Open方法打开文档，并返回打开的文档对象Document
															// Dispatch
															// doc
															// =
				Dispatch.call(docs, "Open", inputFile.trim(), false, true).toDispatch(); // 调用Document对象的SaveAs方法，将文档保存为pdf格式

				Dispatch.call(doc, "SaveAs", pdfFile, wdFormatPDF);// word保存为pdf格式宏，值为17

				// Dispatch.call(doc, "ExportAsFixedFormat",
				// pdfFile,wdFormatPDF);// word保存为pdf格式宏，值为17
				// 关闭文档
				Dispatch.call(doc, "Close", false); // 关闭word应用程序
													// app.invoke("Quit",
													// 0);

			} finally {
				Dispatch.call(doc, "Close", false);
				if (app != null)
					app.invoke("Quit", new Variant[] {});
			}
			ComThread.Release();

			loger.log(Level.INFO, "jacob 调用office组件word转换成功！");
			return true;
		} catch (Exception e) {
			loger.log(Level.SEVERE, "jacob 调用office组件word发生错误！");
			loger.log(Level.SEVERE, e.getMessage());
			return false;
		}
	}

	// excel转换为pdf
	public boolean excel2PDF(String inputFile, String pdfFile) {
		try {
			ActiveXComponent app = new ActiveXComponent("Excel.Application");
			app.setProperty("Visible", false);
			Dispatch excels = app.getProperty("Workbooks").toDispatch();
			Dispatch excel = Dispatch.call(excels, "Open", inputFile, false, true).toDispatch();
			Dispatch.call(excel, "ExportAsFixedFormat", xlTypePDF, pdfFile);
			Dispatch.call(excel, "Close", false);
			app.invoke("Quit");
			return true;
		} catch (Exception e) {
			loger.log(Level.SEVERE, e.getMessage());
			return false;
		}
	}

	// ppt转换为pdf
	public boolean ppt2PDF(String inputFile, String pdfFile) {
		loger.log(Level.INFO, "jacob ppt2PDF()");
		try {
			ActiveXComponent app = new ActiveXComponent("PowerPoint.Application");
			// app.setProperty("Visible", msofalse);
			Dispatch ppts = app.getProperty("Presentations").toDispatch();
			Dispatch ppt = Dispatch.call(ppts, "Open", inputFile, true, // ReadOnly
					true, // Untitled指定文件是否有标题
					false// WithWindow指定文件是否可见
			).toDispatch();
			Dispatch.call(ppt, "SaveAs", pdfFile, ppSaveAsPDF);
			Dispatch.call(ppt, "Close");
			app.invoke("Quit");
			return true;
		} catch (Exception e) {
			loger.log(Level.SEVERE, e.getMessage());
			return false;
		}
	}

}

}