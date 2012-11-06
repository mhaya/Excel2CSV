package org.sheepcloud.tool;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.Iterator;
import java.util.Locale;
import java.util.ResourceBundle;

import org.apache.commons.cli.BasicParser;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.OptionBuilder;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.formula.eval.NotImplementedException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * Convert an excel file format into a comma/tab separated value format.
 * 
 * @author Masaharu Hayashi
 */
public class Excel2CSV {
	/**
	 * Creating locale object from system environment variables.
	 */
	Locale locale = new Locale(System.getProperty("user.language"),
			System.getProperty("user.country"));
	/**
	 * Reading localized language resource.
	 */
	ResourceBundle rb = ResourceBundle.getBundle(
			Excel2CSV.class.getCanonicalName(), locale);
	/**
	 * delimiter for output file format.
	 */
	private String delimiter;
	private boolean flgDoubleQuotes;
	private String charset = "UTF-8";

	/**
	 * Constructor.
	 * 
	 * @param args
	 */
	public Excel2CSV(String[] args) {
		parseCommandLine(args);
	}

	/**
	 * Parsing command line.
	 * 
	 * @param args
	 */
	private void parseCommandLine(String[] args) {
		Options ops = new Options();

		// help
		OptionBuilder.withLongOpt("help");
		OptionBuilder.hasArg(false);
		OptionBuilder.withDescription(rb.getString("OPT_HELP_MSG"));
		ops.addOption(OptionBuilder.create("h"));

		// input file
		OptionBuilder.withLongOpt("input");
		OptionBuilder.hasArg(true);
		OptionBuilder.isRequired(true);
		OptionBuilder.withDescription(rb.getString("OPT_INPUT_MSG"));
		ops.addOption(OptionBuilder.create("i"));

		// input file
		OptionBuilder.withLongOpt("charset");
		OptionBuilder.hasArg(true);
		OptionBuilder.isRequired(false);
		OptionBuilder.withDescription(rb.getString("OPT_CHARSET_MSG"));
		ops.addOption(OptionBuilder.create("ch"));

		OptionBuilder.withDescription(rb.getString("OPT_CSV_MSG"));
		ops.addOption(OptionBuilder.create("csv"));

		OptionBuilder.withDescription(rb.getString("OPT_TABTXT_MSG"));
		ops.addOption(OptionBuilder.create("tab"));

		OptionBuilder.withDescription(rb.getString("OPT_DOUBLE_QUOTES"));
		ops.addOption(OptionBuilder.create("dq"));

		BasicParser parser = new BasicParser();

		try {
			CommandLine cl = parser.parse(ops, args);
			delimiter = ",";
			if (cl.hasOption("tab")) {
				delimiter = "\t";
			} else if (cl.hasOption("csv")) {
				delimiter = ",";
			}

			if (cl.hasOption("dq")) {
				// double quotes
				flgDoubleQuotes = true;
			} else {
				flgDoubleQuotes = false;
			}
			if (cl.hasOption("h")) {
				showHelp(ops);
			}

			charset = "UTF-8";
			if (cl.hasOption("ch")) {
				String val = cl.getOptionValue("ch");
				if (val != null && val.length() > 0) {
					charset = val;
				}
			}

			if (cl.hasOption("i")) {
				String val = cl.getOptionValue("i");
				if (val != null && val.length() > 0) {
					readExcel(val);
				}
			} else {
				showHelp(ops);
			}
		} catch (ParseException e) {
			showHelp(ops);
		}
	}

	/**
	 * Print help message.
	 * 
	 * @param ops
	 */
	private void showHelp(Options ops) {
		HelpFormatter f = new HelpFormatter();
		f.printHelp(rb.getString("USAGE_MSG"), ops);
	}

	/**
	 * Read Excel format file.
	 * 
	 * @param filename
	 */
	private void readExcel(String filename) {
		try {
			FileInputStream fin = new FileInputStream(filename);
			if (flgDoubleQuotes) {
				delimiter = "\"" + delimiter + "\"";
			}
			try {
				Workbook wb = WorkbookFactory.create(fin);
				
				// HSSFFormulaEvaluator.evaluateAllFormulaCells(wb);

				for (int i = 0; i < wb.getNumberOfSheets(); i++) {
					Sheet sheet = wb.getSheetAt(i);
					String sheetname = sheet.getSheetName();

					String fname = sheetname;
					if(delimiter.equals("\t")) {
						fname = fname + ".tsv";
					}else {
						fname = fname + ".csv";
					}
					File file = new File(fname);
					FileOutputStream fout = new FileOutputStream(file);
					OutputStreamWriter ow = new OutputStreamWriter(fout,
							charset);
					BufferedWriter bw = new BufferedWriter(ow);

					for (Iterator<Row> rowIter = sheet.rowIterator(); rowIter
							.hasNext();) {
						Row row = rowIter.next();

						String tmp = "";
						if (flgDoubleQuotes) {
							tmp = "\"";
						}
						if (row != null) {
							for (int k = 0; k < row.getLastCellNum(); k++) {
								Cell cell = row.getCell(k);
								// CellValue celv = evaluator.evaluate(cell);
								if (cell == null) {
									tmp = tmp + delimiter;
									continue;
								}
								switch (cell.getCellType()) {
								case Cell.CELL_TYPE_BLANK:
									tmp = tmp +" "+ delimiter;
									break;
								case Cell.CELL_TYPE_NUMERIC:
									tmp = tmp + getNumericValue(cell)+delimiter;
									break;
								case Cell.CELL_TYPE_STRING:
									tmp = tmp + getStringValue(cell)+delimiter;
									break;
								case Cell.CELL_TYPE_FORMULA:
									try {
										FormulaEvaluator evaluator = wb.getCreationHelper()
												.createFormulaEvaluator();

									CellValue value = evaluator.evaluate(cell);
									
									if (value.getCellType() == Cell.CELL_TYPE_NUMERIC) {
										tmp = tmp + getNumericValue(cell)+delimiter;
									} else if (value.getCellType() == Cell.CELL_TYPE_STRING) {
										tmp = tmp + getStringValue(cell)+delimiter;

									}
									}catch(FormulaParseException e) {
										// error
										tmp = tmp +" "+ delimiter;
										System.err.println(e.getLocalizedMessage());
									}catch(NotImplementedException e) {
										// error
										tmp = tmp +" "+ delimiter;
										System.err.println(e.getLocalizedMessage());
									}
									break;
								default:
									tmp = tmp +" "+ delimiter;
								}

							}
							tmp = tmp.substring(0, tmp.length() - 1);
						}
						bw.write(tmp + "\n");
					}
					bw.flush();
					bw.close();
					ow.close();
					fout.close();
					System.gc();
				}
			} catch (InvalidFormatException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}

	/**
	 * Get String value from a cell.
	 * @param cell
	 * @return
	 */
	private String getStringValue(Cell cell) {
		String ret="";
		ret = cell.getStringCellValue();
		ret = ret.replaceAll("\\n"," ");
		ret = ret.replaceAll("\\r"," ");
		return ret;
	}

	/**
	 * Get Numeric value from cell.
	 * @param cell
	 * @return
	 */
	private String getNumericValue(Cell cell) {
		String ret = "";
		double d = cell.getNumericCellValue();
		if (d == (int) d) {
			ret = ret + (int) d;
		} else {
			ret = ret + d ;
		}
		return ret;
	}

	/**
	 * main method.
	 * 
	 * @param args
	 */
	public static void main(String[] args) {
		new Excel2CSV(args);

	}

}