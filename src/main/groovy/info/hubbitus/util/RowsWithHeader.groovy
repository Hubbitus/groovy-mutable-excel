package info.hubbitus.util


import builders.dsl.spreadsheet.api.Row
import builders.dsl.spreadsheet.query.api.SpreadsheetCriteria
import builders.dsl.spreadsheet.query.api.SpreadsheetCriteriaResult
import builders.dsl.spreadsheet.query.poi.PoiSpreadsheetCriteria
import groovy.transform.stc.ClosureParams
import groovy.transform.stc.SimpleType
import groovy.util.logging.Slf4j
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator

/**
 * Implement iteration and return always header and current row.
 *
 * @author Pavel Alexeev.
 * @since 2019-11-25 16:56.
 */
@Slf4j
public class RowsWithHeader implements Iterable, Closeable {
	Row header
	Set<Row> rows

	final Closure filterRows
	final File outputFile

	/**
	 * @param rows - rows from Excel file
	 * @param outputFile If file provided you may change values in document and at end it will be saved at this path
	 * @param headerNo No of row treated as header. Starts from 1, like in Excel!
	 * @param filterRows
	 */
	public RowsWithHeader(Set<Row> rows, File outputFile = null, int headerNo = 1, Closure filterRows = { true }) {
		this.header = rows.take(headerNo).last()
		this.rows = rows.takeRight(rows.size() - headerNo)
		this.filterRows = filterRows
		this.outputFile = outputFile
	}

	/**
	 * @link http://spockframework.org/spock/docs/1.0/data_driven_testing.html#_closing_of_data_providers
	 * @link https://www.codejava.net/coding/java-example-to-update-existing-excel-files-using-apache-poi
	 */
	@Override
	public void close() throws Throwable {
		if (outputFile) {
			write(outputFile)
		}
	}

	public void write(File outputFile = null) {
		outputFile = outputFile ?: this.outputFile
		if(outputFile){
			try {
				FileOutputStream outputStream = new FileOutputStream(outputFile.absolutePath)
				Workbook workbook = header.sheet.workbook.workbook
				// Recalculate formulas (for conditional formatting)! See https://stackoverflow.com/questions/5937373/using-apache-poi-hssf-how-can-i-refresh-all-formula-cells-at-once/5937590#5937590
				XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook) // Took long time. Around 30 seconds
				workbook.write(outputStream)
				workbook.close()
				outputStream.close()
			}
			catch (Throwable e) {
				log.error('Error happened to write updated Excel file', e)
				throw e
			}
		}
		else {
			throw new IllegalStateException('Please provide name for output file!')
		}
	}

	public static class RowsIterator implements Iterator, Closeable {
		RowsWithHeader rwh
		Iterator rwhIterator

		RowsIterator(RowsWithHeader rwh) {
			this.rwh = rwh
			rwhIterator = this.rwh.rows.iterator()
		}

		RowWithHeader current

		@Override
		boolean hasNext() {
			current = null
			while (rwhIterator.hasNext()) {
				current = new RowWithHeader(rwh.header, rwhIterator.next())
				if (rwh.filterRows(current)) {
					return true
				}
			}
			close()
			return false
		}

		@Override
		RowWithHeader next() {
			return current
		}

		@Override
		void close() throws IOException {
			rwh.close()
		}
	}

	@Override
	Iterator iterator() {
		return new RowsIterator(this)
	}

	/**
	 *
	 * @param source {@see File} or {@see InputStream} to process
	 * @param sheetName Sheet name
	 * @param filterRows Closure to filter only interested rows. F.e.:
	 * <code>
	 *{* 		'no' == it['SkipAutoTest']?.value?.trim()?.toLowerCase()
	 *}* </code>
	 * @return Iterator of rows
	 */
	static RowsWithHeader processExcel(/*File | java.io.InputStream*/ source, String sheetName, File outputFile = null, int headerNo = 1, @ClosureParams(value = SimpleType, options = ["RowsWithHeader"]) Closure filterRows = { true }) {
		SpreadsheetCriteria criteria = source instanceof File ? PoiSpreadsheetCriteria.FACTORY.forFile(source) : PoiSpreadsheetCriteria.FACTORY.forStream(source)

		SpreadsheetCriteriaResult rows
		if (sheetName){
			rows = criteria.query {
				sheet(sheetName) {}
			}
		}
		else {
			rows = criteria.query {
			}
		}
		return new RowsWithHeader(rows.rows as Set<Row>, outputFile, headerNo, filterRows)
	}

	static RowsWithHeader readExcel(/*File | java.io.InputStream*/ source, @ClosureParams(value = SimpleType, options = ["RowsWithHeader"]) Closure filterRows = { true }) {
		processExcel(source, null, null, 1, filterRows)
	}
}
