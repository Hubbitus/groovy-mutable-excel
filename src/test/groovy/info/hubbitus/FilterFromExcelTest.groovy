package info.hubbitus

import builders.dsl.spreadsheet.api.Row
import builders.dsl.spreadsheet.query.api.SpreadsheetCriteria
import builders.dsl.spreadsheet.query.api.SpreadsheetCriteriaResult
import builders.dsl.spreadsheet.query.poi.PoiSpreadsheetCriteria
import groovy.util.logging.Slf4j
import info.hubbitus.util.RowsWithHeader
import spock.lang.Specification

import static info.hubbitus.util.RowsWithHeader.processExcel
import static info.hubbitus.util.RowsWithHeader.readExcel

/**
 * @author Pavel Alexeev.
 * @since 2019-11-24 11:47.
 */
@Slf4j
class FilterFromExcelTest extends Specification{
	def 'simple ReadXls'(){
		expect:
			SpreadsheetCriteria criteria = PoiSpreadsheetCriteria.FACTORY.forStream(FilterFromExcelTest.getResource('/sample.xlsx').newInputStream())

			SpreadsheetCriteriaResult rows = criteria.query {
				sheet('simple') {
				}
			}

			rows.rows.each { Row r->
				println(r.cells)
			}
	}

	StringBuffer sb = new StringBuffer()
	Closure bufferWrite = {
		synchronized (sb) {sb.append(it).append('\n')}
	}

	/**
	 * We do not make there real asserts because of where datapipes, iteration per-row. So it is not convenient provide results.
	 * @see info.hubbitus.util.RowsWithHeader#'RowsWithHeader ReadXls' for results asserions
	 */
	def 'RowsWithHeader ReadXls per row. Like usage demo'(){
		when:
			println(row.toMap())
			println("Column <Two> value: ${row['Two'].value}")
		then:
			noExceptionThrown()

		where:
			row << readExcel(
				FilterFromExcelTest.getResource('/sample.xlsx').newInputStream()){
					'Total:' != it['Enabled']?.value?.toString()?.trim()
				}
	}

	def 'RowsWithHeader ReadXls'(){
		setup:
			RowsWithHeader rows = readExcel(
				FilterFromExcelTest.getResource('/sample.xlsx').newInputStream()){
				'Total:' != it['Enabled']?.value?.toString()?.trim()
			}

		when:
			rows.each{row->
				bufferWrite(row.toMap())
				println(row.toMap())
				println("Column <Two> value: ${row['Two'].value}")
			}
		then:
			sb.toString() == '''{Enabled=yes, One=1.0, Two=2.0, Three=3.0}
{Enabled=no, One=11.0, Two=22.0, Three=33.0}
{Enabled=yes, One=111.0, Two=222.0, Three=333.0}
'''
	}

	def 'RowsWithHeader processExcel. With update'(){
		setup:
			RowsWithHeader rows = processExcel(
				FilterFromExcelTest.getResource('/sample.xlsx').newInputStream()
				,'simple'
				,new File('changed.xlsx')
				,1){
				'Total:' != it['Enabled']?.value?.toString()?.trim()
			}
		when:
			rows.each{row->
				row['One'] = 77 // !!!!
				bufferWrite(row.toMap())
			}
		then:'Please note, value changed!'
			sb.toString() == '''{Enabled=yes, One=77.0, Two=2.0, Three=3.0}
{Enabled=no, One=77.0, Two=22.0, Three=33.0}
{Enabled=yes, One=77.0, Two=222.0, Three=333.0}
'''
		when:'And at end of cycle also new file written automatically!'
			sb = new StringBuffer()
			readExcel(new File('changed.xlsx'))
				.each{row->
					bufferWrite(row.toMap())
				}
		then:
			sb.toString() == '''{Enabled=yes, One=77.0, Two=2.0, Three=3.0}
{Enabled=no, One=77.0, Two=22.0, Three=33.0}
{Enabled=yes, One=77.0, Two=222.0, Three=333.0}
{Enabled=Total:, One=SUM(B2:B4), Two=SUM(C2:C4), Three=SUM(D2:D4)}
'''
	}

	def cleanup(){
		new File('changed.xlsx').delete()
	}
}
