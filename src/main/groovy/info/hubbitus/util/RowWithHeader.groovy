package info.hubbitus.util

import builders.dsl.spreadsheet.api.Cell
import builders.dsl.spreadsheet.api.Row
import groovy.transform.Canonical
import groovy.transform.Memoized
import org.apache.poi.xssf.usermodel.XSSFCell

@Canonical
public class RowWithHeader {
	Row header
	Row row

	@Memoized
	List columnNames() {
		header.cells.collect { it.value }
	}

	public Cell cellByName(String name) {
		if (!(name in columnNames())) {
			throw new IllegalArgumentException("Requested column name [$name] not found in document!")
		}
		row.cells[columnNames().indexOf(name)]
	}

//		@Override
	public Cell getAt(String name) {
		cellByName(name)
	}

//		@Override
	public void putAt(String name, Object value) {
		Cell cell = cellByName(name)

		XSSFCell poiCell
		if (!cell || !cell.xssfCell) {
			poiCell = row.xssfRow.createCell(columnNames().indexOf(name))
		} else {
			poiCell = (XSSFCell) cell.xssfCell
		}
		poiCell.setCellValue(value)
	}

	/**
	 * Returns map of [ColumnName: Value]
	 */
	Map toMap(){
		columnNames().collectEntries{n->
			[(n): cellByName(n).value]
		}
	}
}
