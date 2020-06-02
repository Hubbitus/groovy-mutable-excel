package info.hubbitus

import info.hubbitus.util.RowWithHeader

import static info.hubbitus.util.RowsWithHeader.processExcel

class App {
	static void main(String[] args) {
		processExcel(
			App.getResource('/Отчет ОКВЭД 14.05.2020 .xlsx').newInputStream()
			,'Выручка'
			,null
			,5){
//			'-' != it['Республика Башкортостан']?.value?.toString()?.trim() &&
				'итого' != it['ОКВЭД']?.value?.toString()?.toLowerCase()?.trim()
		}.each { RowWithHeader row->
			println (row.toMap())
		}
	}
}
