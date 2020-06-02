# Demo how to read and write Excel files in simple Groovy DSL syntax

It utilises awesome [Spreadsheet Builder](https://spreadsheet.dsl.builders) library.

Please read theirs documentation and examples. It very easy and powerful.

## Threat Excel like Map for reading (with keys in header column)

Please look at example [sample](src/main/resources/sample.xlsx) file:

| Enabled | One | Two | Three |
|---------|-----|-----|-------|
| yes     |   1 |   2 |     3 |
| no      |  11 |  22 |    33 |
| yes     | 111 | 222 |   333 |
|         |     |     |       |
| Total:  | 123 | 246 |   369 |

```groovy
readExcel(
	new File('/sample.xlsx').newInputStream()){
		'Total:' != it['Enabled']?.value?.toString()?.trim()
	}
		.each{row->
			println(row.toMap())
			println("Column <Two> value: ${row['Two'].value}")
		}
```
Output will be like:
```
[Enabled:yes, One:1.0, Two:2.0, Three:3.0]
Column <Two> value: 2.0
[Enabled:no, One:11.0, Two:22.0, Three:33.0]
Column <Two> value: 22.0
[Enabled:yes, One:111.0, Two:222.0, Three:333.0]
Column <Two> value: 222.0
```

See `info.hubbitus.FilterFromExcelTest.RowsWithHeader ReadXls per row. Like usage demo` for live example

## Read and update (write) file content!

```groovy
			processExcel(
				FilterFromExcelTest.getResource('/sample.xlsx').newInputStream()
				,'simple'
				,new File('changed.xlsx')
				,1){
				'Total:' != it['Enabled']?.value?.toString()?.trim()
			}
				.each{row->
					row['One'] = 77 // !!!!
					bufferWrite(row.toMap())
				}
```

Please note, value changed and written into file `changed.xlsx` automatically!

# Development

Fur build you may just run:

	./gradlew hadowJar

Then just run as usual:

	java -Xmx1400m -jar build/libs/groovy-mutable-excel-1.0-SNAPSHOT-all.jar

> **Warn**: Such run provided only for demo-purpose and use bundled in resources excel file.
