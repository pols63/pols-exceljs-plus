import path from 'node:path'
import stream from 'node:stream'
import exceljs, { Style } from 'exceljs'
import { PUtilsString } from 'pols-utils'
import { PDate } from 'pols-date'

export type PColumn = {
	label?: string
	color?: string
	backgroundColor?: string
	width?: number
	span?: number
	colSpan?: number
	rowSpan?: number
	children?: PColumn[]
}

export type PValue = string | number | boolean | null | undefined | Date | PDate

export type PCell = PValue | {
	value: PValue
	color?: string
	backgroundColor?: string
	span?: number
	numberFormat?: string
	vAlign?: Style['alignment']['vertical']
	hAlign?: Style['alignment']['horizontal']
	wrapText?: boolean
}

export type PPage = {
	name: string
	title?: string
	columns: PColumn[][]
	rows: PCell[][]
	defaultNumberFormat?: string
}

type PSkips = Record<string, number[]>

/* Pintado de columnas */
const cellColumnPaint = (sheet: exceljs.Worksheet, r: number, c: number, column: PColumn, skips: PSkips) => {
	skips[r.toString()]?.sort((a, b) => {
		if (a < b) return -1
		if (a > b) return 1
		return 0
	})
	const row = sheet.getRow(r)
	row.font = {
		name: "Calibri",
		bold: true,
		size: 8,
	}
	row.alignment = {
		vertical: "top",
		wrapText: true,
	}

	let cFinal = c
	const skiped = skips[r.toString()]
	if (skiped) {
		const indexSkiped = skiped.indexOf(cFinal)
		if (indexSkiped > -1) {
			for (let i = indexSkiped; i < skiped.length; i++) {
				const cSkiped = skiped[i]
				if (cFinal == cSkiped) {
					cFinal++
				} else {
					break
				}
			}
		}
	}

	const cell = sheet.getCell(r, cFinal)
	cell.border = {
		top: { style: "thin" },
		left: { style: "thin" },
		bottom: { style: "thin" },
		right: { style: "thin" },
	}
	if (column.backgroundColor) {
		if (column.backgroundColor.match(/^#?[0-9A-F]{3}$/i)) {
			column.backgroundColor = `#${column.backgroundColor[1]}${column.backgroundColor[1]}${column.backgroundColor[2]}${column.backgroundColor[2]}${column.backgroundColor[3]}${column.backgroundColor[3]}`
		} else if (!column.backgroundColor.match(/^#?[0-9A-F]{6}$/i)) {
			throw new Error(`'backgroundColor' para la columna ${cFinal} no tiene el formato correcto`)
		}
		cell.fill = {
			type: "pattern",
			pattern: "solid",
			fgColor: {
				argb: `00${column.backgroundColor.replace('#', '')}`,
			},
			bgColor: {
				argb: `00${column.backgroundColor.replace('#', '')}`,
			},
		}
	}

	if (column.color) {
		if (column.color.match(/^#?[0-9A-F]{3}$/i)) {
			column.color = `#${column.color[1]}${column.color[1]}${column.color[2]}${column.color[2]}${column.color[3]}${column.color[3]}`
		} else if (!column.color.match(/^#?[0-9A-F]{6}$/i)) {
			throw new Error(`'color' para la columna ${cFinal} no tiene el formato correcto`)
		}
		cell.value = {
			richText: [{
				font: { color: { argb: `00${column.color.replace('#', '')}` }, size: 8, bold: true }, text: column.label ?? ''
			}]
		}
	} else {
		if ('label' in column) cell.value = column.label ?? ''
	}

	if (column.width) sheet.getColumn(cFinal).width = column.width

	if (column.colSpan || column.rowSpan) {
		const colSpan = Math.max(Math.ceil(column.colSpan ?? 0), 1) - 1
		const rowSpan = Math.max(Math.ceil(column.rowSpan ?? 0), 1) - 1
		if (colSpan || rowSpan) {
			try {
				sheet.mergeCells(r, cFinal, r + rowSpan, cFinal + colSpan)
			} catch (error) {
				throw new Error(`Error al aplicar merge en ${r}, ${cFinal}, ${r + rowSpan}, ${cFinal + colSpan}: ${error.message}`)
			}
			if (colSpan) cFinal += colSpan
		}
		/* Si se ha definido un rowSpan, se registran los saltos */
		if (rowSpan) {
			for (let i = 1; i <= rowSpan; i++) {
				let reference = skips[(r + i).toString()]
				if (!reference) {
					reference = []
					skips[(r + i).toString()] = reference
				}
				for (let j = 0; j <= colSpan; j++) {
					reference.push(cFinal + j)
				}
			}
		}
	}
	return cFinal + 1
}

const setValueCell = ({ sheetCell, value }: {
	sheetCell: exceljs.Cell
	value: PCell
}) => {

	if (value == null) {
		sheetCell.value = null
	} else if (typeof value == 'string' || typeof value == 'number' || typeof value == 'boolean' || value instanceof Date || 'utcTimestamp' in value) {
		if (typeof value == 'string' || typeof value == 'number' || typeof value == 'boolean' || value instanceof Date) {
			sheetCell.value = value
		} else {
			sheetCell.value = new Date(value.utcTimestamp)
		}
	} else if (value != null && typeof value == 'object' && 'value' in value) {
		if (value.backgroundColor) {
			sheetCell.fill = {
				type: "pattern",
				pattern: "solid",
				fgColor: {
					argb: `00${(value.backgroundColor).replace('#', '')}`,
				},
				bgColor: {
					argb: `00${(value.backgroundColor).replace('#', '')}`,
				},
			}
		}
		if (value.color) sheetCell.font.color.argb = `00${value.color.replace('#', '')}`
		if (value.numberFormat) sheetCell.numFmt = value.numberFormat
		if (value.vAlign) sheetCell.alignment.vertical = value.vAlign
		if (value.hAlign) sheetCell.alignment.horizontal = value.hAlign
		if (value.wrapText) sheetCell.alignment.wrapText = value.wrapText

		try {
			setValueCell({
				sheetCell,
				value: value.value
			})
		} catch (error) {
			throw new Error(`Se ha intentado asignar un valor no válido para la celda ${sheetCell.row},${sheetCell.col}: ${value.value}. ${error.message}`)
		}
	}
}

export const report = async (...pages: PPage[]) => {
	const workbook = new exceljs.Workbook
	for (const page of pages) {
		const sheet = workbook.addWorksheet(page.name)
		let r = 1
		if (page.title) {
			sheet.getCell(1, 1).value = page.title
			sheet.getCell(1, 1).font = {
				name: "Calibri",
				bold: true,
				size: 14,
			}
			r = 2
		}

		const skips: PSkips = {}
		for (const rowColumn of page.columns) {
			let c = 1
			for (const column of rowColumn) {
				c = cellColumnPaint(sheet, r, c, column, skips)
			}
			r++
		}

		for (const row of page.rows) {
			let c = 1
			const rowInSheet = sheet.getRow(r)
			rowInSheet.font = {
				name: "Calibri",
				size: 8,
			}
			rowInSheet.alignment = {
				vertical: "top",
				wrapText: true,
			}
			for (const cell of row) {
				const sheetCell = sheet.getCell(r, c++)
				setValueCell({
					sheetCell,
					value: cell
				})
			}
			r++
		}
	}

	const passThrough = new stream.PassThrough
	await workbook.xlsx.write(passThrough)
	return stream.Readable.from(passThrough)
}

export type PSchemaResult<T, K extends readonly string[]> = [T] extends [never] ? Record<K[number], string | number | Date | null> : T

export type Worksheet = exceljs.Worksheet & {
	setValues: (r: string | number, c: string | number, values: PCell[][]) => void
	setRowValues: (r: string | number, c: string | number, values: PCell[]) => void
	setColumnValues: (r: string | number, c: string | number, values: PCell[]) => void
	getValuesBySchema: <T = never, K extends readonly string[] = readonly string[]>(r: number, c: number, readMode: 'row' | 'column', schema: K) => PSchemaResult<T, K>
	getValue: <T = string | number | Date | null>(r: number, c: number) => T
}

export type WorksheetCell = null | string | number | Date | {
	formula: string,
	result: string | number | Date
} | {
	text: string,
	hyperlink: string
}

export class Xls extends exceljs.Workbook {
	async readFromReadableStream(readableStream: stream.Readable) {
		await this.xlsx.read(readableStream)
	}

	async readFromBase64(content: string) {
		await this.xlsx.read(PUtilsString.toReadableStream(content, 'base64'))
	}

	async readFile(...filePath: string[]) {
		await this.xlsx.readFile(path.join(...filePath))
	}

	async writeFile(...filePath: string[]) {
		const fullFilePath = path.join(...filePath)
		await this.xlsx.writeFile(fullFilePath)
		return fullFilePath
	}

	async toReadableStream() {
		const buffer = await this.xlsx.writeBuffer()
		const readableStream = new stream.Readable
		readableStream.push(buffer)
		readableStream.push(null)
		return readableStream
	}

	getWorksheet(indexOrName: string | number): Worksheet {
		let sheet: Worksheet
		try {
			sheet = (typeof indexOrName == 'string' ? super.getWorksheet(indexOrName) : this.worksheets[indexOrName]) as Worksheet
		} catch {
			return
		}
		if (!sheet) throw new Error(`No se encontró el worksheet '${indexOrName}'`)

		sheet.setValues = (r: number, c: number, values: PCell[][]) => {
			for (const [i, rows] of values.entries()) {
				for (const [j, value] of rows.entries()) {
					const cell = sheet.getCell(r + i, c + j)
					setValueCell({
						sheetCell: cell,
						value
					})
				}
			}
		}

		sheet.setRowValues = (r: number, c: number, values: PCell[]) => {
			for (const [i, value] of values.entries()) {
				const cell = sheet.getCell(r, c + i)
				setValueCell({
					sheetCell: cell,
					value
				})
			}
		}

		sheet.setColumnValues = (r: number, c: number, values: PCell[]) => {
			for (const [i, value] of values.entries()) {
				const cell = sheet.getCell(r + i, c)
				setValueCell({
					sheetCell: cell,
					value
				})
			}
		}

		sheet.getValuesBySchema = <T = never, K extends readonly string[] = readonly string[]>(r: number, c: number, readMode: 'row' | 'column', schema: K) => {
			const response: Partial<PSchemaResult<T, K>> = {}
			for (const [i, property] of schema.entries()) {
				let value: WorksheetCell
				/* Se obtiene el valor de acuerdo al modo de lectura */
				if (readMode == 'row') {
					value = sheet.getCell(r, c + i).value as WorksheetCell
				} else {
					value = sheet.getCell(r + i, c).value as WorksheetCell
				}
				/* Se ientifica el tipo del valor */
				if (value && typeof value == 'object') {
					if (value instanceof Date) {
						const date = new Date(value.toISOString().replace('T', ' ').replace('Z', ''))
						response[property] = isNaN(date.getTime()) ? null : date
					} else if ('result' in value) { /* Cuando la celda es una fórmula */
						if (value.result && typeof value.result == 'object' && 'error' in value.result) {
							response[property] = value.result.error
						} else {
							if (value.result instanceof Date) {
								const date = new Date(value.result.toISOString().replace('T', ' ').replace('Z', ''))
								response[property] = isNaN(date.getTime()) ? null : date
							} else {
								response[property] = value.result
							}
						}
					} else if ('text' in value) { /* Cuando la celda es un hipervínculo */
						response[property] = value.text
					} else if ('richText' in value) { /* Cuando la celda es un hipervínculo */
						response[property] = (value as any).richText.map(v => v.text).join(' ')
					}
				} else if (value != null) {
					response[property] = value
				}
				if (typeof response[property] == 'string') response[property] = response[property].trim().replace(/\s/g, ' ')
			}
			return response as PSchemaResult<T, K>
		}

		sheet.getValue = <T = string | number | Date | null>(r: number, c: number): T => {
			let value = sheet.getCell(r, c).value
			if (value && typeof value == 'object') {
				if (value instanceof Date) {
					const date = new Date(value.toISOString().replace('T', ' ').replace('Z', ''))
					value = isNaN(date.getTime()) ? null : date
				} else if ('result' in value) { /* Cuando la celda es una fórmula */
					if (value.result && typeof value.result == 'object' && 'error' in value.result) {
						value = value.result.error
					} else {
						value = value.result
					}
				} else if ('text' in value) { /* Cuando la celda es un hipervínculo */
					value = value.text
				}
			}
			if (typeof value == 'string') {
				value = value.trim().replace(/\s/g, ' ')
				if (!value) value = null
			}
			return value as T
		}

		return sheet as Worksheet
	}
}