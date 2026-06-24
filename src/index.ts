import path from 'node:path'
import stream from 'node:stream'
import exceljs, { Style } from 'exceljs'
import { PUtilsString } from 'pols-utils'
import { PDate } from 'pols-date'

export type PHeaderCell = {
	label?: string
	color?: string
	backgroundColor?: string
	width?: number
	span?: number
	colSpan?: number
	rowSpan?: number
	children?: PHeaderCell[]
}

export type PValue = string | number | boolean | null | undefined | Date | PDate

export type PCellStyle = {
	color?: string
	backgroundColor?: string
	span?: number
	numberFormat?: string
	vAlign?: Style['alignment']['vertical']
	hAlign?: Style['alignment']['horizontal']
	wrapText?: boolean
	border?: {
		top?: { style?: Style['border']['top']['style'], color?: string }
		bottom?: { style?: Style['border']['bottom']['style'], color?: string }
		left?: { style?: Style['border']['left']['style'], color?: string }
		right?: { style?: Style['border']['right']['style'], color?: string }
	}
}

export type PCellDefinition = {
	value: PValue
} & PCellStyle

export type PDataCell = PValue | PCellDefinition

export type PPage = {
	name: string
	title?: string
	headers: PHeaderCell[][]
	rows: PDataCell[][]
	defaultNumberFormat?: string
}

type PSkips = Record<string, number[]>

/* Pintado de columnas */
const cellColumnPaint = (sheet: exceljs.Worksheet, r: number, c: number, column: PHeaderCell, skips: PSkips) => {
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

const setBorder = (_type: 'top' | 'bottom' | 'left' | 'right', toSet: Style['border'], reference?: PCellStyle['border']) => {
	if (!reference?.[_type]) return
	toSet[_type] = {}
	if (reference[_type].style) toSet[_type].style = reference[_type].style
	if (reference[_type].color) toSet[_type].color = { argb: PUtilsString.padStart(reference[_type].color.replace(/^#/, ''), 8) }
}

const setValueCell = ({ sheetCell, value }: {
	sheetCell: exceljs.Cell
	value: PDataCell
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
					argb: PUtilsString.padStart(value.backgroundColor.replace(/^#/, ''), 8),
				},
				bgColor: {
					argb: PUtilsString.padStart(value.backgroundColor.replace(/^#/, ''), 8),
				},
			}
		}
		if (value.border) {
			const toSet: Style['border'] = {}
			setBorder('top', toSet, value.border)
			setBorder('bottom', toSet, value.border)
			setBorder('left', toSet, value.border)
			setBorder('right', toSet, value.border)
			sheetCell.border = toSet
		}
		if (value.color) sheetCell.font = {
			...sheetCell.font,
			color: { argb: PUtilsString.padStart(value.color.replace(/^#/, ''), 8) }
		}
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

export type PSchemaType = 'string' | 'number' | 'date' | 'boolean' | 'any' | StringConstructor | NumberConstructor | DateConstructor | BooleanConstructor | (string & {})

export type PSchemaItem = PSchemaType | {
	type?: PSchemaType
	cellIndex?: number
	parse?: (value: any) => any
}

export type PSchema = Record<string, PSchemaItem>

export type InferSchemaType<T> =
	T extends 'string' | StringConstructor ? string :
	T extends 'number' | NumberConstructor ? number :
	T extends 'date' | DateConstructor ? Date :
	T extends 'boolean' | BooleanConstructor ? boolean :
	T extends 'any' ? any :
	T extends { type: infer U } ? InferSchemaType<U> :
	any;

export type PSchemaResult<T, S extends PSchema> = [T] extends [never]
	? { [K in keyof S]: InferSchemaType<S[K]> | null }
	: T;

export type PTableSchemaItem = {
	type?: PSchemaType
	headerName: string | RegExp
	parse?: (value: any, prevValue?: any) => any
}

export type PTableSchema = Record<string, PTableSchemaItem>

export type PTableSchemaResult<T, S extends PTableSchema> = [T] extends [never]
	? { [K in keyof S]: InferSchemaType<S[K]> | null }
	: T;

export type Worksheet = exceljs.Worksheet & {
	setValues: (r: string | number, c: string | number, values: PDataCell[][], defaultStyle?: Omit<PCellDefinition, 'value'>) => void
	setRowValues: (r: string | number, c: string | number, values: PDataCell[], defaultStyle?: Omit<PCellDefinition, 'value'>) => void
	setColumnValues: (r: string | number, c: string | number, values: PDataCell[], defaultStyle?: Omit<PCellDefinition, 'value'>) => void
	getValuesBySchema: <T = never, S extends PSchema = PSchema>(schema: S, readMode: 'row' | 'column', r: number, c: number) => PSchemaResult<T, S>
	getTableValues: <T = never, S extends PTableSchema = PTableSchema>(schema: S, r: number, c: number) => PTableSchemaResult<T, S>[]
	getValue: <T = string | number | Date | null>(r: number, c: number) => T
}

export type WorksheetCell = null | string | number | Date | {
	formula: string,
	result: string | number | Date
} | {
	text: string,
	hyperlink: string
}

const getParsedCellValue = (sheet: Worksheet, row: number, col: number) => {
	let value: WorksheetCell = sheet.getCell(row, col).value as WorksheetCell
	if (value && typeof value == 'object') {
		if (value instanceof Date) {
			if (isNaN(value.getTime())) {
				return null
			} else {
				const date = new Date(value.toISOString().replace('T', ' ').replace('Z', ''))
				return isNaN(date.getTime()) ? null : date
			}
		} else if ('result' in value) { /* Cuando la celda es una fórmula */
			if (value.result && typeof value.result == 'object' && 'error' in value.result) {
				return value.result.error
			} else {
				if (value.result instanceof Date) {
					if (isNaN(value.result.getTime())) {
						return null
					} else {
						const date = new Date(value.result.toISOString().replace('T', ' ').replace('Z', ''))
						return isNaN(date.getTime()) ? null : date
					}
				} else {
					return value.result
				}
			}
		} else if ('text' in value) { /* Cuando la celda es un hipervínculo */
			return value.text
		} else if ('richText' in value) { /* Cuando la celda es un hipervínculo */
			return (value as any).richText.map(v => v.text).join(' ')
		}
	}
	if (value != null && typeof value == 'string') {
		value = value.trim().replace(/\s/g, ' ')
	}
	return value
}

const convertValueType = (val: any, type: any) => {
	if (val != null) {
		if (type === 'string' || type === String) {
			val = String(val)
		} else if (type === 'number' || type === Number) {
			const num = Number(val)
			val = isNaN(num) ? null : num
		} else if (type === 'boolean' || type === Boolean) {
			val = Boolean(val)
		} else if (type === 'date' || type === Date) {
			if (!(val instanceof Date)) {
				const d = new Date(val)
				val = isNaN(d.getTime()) ? null : d
			}
		}
	}
	if (val === undefined || val === null || val === '') {
		val = null
	}
	return val
}

export class PXls extends exceljs.Workbook {
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

	private decorateWorksheet(sheet: exceljs.Worksheet): Worksheet {
		if (!sheet) return sheet as any
		const ws = sheet as Worksheet
		if (ws.getValuesBySchema) return ws

		ws.setValues = (r: number, c: number, values: PDataCell[][], defaultStyle?: PCellStyle) => {
			for (const [i, rows] of values.entries()) {
				ws.setRowValues(r + i, c, rows, defaultStyle)
			}
		}

		ws.setRowValues = (r: number, c: number, values: PDataCell[], defaultStyle?: PCellStyle) => {
			let col = c
			for (const [i, value] of values.entries()) {
				const cell = ws.getCell(r, col + i)
				const processedValue = (typeof value == 'object' && value != null && 'value' in value) ? {
					...(defaultStyle ?? {}),
					...value
				} : {
					value: value as PValue,
					...(defaultStyle ?? {})
				}
				if ((processedValue.span ?? 0) > 1) {
					ws.mergeCells(r, col + i, r, col + processedValue.span - 1)
					col += processedValue.span - 1
				}
				setValueCell({
					sheetCell: cell,
					value: processedValue
				})
			}
		}

		ws.setColumnValues = (r: number, c: number, values: PDataCell[], defaultStyle?: PCellStyle) => {
			for (const [i, value] of values.entries()) {
				const cell = ws.getCell(r + i, c)
				const processedValue = (typeof value == 'object' && value != null && 'value' in value) ? {
					...(defaultStyle ?? {}),
					...value
				} : {
					value: value as PValue,
					...(defaultStyle ?? {})
				}
				if ((processedValue.span ?? 0) > 1) {
					ws.mergeCells(r, c, r, c + processedValue.span - 1)
				}
				setValueCell({
					sheetCell: cell,
					value: processedValue
				})
			}
		}

		ws.getValuesBySchema = <T = never, S extends PSchema = PSchema>(schema: S, readMode: 'row' | 'column', r: number, c: number) => {
			const response: any = {}
			let autoIndex = 0
			for (const [key, item] of Object.entries(schema)) {
				const normalized = typeof item === 'string' ? { type: item } :
					typeof item === 'function' ? (item === Number || item === String || item === Boolean || item === Date ? { type: item } : { parse: item }) :
						(item && typeof item === 'object' ? item as any : {})

				const cellIdx = typeof normalized.cellIndex === 'number' ? normalized.cellIndex : autoIndex
				autoIndex++

				const row = readMode == 'row' ? r : r + cellIdx
				const col = readMode == 'row' ? c + cellIdx : c

				let val = getParsedCellValue(ws, row, col)
				val = convertValueType(val, normalized.type)

				// Custom parse
				if (normalized.parse && typeof normalized.parse === 'function') {
					val = normalized.parse(val)
				}

				if (val === undefined || val === null || val === '') {
					val = null
				}

				response[key] = val
			}
			return response as PSchemaResult<T, S>
		}

		ws.getTableValues = <T = never, S extends PTableSchema = PTableSchema>(schema: S, r: number, c: number): PTableSchemaResult<T, S>[] => {
			const headers: { colIndex: number; headerText: string }[] = []
			let col = c
			while (true) {
				const val = getParsedCellValue(ws, r, col)
				if (val === null || val === undefined || val === '') {
					break
				}
				headers.push({ colIndex: col, headerText: String(val) })
				col++
			}

			const keyToCols = new Map<string, number[]>()
			const allMatchedCols = new Set<number>()
			const matchesHeader = (headerText: string, pattern: string | RegExp): boolean => {
				if (pattern instanceof RegExp) {
					return pattern.test(headerText)
				}
				return headerText === pattern
			}

			for (const [key, item] of Object.entries(schema)) {
				const matchedCols: number[] = []
				const pattern = item?.headerName
				if (pattern != null) {
					for (const h of headers) {
						if (matchesHeader(h.headerText, pattern)) {
							matchedCols.push(h.colIndex)
							allMatchedCols.add(h.colIndex)
						}
					}
				}
				keyToCols.set(key, matchedCols)
			}

			const results: any[] = []
			let rowIdx = r + 1

			while (true) {
				if (rowIdx > ws.rowCount) {
					break
				}

				let rowIsEmpty = true
				for (const colIdx of allMatchedCols) {
					const val = getParsedCellValue(ws, rowIdx, colIdx)
					if (val !== null && val !== undefined && val !== '') {
						rowIsEmpty = false
						break
					}
				}

				if (rowIsEmpty) {
					break
				}

				const rowObj: any = {}
				for (const [key, item] of Object.entries(schema)) {
					const matchedCols = keyToCols.get(key) || []
					if (matchedCols.length === 0) {
						rowObj[key] = null
						continue
					}

					let prevVal: any = undefined
					let hasPrev = false

					for (const colIdx of matchedCols) {
						let val = getParsedCellValue(ws, rowIdx, colIdx)
						val = convertValueType(val, item.type)

						if (item.parse && typeof item.parse === 'function') {
							val = item.parse(val, hasPrev ? prevVal : undefined)
						}

						if (val === undefined || val === null || val === '') {
							val = null
						}

						prevVal = val
						hasPrev = true
					}

					rowObj[key] = prevVal
				}

				results.push(rowObj)
				rowIdx++
			}

			return results as PTableSchemaResult<T, S>[]
		}

		ws.getValue = <T = string | number | Date | null>(r: number, c: number): T => {
			let value = ws.getCell(r, c).value
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

		return ws
	}

	getWorksheet(indexOrName: string | number): Worksheet {
		let sheet: exceljs.Worksheet
		try {
			sheet = typeof indexOrName == 'string' ? super.getWorksheet(indexOrName) : this.worksheets[indexOrName]
		} catch {
			return
		}
		if (!sheet) throw new Error(`No se encontró el worksheet '${indexOrName}'`)
		return this.decorateWorksheet(sheet)
	}

	addWorksheet(name: string, options?: exceljs.AddWorksheetOptions): Worksheet {
		const sheet = super.addWorksheet(name, options)
		return this.decorateWorksheet(sheet)
	}

	// @ts-ignore
	get worksheets(): Worksheet[] {
		// @ts-ignore
		return super.worksheets.map(sheet => this.decorateWorksheet(sheet))
	}

	eachSheet(iteratee: (worksheet: Worksheet, id: number) => void) {
		this.worksheets.forEach((sheet) => {
			iteratee(sheet, sheet.id)
		})
	}

	static async createReport(...pages: PPage[]) {
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
			for (const rowColumn of page.headers) {
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
}