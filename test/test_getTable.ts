import assert from 'node:assert'
import { PXls } from '../src/index'

async function runTests() {
	const workbook = new PXls()
	workbook.addWorksheet('TableSheet')
	const sheet = workbook.getWorksheet('TableSheet')

	// Set up headers in Row 2
	// Col 2 (B) = "Nombres"
	// Col 3 (C) = "Edades"
	// Col 4 (D) = "Monto 1"
	// Col 5 (E) = "Monto 2"
	// Col 6 (F) = null (empty, should stop header scanning here)
	sheet.getCell(2, 2).value = "Nombres"
	sheet.getCell(2, 3).value = "Edades"
	sheet.getCell(2, 4).value = "Monto 1"
	sheet.getCell(2, 5).value = "Monto 2"
	sheet.getCell(2, 6).value = null

	// Row 3 (first data row)
	sheet.getCell(3, 2).value = "Juan"
	sheet.getCell(3, 3).value = 25
	sheet.getCell(3, 4).value = 100
	sheet.getCell(3, 5).value = 50
	sheet.getCell(3, 6).value = 999 // Should be ignored since F is not in the header range

	// Row 4
	sheet.getCell(4, 2).value = "Maria"
	sheet.getCell(4, 3).value = 30
	sheet.getCell(4, 4).value = 200
	sheet.getCell(4, 5).value = 150

	// Row 5 (completely empty for matched columns)
	sheet.getCell(5, 2).value = null
	sheet.getCell(5, 3).value = null
	sheet.getCell(5, 4).value = null
	sheet.getCell(5, 5).value = null

	// Row 6 (data that should NOT be read)
	sheet.getCell(6, 2).value = "Pedro"
	sheet.getCell(6, 3).value = 40

	console.log('--- Test Case 1: Basic schema & String/RegExp headerName matching ---')
	const schema1 = {
		nombre: { type: String, headerName: 'Nombres' },
		edad: { type: Number, headerName: /Edades/ }
	}
	// Starting at Row 2, Col 2 (B2)
	const res1 = sheet.getTableValues(schema1, 2, 2)
	console.log('Result 1:', res1)
	assert.strictEqual(res1.length, 2)
	assert.deepStrictEqual(res1[0], { nombre: 'Juan', edad: 25 })
	assert.deepStrictEqual(res1[1], { nombre: 'Maria', edad: 30 })

	console.log('--- Test Case 2: Multi-column matching (last one wins by default) ---')
	const schema2 = {
		nombre: { type: String, headerName: 'Nombres' },
		monto: { type: Number, headerName: /Monto/ } // Matches "Monto 1" and "Monto 2"
	}
	const res2 = sheet.getTableValues(schema2, 2, 2)
	console.log('Result 2:', res2)
	assert.strictEqual(res2.length, 2)
	assert.deepStrictEqual(res2[0], { nombre: 'Juan', monto: 50 })
	assert.deepStrictEqual(res2[1], { nombre: 'Maria', monto: 150 })

	console.log('--- Test Case 3: Multi-column matching with parse (sum accumulation) ---')
	const schema3 = {
		nombre: { type: String, headerName: 'Nombres' },
		montoTotal: {
			type: Number,
			headerName: /Monto/,
			parse: (val: any, prevVal?: any) => {
				return (prevVal || 0) + (val || 0)
			}
		}
	}
	const res3 = sheet.getTableValues(schema3, 2, 2)
	console.log('Result 3:', res3)
	assert.strictEqual(res3.length, 2)
	assert.deepStrictEqual(res3[0], { nombre: 'Juan', montoTotal: 150 }) // 100 + 50
	assert.deepStrictEqual(res3[1], { nombre: 'Maria', montoTotal: 350 }) // 200 + 150

	console.log('--- Test Case 4: Schema with no matching columns ---')
	const schema4 = {
		nonExistent: { type: String, headerName: 'Inexistente' }
	}
	const res4 = sheet.getTableValues(schema4, 2, 2)
	console.log('Result 4:', res4)
	assert.strictEqual(res4.length, 0) // Stops immediately because no columns are matched

	console.log('All getTable tests passed successfully!')
}

runTests().catch(err => {
	console.error('Test failed:', err)
	process.exit(1)
})
