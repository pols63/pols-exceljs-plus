import assert from 'node:assert'
import { PXls } from '../src/index'

async function runTests() {
	const workbook = new PXls()
	workbook.addWorksheet('TestSheet')
	const sheet = workbook.getWorksheet('TestSheet')

	// Write some test data in row 1
	// Columns: A1 (1,1) = 123, B1 (1,2) = "John Doe", C1 (1,3) = "2026-06-24", D1 (1,4) = null (empty), E1 (1,5) = "  Custom text  ", F1 (1,6) = "extra"
	sheet.getCell(1, 1).value = 123
	sheet.getCell(1, 2).value = "John Doe"
	sheet.getCell(1, 3).value = new Date("2026-06-24T12:00:00.000Z")
	sheet.getCell(1, 4).value = null
	sheet.getCell(1, 5).value = "  Custom text  "
	sheet.getCell(1, 6).value = "extra"

	// 1. Basic schema verification (sequential read from column 1 onwards)
	const basicSchema = {
		id: Number,
		name: String,
		birthDate: Date,
	}
	const res1 = sheet.getValuesBySchema(basicSchema, 'row', 1, 1)
	console.log('Result 1 (Basic):', res1)
	assert.strictEqual(res1.id, 123)
	assert.strictEqual(res1.name, "John Doe")
	assert.ok(res1.birthDate instanceof Date)
	assert.strictEqual(res1.birthDate.toISOString().substring(0, 10), "2026-06-24")

	// 2. Custom cellIndex & Null value verification
	const indexAndNullSchema = {
		name: { type: String, cellIndex: 1 },         // Col 2 ("John Doe")
		extra: { type: String, cellIndex: 5 },        // Col 6 ("extra")
		missingNull: { type: String, cellIndex: 3 }, // Col 4 (null)
	}
	const res2 = sheet.getValuesBySchema(indexAndNullSchema, 'row', 1, 1)
	console.log('Result 2 (Index & Null):', res2)
	assert.strictEqual(res2.name, "John Doe")
	assert.strictEqual(res2.extra, "extra")
	assert.strictEqual(res2.missingNull, null)

	// 3. Null representation check (if no content, returns null)
	const nullCheckSchema = {
		emptyCell: { type: String, cellIndex: 3 }, // Col 4 (null)
	}
	const res3 = sheet.getValuesBySchema(nullCheckSchema, 'row', 1, 1)
	console.log('Result 3 (Null check):', res3)
	assert.strictEqual(res3.emptyCell, null)

	// 4. Custom parser verification
	const customParserSchema = {
		name: { parse: (val: any) => typeof val === 'string' ? val.toUpperCase() : val, cellIndex: 1 },
		trimmed: { type: String, cellIndex: 4 }, // A5 has leading/trailing spaces
	}
	const res4 = sheet.getValuesBySchema(customParserSchema, 'row', 1, 1)
	console.log('Result 4 (Custom Parser & Trimming):', res4)
	assert.strictEqual(res4.name, "JOHN DOE")
	assert.strictEqual(res4.trimmed, "Custom text") // check replacement/trimming

	// 5. Short-hand string types verification
	const shorthandStringTypes = {
		id: 'number',
		name: 'string',
		birthDate: 'date',
	}
	const res5 = sheet.getValuesBySchema(shorthandStringTypes, 'row', 1, 1)
	console.log('Result 5 (Shorthand string types):', res5)
	assert.strictEqual(res5.id, 123)
	assert.strictEqual(res5.name, "John Doe")
	assert.ok(res5.birthDate instanceof Date)

	// 6. Verification of consistent decoration on other worksheet access paths
	// Test worksheets getter
	const sheetsFromGetter = workbook.worksheets
	assert.ok(sheetsFromGetter.length > 0)
	assert.ok(typeof sheetsFromGetter[0].getValuesBySchema === 'function')
	console.log('Worksheets getter decoration check passed!')

	// Test eachSheet iterator
	let iteratorCalled = false
	workbook.eachSheet((ws) => {
		assert.ok(typeof ws.getValuesBySchema === 'function')
		iteratorCalled = true
	})
	assert.ok(iteratorCalled)
	console.log('eachSheet iterator decoration check passed!')

	console.log('All tests passed successfully!')
}

runTests().catch(err => {
	console.error('Test failed:', err)
	process.exit(1)
})
