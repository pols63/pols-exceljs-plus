import { PDate } from 'pols-date'
import { rules } from '../src/index'

const original = {
	uno: new PDate,
	dos: 'qwe',
}

const otro = 56

const resultados = rules().isObject({
	uno: rules().isDate(),
	dos: rules({required: true}).isDate(),
	// tres: rules({ required: true }).isNaturalNoZero(),
	// cuatro: rules({ label: 'Cuatro', default: {} }).isObject({
	// 	aaa: rules().isDateTime()
	// }),
	// cinco: rules().isArray(i => rules({ label: `Elem ${i + 1}` }).isObject({
	// 	ccc: rules().isNumber()
	// }))
}).validate<typeof original>('')

console.log(original, resultados)

// console.log(rules().isTime().validate('23:59:59'))