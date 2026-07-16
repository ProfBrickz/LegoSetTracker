// Classes
/**
 * @template K
 * @template V
 * @extends {Map<K, V>}
 */
export class ClassMap extends Map {
	/**
	 * @private
	 * @type {new (...args: any[]) => V}
	 */
	classType;

	/**
	 * @param {new (...args: any[]) => V} classType
	 * @param {Iterable<readonly [K, V]>} [iterable]
	*/
	constructor(classType, iterable) {
		super(iterable);
		this.classType = classType;
	}

	/**
	 * @param {K} key
	 * @param {V} value
	 * @throws {TypeError} if value is not an instance of classTyp
	 */
	set(key, value) {
		if (!(value instanceof this.classType)) {
			throw new TypeError(`Value must be an instance of ${this.classType.name}`);
		}

		return super.set(key, value);
	}
}

export class TypedClass {
	/**
	 * @private
	 * @type {Symbol}
	 */
	brand = Symbol("");
}
