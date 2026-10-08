import { XlsxError, XlsxLimitError } from './model.js';
import { checkpoint } from './async.js';

// RFC 1951: stored blocks and canonical fixed/dynamic Huffman codes.
function invalid(): never {
	throw new XlsxError('Invalid DEFLATE data.');
}
class Bits {
	offset = 0;
	constructor(readonly bytes: Uint8Array) {}
	read(count: number): number {
		if (this.offset + count > this.bytes.length * 8) invalid();
		let value = 0;
		for (let bit = 0; bit < count; bit++, this.offset++)
			value |= ((this.bytes[this.offset >>> 3] >>> (this.offset & 7)) & 1) << bit;
		return value;
	}
}
class Huffman {
	private codes = new Map<number, number>();
	private maximum = 0;
	constructor(lengths: number[], allowEmpty = false) {
		const counts = new Array<number>(16).fill(0);
		for (const length of lengths) {
			if (length < 0 || length > 15) invalid();
			if (length) counts[length]++;
		}
		let available = 1,
			used = 0;
		for (let length = 1; length <= 15; length++) {
			available = available * 2 - counts[length];
			if (available < 0) invalid();
			if (counts[length]) this.maximum = length;
			used += counts[length];
		}
		if (!used && !allowEmpty) invalid();
		if (used && available && !(used === 1 && this.maximum === 1)) invalid();
		let code = 0;
		const next = new Array<number>(16).fill(0);
		for (let length = 1; length <= 15; length++) {
			code = (code + counts[length - 1]) * 2;
			next[length] = code;
		}
		lengths.forEach((length, symbol) => {
			if (length) this.codes.set((length << 16) | next[length]++, symbol);
		});
	}
	decode(bits: Bits): number {
		let code = 0;
		for (let length = 1; length <= this.maximum; length++) {
			code = code * 2 + bits.read(1);
			const symbol = this.codes.get((length << 16) | code);
			if (symbol !== undefined) return symbol;
		}
		return invalid();
	}
}
const fixedLiteral = new Huffman(
	Array.from({ length: 288 }, (_, i) => (i < 144 ? 8 : i < 256 ? 9 : i < 280 ? 7 : 8))
);
const fixedDistance = new Huffman(new Array(32).fill(5));
const codeOrder = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];
// The code-length alphabet's exact order is specified in §3.2.7.
const lengthBase = [
	3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131,
	163, 195, 227, 258
];
const lengthExtra = [
	0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0
];
const distanceBase = [
	1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049,
	3073, 4097, 6145, 8193, 12289, 16385, 24577
];
const distanceExtra = [
	0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13
];
function dynamic(bits: Bits): [Huffman, Huffman] {
	const literalCount = bits.read(5) + 257,
		distanceCount = bits.read(5) + 1,
		codeCount = bits.read(4) + 4;
	if (literalCount > 286) invalid();
	const codeLengths = new Array<number>(19).fill(0);
	for (let i = 0; i < codeCount; i++) codeLengths[codeOrder[i]] = bits.read(3);
	const codes = new Huffman(codeLengths);
	const lengths: number[] = [];
	while (lengths.length < literalCount + distanceCount) {
		const value = codes.decode(bits);
		if (value < 16) lengths.push(value);
		else {
			if (value === 16 && !lengths.length) invalid();
			const count =
				value === 16 ? bits.read(2) + 3 : value === 17 ? bits.read(3) + 3 : bits.read(7) + 11;
			if (lengths.length + count > literalCount + distanceCount) invalid();
			const length = value === 16 ? lengths.at(-1)! : 0;
			for (let i = 0; i < count; i++) lengths.push(length);
		}
	}
	if (!lengths[256]) invalid();
	return [
		new Huffman(lengths.slice(0, literalCount)),
		new Huffman(lengths.slice(literalCount), true)
	];
}
export async function inflate(
	input: Uint8Array,
	maximum: number,
	signal?: AbortSignal
): Promise<Uint8Array> {
	signal?.throwIfAborted();
	const output = new Uint8Array(maximum),
		bits = new Bits(input);
	let size = 0,
		work = 0,
		final = false;
	const reserve = (count: number) => {
		if (size + count > maximum) throw new XlsxLimitError('A workbook archive entry is too large.');
	};
	while (!final) {
		if (++work >= 4096) {
			work = 0;
			await checkpoint(signal);
		}
		final = !!bits.read(1);
		const type = bits.read(2);
		if (type === 0) {
			bits.offset = Math.ceil(bits.offset / 8) * 8;
			const length = bits.read(16),
				inverse = bits.read(16);
			if ((length ^ inverse) !== 65535 || bits.offset / 8 + length > input.length) invalid();
			reserve(length);
			output.set(input.subarray(bits.offset / 8, bits.offset / 8 + length), size);
			size += length;
			bits.offset += length * 8;
			await checkpoint(signal);
		} else if (type === 1 || type === 2) {
			const [literals, distances] = type === 1 ? [fixedLiteral, fixedDistance] : dynamic(bits);
			for (;;) {
				if (++work >= 4096) {
					work = 0;
					await checkpoint(signal);
				}
				const symbol = literals.decode(bits);
				if (symbol < 256) {
					reserve(1);
					output[size++] = symbol;
				} else if (symbol === 256) break;
				else {
					const index = symbol - 257;
					if (index >= lengthBase.length) invalid();
					const length = lengthBase[index] + bits.read(lengthExtra[index]);
					const distanceCode = distances.decode(bits);
					if (distanceCode >= distanceBase.length) invalid();
					const distance = distanceBase[distanceCode] + bits.read(distanceExtra[distanceCode]);
					if (distance > size) invalid();
					reserve(length);
					for (let i = 0; i < length; i++, size++) output[size] = output[size - distance];
				}
			}
		} else invalid();
	}
	if (Math.ceil(bits.offset / 8) !== input.length) invalid();
	signal?.throwIfAborted();
	return output.subarray(0, size);
}
