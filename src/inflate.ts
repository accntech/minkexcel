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
function codeCounts(lengths: number[]): number[] {
	const counts = new Array<number>(16).fill(0);
	for (const length of lengths) {
		if (length < 0 || length > 15) invalid();
		if (length) counts[length]++;
	}
	return counts;
}
function maximumCodeLength(counts: number[], allowEmpty: boolean): number {
	let available = 1;
	let used = 0;
	let maximum = 0;
	for (let length = 1; length <= 15; length++) {
		available = available * 2 - counts[length];
		if (available < 0) invalid();
		if (counts[length]) maximum = length;
		used += counts[length];
	}
	if (!used && !allowEmpty) invalid();
	if (used && available && !(used === 1 && maximum === 1)) invalid();
	return maximum;
}
function codeStarts(counts: number[]): number[] {
	let code = 0;
	const next = new Array<number>(16).fill(0);
	for (let length = 1; length <= 15; length++) {
		code = (code + counts[length - 1]) * 2;
		next[length] = code;
	}
	return next;
}
class Huffman {
	private codes = new Map<number, number>();
	private maximum: number;
	constructor(lengths: number[], allowEmpty = false) {
		const counts = codeCounts(lengths);
		this.maximum = maximumCodeLength(counts, allowEmpty);
		const next = codeStarts(counts);
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
// Build once on the first fixed block, so importing an unused reader does no work.
let fixedTrees: [Huffman, Huffman] | undefined;
function fixed(): [Huffman, Huffman] {
	return (fixedTrees ??= [
		new Huffman(
			Array.from({ length: 288 }, (_, i) => (i < 144 ? 8 : i < 256 ? 9 : i < 280 ? 7 : 8))
		),
		new Huffman(new Array(32).fill(5))
	]);
}
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
function repeatCodeLength(value: number, bits: Bits, lengths: number[], total: number): void {
	if (value === 16 && !lengths.length) invalid();
	const count = value === 16 ? bits.read(2) + 3 : value === 17 ? bits.read(3) + 3 : bits.read(7) + 11;
	if (lengths.length + count > total) invalid();
	const length = value === 16 ? lengths.at(-1)! : 0;
	for (let index = 0; index < count; index++) lengths.push(length);
}
function dynamicLengths(bits: Bits, codes: Huffman, total: number): number[] {
	const lengths: number[] = [];
	while (lengths.length < total) {
		const value = codes.decode(bits);
		if (value < 16) lengths.push(value);
		else repeatCodeLength(value, bits, lengths, total);
	}
	return lengths;
}
function dynamic(bits: Bits): [Huffman, Huffman] {
	const literalCount = bits.read(5) + 257;
	const distanceCount = bits.read(5) + 1;
	const codeCount = bits.read(4) + 4;
	if (literalCount > 286) invalid();
	const codeLengths = new Array<number>(19).fill(0);
	for (let index = 0; index < codeCount; index++) codeLengths[codeOrder[index]] = bits.read(3);
	const lengths = dynamicLengths(bits, new Huffman(codeLengths), literalCount + distanceCount);
	if (!lengths[256]) invalid();
	return [new Huffman(lengths.slice(0, literalCount)), new Huffman(lengths.slice(literalCount), true)];
}
class InflateOutput {
	readonly bytes: Uint8Array;
	size = 0;
	work = 0;
	constructor(maximum: number) { this.bytes = new Uint8Array(maximum); }
	private reserve(count: number): void {
		if (this.size + count > this.bytes.length)
			throw new XlsxLimitError('A workbook archive entry is too large.');
	}
	literal(value: number): void {
		this.reserve(1);
		this.bytes[this.size++] = value;
	}
	stored(value: Uint8Array): void {
		this.reserve(value.length);
		this.bytes.set(value, this.size);
		this.size += value.length;
	}
	copy(length: number, distance: number): void {
		if (distance > this.size) invalid();
		this.reserve(length);
		const end = this.size + length;
		for (let index = this.size; index < end; index++) this.bytes[index] = this.bytes[index - distance];
		this.size = end;
	}
}
function storedBlock(bits: Bits, output: InflateOutput): void {
	bits.offset = Math.ceil(bits.offset / 8) * 8;
	const length = bits.read(16);
	const inverse = bits.read(16);
	if ((length ^ inverse) !== 65535 || bits.offset / 8 + length > bits.bytes.length) invalid();
	output.stored(bits.bytes.subarray(bits.offset / 8, bits.offset / 8 + length));
	bits.offset += length * 8;
}
function backReference(symbol: number, bits: Bits, distances: Huffman, output: InflateOutput): void {
	const index = symbol - 257;
	if (index >= lengthBase.length) invalid();
	const length = lengthBase[index] + bits.read(lengthExtra[index]);
	const distanceCode = distances.decode(bits);
	if (distanceCode >= distanceBase.length) invalid();
	const distance = distanceBase[distanceCode] + bits.read(distanceExtra[distanceCode]);
	output.copy(length, distance);
}
async function compressedBlock(bits: Bits, trees: [Huffman, Huffman], output: InflateOutput, signal?: AbortSignal): Promise<void> {
	const [literals, distances] = trees;
	for (;;) {
		if (++output.work >= 4096) {
			output.work = 0;
			await checkpoint(signal);
		}
		const symbol = literals.decode(bits);
		if (symbol === 256) return;
		if (symbol < 256) output.literal(symbol);
		else backReference(symbol, bits, distances, output);
	}
}
export async function inflate(input: Uint8Array, maximum: number, signal?: AbortSignal): Promise<Uint8Array> {
	signal?.throwIfAborted();
	if (signal) await checkpoint(signal);
	const output = new InflateOutput(maximum);
	const bits = new Bits(input);
	let final = false;
	while (!final) {
		if (++output.work >= 4096) {
			output.work = 0;
			await checkpoint(signal);
		}
		final = !!bits.read(1);
		const type = bits.read(2);
		if (type === 0) {
			storedBlock(bits, output);
			await checkpoint(signal);
		} else if (type === 1 || type === 2) {
			await compressedBlock(bits, type === 1 ? fixed() : dynamic(bits), output, signal);
		} else invalid();
	}
	if (Math.ceil(bits.offset / 8) !== input.length) invalid();
	signal?.throwIfAborted();
	return output.bytes.subarray(0, output.size);
}
