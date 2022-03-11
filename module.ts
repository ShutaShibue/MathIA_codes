import { CellAddress, readFile, utils, writeFile } from "xlsx";
/**
 *
 * @param {String} area - Fetch area "A1:D4"
 * @param {String} sheetname - sheetname
 * @param {String} filename - Excel Worksheet
 * @returns {Array}
 */
export async function read(
  area: string,
  sheetname: string,
  filename: string
): Promise<any[][]> {
  const book = readFile(filename);
  if (!book) throw new Error("Book not found!");
  const ws = book.Sheets[sheetname];
  if (!book) throw new Error("Sheet not found!");
  let arr: Array<any> = [];
  let decodeRange = await getdecodeRange(area);
  for (
    let colIdx: any = decodeRange.s.c, m = 0;
    colIdx <= decodeRange.e.c;
    colIdx++, m++
  ) {
    arr[colIdx] = [];
    for (
      let rowIdx = decodeRange.s.r, n = 0;
      rowIdx <= decodeRange.e.r;
      rowIdx++, n++
    ) {
      // セルのアドレスを取得する
      let address = await getencodeRange({ r: rowIdx, c: colIdx });
      let cell = ws[address];
      let k:any;
      if (cell?.v == "undefined") k = "";
      else if (!isNaN(cell.v)) k = Math.round(cell.v * 1000) / 1000;
      else k = cell.v;
      arr[m][n] = k;
    }
  }
  return arr;
}

/**
 *
 * @param {Array} data data to write in
 * @param {String} area area to write in, eg. A1:D3
 * @param {String}  sheetname sheetname to write
 * @param {String} filename The book of worksheet
 */
export async function write(
  data: Array<any>,
  area: string,
  sheetname: string,
  filename: string
) {
    const book = readFile(filename);
  if (!book) throw new Error("Book not found!");
    const ws = book.Sheets[sheetname];
  if (!ws) throw new Error("Sheet not found!");
  const decodeRange = await getdecodeRange(area);
  for (
    let colIdx = decodeRange.s.c, m = 0;
    colIdx <= decodeRange.e.c;
    colIdx++, m++
  ) {
    for (
      let rowIdx = decodeRange.s.r, n = 0;
      rowIdx <= decodeRange.e.r;
      rowIdx++, n++
    ) {
      const address = await getencodeRange({ r: rowIdx, c: colIdx });
      if (data[m][n] === null || undefined) {
      } else if (isNaN(data[m][n])) {
        ws[address] = {
          t: "f",
          f: data[m][n],
        };
      } else {
        ws[address] = {
          t: "n",
          v: data[m][n],
        };
      }
    }
  }
  book.Sheets[sheetname] = ws;
  writeFile(book, filename);
  return 0;
}
/**
 * 
 * @param {String} range input of Conversion
 * @returns {Range} - { s: { c: start col, r: start row }, e: { c: end col , r: end row } }

 */
 function getdecodeRange(range: string) {
  return utils.decode_range(range);
}

/**
 * 
 * @param {CellAddress} a Celladdress
 * @returns {String} Output string
 */
async function getencodeRange(a: CellAddress) {
  return utils.encode_cell(a);
}
