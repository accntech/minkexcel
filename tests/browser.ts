import { chromium } from "../tools/node_modules/@playwright/test/index.mjs";
import { strict as assert } from "node:assert";
import { unzip } from "./helpers.js";

// Optional tooling is isolated from the dependency-free library and unit tests.
const build = await Bun.build({
  entrypoints: [new URL("../src/index.ts", import.meta.url).pathname],
  target: "browser",
  format: "esm",
});
if (!build.success)
  throw new AggregateError(build.logs, "Browser bundle failed");
const module = await build.outputs[0].text();
const worker = `import {readWorkbook,writeWorkbook} from '/xlsx.js';
self.onmessage=async()=>{
 try {
 const foreign=await readWorkbook(new Uint8Array(await (await fetch('/fixture.xlsx')).arrayBuffer()));
 const bytes=await writeWorkbook(foreign);
 const roundtrip=await readWorkbook(bytes);
 self.postMessage({tin:roundtrip.worksheets[0].getCell('A3').value,date:roundtrip.worksheets[0].getCell('F3').value.toISOString(),bytes},[bytes.buffer]);
 } catch(error) { self.postMessage({error:String(error)}); }
};`;
const html = `<!doctype html><html lang="en"><meta charset="utf-8"><title>MinkExcel browser compatibility</title><button>Import and export</button><output>Ready</output><script type="module">
import {Workbook,readWorkbook,writeWorkbook} from '/xlsx.js';
document.querySelector('button').onclick=async()=>{
 try {
 const book=new Workbook();book.addWorksheet('Browser').addRow(['001234567','α & <Co> 🧾',12.34,true,{formula:'1+2',result:3}]);
 const bytes=await writeWorkbook(book);const blob=new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
 const actual=await readWorkbook(new Uint8Array(await blob.arrayBuffer()));
 const worker=new Worker('/worker.js',{type:'module'});
 const imported=await new Promise((resolve,reject)=>{worker.onmessage=e=>e.data.error ? reject(new Error(e.data.error)) : resolve(e.data);worker.onerror=reject;worker.postMessage('run');});worker.terminate();
 window.result={values:actual.worksheets[0].getRow(1).values.slice(1),tin:imported.tin,date:imported.date,bytes:Array.from(imported.bytes)};
 document.querySelector('output').textContent='Passed';
 } catch(error) {document.querySelector('output').textContent=String(error);}
};</script></html>`;
const server = Bun.serve({
  hostname: "127.0.0.1",
  port: 0,
  fetch(request) {
    switch (new URL(request.url).pathname) {
      case "/xlsx.js":
        return new Response(module, {
          headers: { "Content-Type": "text/javascript" },
        });
      case "/worker.js":
        return new Response(worker, {
          headers: { "Content-Type": "text/javascript" },
        });
      case "/fixture.xlsx":
        return new Response(
          Bun.file(new URL("fixtures/exceljs-1900.xlsx", import.meta.url)),
        );
      default:
        return new Response(html, { headers: { "Content-Type": "text/html" } });
    }
  },
});
let browser;
try {
  browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(String(error)));
  await page.goto(server.url.href);
  await page.getByRole("button", { name: "Import and export" }).click();
  await page.getByText("Passed", { exact: true }).waitFor();
  const result = (await page.evaluate(
    () => (window as Window & { result: unknown }).result,
  )) as {
    values: unknown[];
    tin: string;
    date: string;
    bytes: number[];
  };
  assert.deepEqual(result.values, [
    "001234567",
    "α & <Co> 🧾",
    12.34,
    true,
    { formula: "1+2", result: 3 },
  ]);
  assert.equal(result.tin, "001234567");
  assert.equal(result.date, "2026-10-08T12:30:00.000Z");
  assert.deepEqual(errors, []);
  const exported = new Uint8Array(result.bytes);
  assert.equal(new DataView(exported.buffer).getUint16(8, true), 8);
  assert.match(unzip(exported).get("xl/sharedStrings.xml")!, /001234567/);
  console.log(
    "Browser module and Web Worker: compressed import, export, Blob roundtrip and scalar values passed.",
  );
} finally {
  await browser?.close();
  server.stop(true);
}
