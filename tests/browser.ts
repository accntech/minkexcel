import { chromium } from "../tools/node_modules/@playwright/test/index.mjs";
import { strict as assert } from "node:assert";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { unzip } from "./helpers.js";

// Optional tooling is isolated from the dependency-free library and unit tests.
// Use a consumer entry outside the side-effect-free package. Bun 1.4.0 drops
// definitions when the package's re-export barrel is itself the build entry.
const directory = await mkdtemp(join(tmpdir(), "minkexcel-browser-"));
let module: string;
try {
  const entrypoint = join(directory, "consumer.js");
  await writeFile(
    entrypoint,
    `export * from ${JSON.stringify(new URL("../src/index.ts", import.meta.url).pathname)};`,
  );
  const build = await Bun.build({
    entrypoints: [entrypoint],
    target: "browser",
    format: "esm",
  });
  if (!build.success)
    throw new AggregateError(build.logs, "Browser bundle failed");
  module = await build.outputs[0].text();
} finally {
  await rm(directory, { recursive: true, force: true });
}
const worker = `import {readWorkbook,writeWorkbook} from '/xlsx.js';
self.onmessage=async()=>{
 try {
 const foreign=await readWorkbook(new Uint8Array(await (await fetch('/fixture.xlsx')).arrayBuffer()),{preserveTemplate:true});
 foreign.getWorksheet(1).getCell('A3').value='000000999';
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
  assert.equal(result.tin, "000000999");
  assert.equal(result.date, "2026-10-08T12:30:00.000Z");
  assert.deepEqual(errors, []);
  const exported = new Uint8Array(result.bytes);
  assert.equal(new DataView(exported.buffer).getUint16(8, true), 8);
  assert.match(unzip(exported).get("xl/sharedStrings.xml")!, /001234567/);
	const fixtureParts = unzip(new Uint8Array(await Bun.file(new URL('fixtures/exceljs-1900.xlsx', import.meta.url)).arrayBuffer()));
	assert.equal(unzip(exported).get('xl/styles.xml'), fixtureParts.get('xl/styles.xml'), 'Worker template export retains foreign styles');
	assert.equal(unzip(exported).get('xl/worksheets/_rels/sheet1.xml.rels'), fixtureParts.get('xl/worksheets/_rels/sheet1.xml.rels'), 'Worker template export retains hyperlink relationships');
  console.log(
    "Browser module and Web Worker: compressed import, export, Blob roundtrip, scalar values and template preservation passed.",
  );
} finally {
  await browser?.close();
  server.stop(true);
}
