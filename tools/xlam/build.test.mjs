import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";
import CFB from "cfb";
import { unzipSync } from "fflate";
import {
  compareCfbNames,
  compress,
  decompress,
  misplacedDeclarations,
  parseDir,
  relinkDirectoryTree,
  removeDesignerStorage,
  removeModulesFromDir,
  removeModulesFromNameMap,
  removeModulesFromProjectText,
  sourceBody,
  splitStoredSource,
  validateDirectoryTree,
} from "./build.mjs";

const SRC_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../../src");

function seededBytes(length, seed, alphabetSize) {
  const out = Buffer.alloc(length);
  let state = seed;
  for (let i = 0; i < length; i++) {
    state = (state * 1103515245 + 12345) & 0x7fffffff;
    out[i] = state % alphabetSize;
  }
  return out;
}

test("compression round-trips every VBA source file", () => {
  const files = [];
  const walk = (dir) => {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      const full = path.join(dir, entry.name);
      if (entry.isDirectory()) walk(full);
      else files.push(full);
    }
  };
  walk(SRC_DIR);
  for (const file of files) {
    const data = fs.readFileSync(file);
    assert.ok(decompress(compress(data)).equals(data), file);
  }
});

test("compression round-trips edge sizes and low/high entropy data", () => {
  for (const length of [0, 1, 3, 4095, 4096, 4097, 8192, 12289]) {
    for (const alphabet of [1, 4, 256]) {
      const data = seededBytes(length, length + alphabet, alphabet);
      if (alphabet === 256 && length % 4096 !== 0 && length > 3800) continue; // incompressible tail
      assert.ok(decompress(compress(data)).equals(data), `len=${length} alphabet=${alphabet}`);
    }
  }
});

test("sourceBody strips class, form, and module export headers", () => {
  const cls = 'VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  \'True\r\nEND\r\nAttribute VB_Name = "X"\r\nAttribute VB_Exposed = False\r\nOption Explicit\r\n';
  assert.equal(sourceBody(cls), "Option Explicit\n");

  const frm = 'VERSION 5.00\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmX\n   Caption = "X"\n   ClientHeight = 100\nEnd\nAttribute VB_Name = "frmX"\nOption Explicit\n';
  assert.equal(sourceBody(frm), "Option Explicit\n");

  const bas = 'Attribute VB_Name = "ModX"\nSub A()\nEnd Sub\n';
  assert.equal(sourceBody(bas), "Sub A()\nEnd Sub\n");
});

test("splitStoredSource keeps the stored Attribute header", () => {
  const stored = 'Attribute VB_Name = "X"\r\nAttribute VB_Base = "0{1}"\r\nOption Explicit\r\n';
  const { header, body } = splitStoredSource(stored);
  assert.deepEqual(header, ['Attribute VB_Name = "X"', 'Attribute VB_Base = "0{1}"']);
  assert.equal(body, "Option Explicit\n");
});

test("CFB names order by length, then case-insensitively", () => {
  // Regression: case-sensitive order put ModCellFormat before clsFormatType,
  // which made Excel for Mac fail to load VBA/clsFormatType.
  assert.ok(compareCfbNames("clsFormatType", "ModCellFormat") < 0);
  assert.ok(compareCfbNames("dir", "ModCAGR") < 0);
  assert.equal(compareCfbNames("vba", "VBA"), 0);
});

function fakeCfb(childNames) {
  return {
    FileIndex: [
      { name: "Root Entry", type: 5 },
      { name: "VBA", type: 1 },
      ...childNames.map((name) => ({ name, type: 2 })),
    ],
    FullPaths: [
      "Root Entry/",
      "Root Entry/VBA/",
      ...childNames.map((name) => `Root Entry/VBA/${name}`),
    ],
  };
}

test("relinked directory trees are searchable and valid red-black trees", () => {
  const pool = ["dir", "_VBA_PROJECT", "ModCellFormat", "clsFormatType", "frmPrecedents",
    "TraceUtils", "clsUISettings", "RibbonCallbacks", "ThisWorkbook", "Sheet1", "a", "B"];
  for (let count = 0; count <= 70; count++) {
    const names = Array.from({ length: count }, (_, i) => `${pool[i % pool.length]}${i}`);
    const cfb = fakeCfb(names);
    relinkDirectoryTree(cfb);
    assert.deepEqual(validateDirectoryTree(cfb.FileIndex), [], `count=${count}`);
  }
});

test("validateDirectoryTree catches a case-sensitively ordered chain", () => {
  const cfb = fakeCfb(["ModCellFormat", "clsFormatType"]);
  // What cfb produced: storage -> ModCellFormat -R-> clsFormatType
  cfb.FileIndex.forEach((e) => Object.assign(e, { L: -1, R: -1, C: -1, color: 1 }));
  cfb.FileIndex[0].C = 1;
  cfb.FileIndex[1].C = 2;
  cfb.FileIndex[2].R = 3;
  assert.ok(validateDirectoryTree(cfb.FileIndex).some((p) => p.includes("clsFormatType")));
});

function templateStreams() {
  const xlam = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../../dist/Breakdown.xlam");
  const zip = unzipSync(fs.readFileSync(xlam));
  const cfb = CFB.read(Buffer.from(zip["xl/vbaProject.bin"]), { type: "buffer" });
  const read = (name) => Buffer.from(CFB.find(cfb, `/${name}`).content);
  return { dir: decompress(read("VBA/dir")), project: read("PROJECT"), nameMap: read("PROJECTwm") };
}

test("removing a module from dir drops its record and lowers the count", () => {
  const { dir } = templateStreams();
  const { modules, moduleCountPos } = parseDir(dir);
  const victim = modules.find((m) => m.kind === "procedural");
  assert.equal(dir.readUInt16LE(moduleCountPos), modules.length);

  const after = parseDir(removeModulesFromDir(dir, moduleCountPos, [victim]));
  assert.equal(after.modules.length, modules.length - 1);
  assert.ok(!after.modules.some((m) => m.name === victim.name));
  assert.deepEqual(
    after.modules.map((m) => [m.name, m.streamName, m.kind]),
    modules.filter((m) => m !== victim).map((m) => [m.name, m.streamName, m.kind])
  );
});

test("the name map round-trips and drops only the removed module", () => {
  const { nameMap, dir } = templateStreams();
  assert.ok(removeModulesFromNameMap(nameMap, []).equals(nameMap));

  const victim = parseDir(dir).modules.find((m) => m.kind === "procedural").name;
  const after = removeModulesFromNameMap(nameMap, [victim]);
  const utf16 = Buffer.from(victim + "\0", "utf16le");
  assert.ok(nameMap.includes(utf16) && !after.includes(utf16));
  // One entry: the MBCS name and its null, then the UTF-16 name and its null
  assert.equal(after.length, nameMap.length - (victim.length + 1) - utf16.length);
});

test("PROJECT text drops the Module=, Class=, BaseClass= and [Workspace] lines only", () => {
  const text = [
    'ID="{X}"', "Document=ThisWorkbook/&H00000000", "Module=Keep", "Module=Gone", "Class=GoneClass",
    "Class=KeepClass", "BaseClass=frmKeep", "BaseClass=frmGone", 'Name="VBAProject"', "",
    "[Host Extender Info]", "&H00000001={3832D640};VBE;&H00000000", "",
    "[Workspace]", "Keep=0, 0, 0, 0, C", "Gone=0, 0, 0, 0, C", "GoneClass=0, 0, 0, 0, C",
    "frmGone=0, 0, 0, 0, C, 1, 1, 1, 1, C", "GoneToo=1, 1, 1, 1, C", "",
  ].join("\r\n");
  const after = removeModulesFromProjectText(text, ["Gone", "GoneClass", "frmGone"]).split("\r\n");
  assert.ok(!after.includes("Module=Gone") && !after.includes("Gone=0, 0, 0, 0, C"));
  assert.ok(!after.includes("Class=GoneClass") && !after.includes("GoneClass=0, 0, 0, 0, C"));
  assert.ok(!after.includes("BaseClass=frmGone") && !after.includes("frmGone=0, 0, 0, 0, C, 1, 1, 1, 1, C"));
  assert.ok(after.includes("Module=Keep") && after.includes("Keep=0, 0, 0, 0, C"));
  assert.ok(after.includes("Class=KeepClass") && after.includes("BaseClass=frmKeep"));
  assert.ok(after.includes("GoneToo=1, 1, 1, 1, C"), "only exact names are removed");
});

test("removing a form's designer storage leaves a valid tree without it", () => {
  const zip = unzipSync(fs.readFileSync(path.join(SRC_DIR, "../dist/Breakdown.xlam")));
  const cfb = CFB.read(Buffer.from(zip["xl/vbaProject.bin"]), { type: "buffer" });
  const root = cfb.FullPaths[0];
  const form = cfb.FullPaths.find((p) => p !== root && p.endsWith("/") && !p.endsWith("/VBA/"));
  const name = form.slice(root.length, -1);
  const others = cfb.FullPaths.filter((p) => !p.startsWith(form));

  assert.ok(removeDesignerStorage(cfb, name));
  assert.ok(!cfb.FullPaths.some((p) => p.startsWith(form)), "storage and its streams are gone");
  assert.deepEqual(cfb.FullPaths, others, "nothing else is touched");
  assert.equal(removeDesignerStorage(cfb, name), false);

  CFB.utils.cfb_gc(cfb);
  relinkDirectoryTree(cfb);
  assert.deepEqual(validateDirectoryTree(cfb.FileIndex), []);
  const reread = CFB.read(CFB.write(cfb, { type: "buffer" }), { type: "buffer" });
  assert.ok(!reread.FullPaths.some((p) => p.includes(`/${name}/`)));
});

test("declarations after the first procedure are reported", () => {
  const good = [
    "Option Explicit",
    "Private Const A As Long = 1",
    "Private WithEvents lst As MSForms.ListBox",
    "#If Mac Then",
    "#Else",
    "    Private Declare PtrSafe Function GetDC Lib \"user32\" (ByVal h As LongPtr) As LongPtr",
    "#End If",
    "Private Sub One()",
    "    Dim local As Long",
    "    Const INSIDE As Long = 2",
    "End Sub",
    "' A comment between procedures is fine",
    "Public Property Get Two() As Long",
    "End Property",
  ].join("\r\n");
  assert.deepEqual(misplacedDeclarations(good), []);

  const bad = ["Private Sub One()", "End Sub", "", "Private Const LATE As Long = 6", "Dim stray As Long"].join("\n");
  assert.deepEqual(misplacedDeclarations(bad), ["line 4: Private Const LATE As Long = 6", "line 5: Dim stray As Long"]);
});

test("every src file keeps its declarations before its procedures", () => {
  const files = [];
  const walk = (dir) => {
    for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
      const full = path.join(dir, e.name);
      if (e.isDirectory()) walk(full);
      else if (/\.(bas|cls|frm)$/i.test(e.name)) files.push(full);
    }
  };
  walk(SRC_DIR);
  for (const file of files) assert.deepEqual(misplacedDeclarations(fs.readFileSync(file, "latin1")), [], file);
});
