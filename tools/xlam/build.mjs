#!/usr/bin/env node
/**
 * Builds dist/Breakdown.xlam from the VBA source in src/.
 *
 *   node build.mjs           rebuild dist/Breakdown.xlam in place
 *   node build.mjs --check   verify dist/Breakdown.xlam matches src/ (CI)
 *
 * The existing dist/Breakdown.xlam is the template: it holds the UserForm
 * designer binaries (.frx is not in git), the sheets, and each module's
 * stored `Attribute` header. The build swaps in the code bodies from src/
 * and the ribbon XML from src/ribbon/customUI14.xml.
 *
 * Modules are written as source only (no compiled p-code) and the
 * _VBA_PROJECT stream is reset to version 0xFFFF [MS-OVBA 2.3.4.1], so
 * Excel recompiles the project from source when it loads. That is what
 * makes one file work on both Windows and Mac (32/64-bit, any VBA build).
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import CFB from "cfb";
import { unzipSync, zipSync } from "fflate";

const REPO_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const SRC_DIR = path.join(REPO_ROOT, "src");
const XLAM_PATH = path.join(REPO_ROOT, "dist", "Breakdown.xlam");
const RIBBON_SRC = path.join(SRC_DIR, "ribbon", "customUI14.xml");
const RIBBON_ZIP_PATH = "customUI/customUI14.xml";
const VBA_BIN_ZIP_PATH = "xl/vbaProject.bin";
const VBA_PROJECT_SOURCE_ONLY = Buffer.from([0xcc, 0x61, 0xff, 0xff, 0x00, 0x00, 0x00]);
// Fixed timestamp so rebuilding unchanged source gives a byte-identical file.
const ZIP_MTIME = new Date("2000-01-01T00:00:00Z");

// cfb (SheetJS) plants a "\u0001Sh33tJ5" placeholder stream in every file it
// writes, via an internal CFB.find() lookup. Answer that lookup as "present"
// so the vbaProject.bin we emit contains only Excel's own streams.
const cfbFind = CFB.find;
CFB.find = (cfb, name) => (name === "/\u0001Sh33tJ5" ? {} : cfbFind(cfb, name));

// ---------------------------------------------------------------------------
// Directory tree [MS-CFB 2.6.4]
//
// Each storage's children form a red-black tree ordered by name length,
// then by uppercase name. Office finds streams by searching that tree.
// cfb orders same-length names case-sensitively ("ModX" before "clsX") and
// links siblings as one unbalanced chain, so Excel for Mac cannot find
// streams like VBA/clsFormatType. relinkDirectoryTree() replaces those
// links with a correctly ordered, balanced tree before writing.
// ---------------------------------------------------------------------------

const RED = 0;
const BLACK = 1;
const NOSTREAM = -1;

function compareCfbNames(a, b) {
  if (a.length !== b.length) return a.length - b.length;
  const upperA = a.toUpperCase();
  const upperB = b.toUpperCase();
  return upperA < upperB ? -1 : upperA > upperB ? 1 : 0;
}

function parentPath(fullPath) {
  const trimmed = fullPath.endsWith("/") ? fullPath.slice(0, -1) : fullPath;
  return trimmed.slice(0, trimmed.lastIndexOf("/") + 1);
}

function isStorage(entry) {
  return entry.type === 1 || entry.type === 5;
}

function relinkDirectoryTree(cfb) {
  const { FileIndex: entries, FullPaths: paths } = cfb;
  for (const entry of entries) {
    entry.L = entry.R = entry.C = NOSTREAM;
    entry.color = BLACK;
  }

  for (let storage = 0; storage < entries.length; storage++) {
    if (!isStorage(entries[storage])) continue;
    const children = [];
    for (let i = 1; i < entries.length; i++) {
      if (parentPath(paths[i]) === paths[storage]) children.push(i);
    }
    children.sort((a, b) => compareCfbNames(entries[a].name, entries[b].name));

    // Midpoint splits give a minimal-height tree whose levels are all full
    // except the deepest; colouring only that level red satisfies the
    // red-black rules (no red-red edge, equal black height on every path).
    const redDepth = Math.floor(Math.log2(children.length + 1));
    const link = (lo, hi, depth) => {
      if (lo > hi) return NOSTREAM;
      const mid = (lo + hi) >> 1;
      const node = entries[children[mid]];
      node.L = link(lo, mid - 1, depth + 1);
      node.R = link(mid + 1, hi, depth + 1);
      node.color = depth === redDepth ? RED : BLACK;
      return children[mid];
    };
    entries[storage].C = link(0, children.length - 1, 0);
  }
}

/** Returns a list of problems: children Office could not find, or red-black violations. */
function validateDirectoryTree(entries) {
  const problems = [];
  const blackHeight = (sid, parentIsRed) => {
    if (sid === NOSTREAM) return 1;
    const node = entries[sid];
    const red = node.color === RED;
    if (red && parentIsRed) problems.push(`red node ${node.name} has a red parent`);
    const left = blackHeight(node.L, red);
    const right = blackHeight(node.R, red);
    if (left !== right) problems.push(`unequal black height under ${node.name}`);
    return left + (red ? 0 : 1);
  };
  const collect = (sid, out) => {
    if (sid === NOSTREAM) return out;
    out.push(sid);
    collect(entries[sid].L, out);
    collect(entries[sid].R, out);
    return out;
  };

  for (const storage of entries) {
    if (!storage || !isStorage(storage) || storage.C === NOSTREAM) continue;
    blackHeight(storage.C, false);
    for (const sid of collect(storage.C, [])) {
      const name = entries[sid].name;
      let cursor = storage.C;
      while (cursor !== NOSTREAM) {
        const order = compareCfbNames(name, entries[cursor].name);
        if (order === 0) break;
        cursor = order < 0 ? entries[cursor].L : entries[cursor].R;
      }
      if (cursor !== sid) problems.push(`lookup of ${storage.name}/${name} fails`);
    }
  }
  return problems;
}

// ---------------------------------------------------------------------------
// MS-OVBA 2.4.1 compression
// ---------------------------------------------------------------------------

function copyTokenBitCount(difference) {
  return Math.max(Math.ceil(Math.log2(difference)), 4);
}

function decompress(bytes) {
  if (bytes[0] !== 0x01) throw new Error("Invalid compressed container signature");
  const out = [];
  let pos = 1;
  while (pos < bytes.length) {
    const header = bytes[pos] | (bytes[pos + 1] << 8);
    const chunkSize = (header & 0x0fff) + 3;
    const isCompressed = (header & 0x8000) !== 0;
    const chunkEnd = Math.min(pos + chunkSize, bytes.length);
    pos += 2;
    const chunkStart = out.length;
    if (!isCompressed) {
      for (let i = 0; i < 4096; i++) out.push(bytes[pos + i]);
      pos += 4096;
      continue;
    }
    while (pos < chunkEnd) {
      const flags = bytes[pos++];
      for (let bit = 0; bit < 8 && pos < chunkEnd; bit++) {
        if ((flags & (1 << bit)) === 0) {
          out.push(bytes[pos++]);
        } else {
          const token = bytes[pos] | (bytes[pos + 1] << 8);
          pos += 2;
          const bitCount = copyTokenBitCount(out.length - chunkStart);
          const lengthMask = 0xffff >> bitCount;
          const length = (token & lengthMask) + 3;
          const offset = (token >> (16 - bitCount)) + 1;
          const from = out.length - offset;
          for (let i = 0; i < length; i++) out.push(out[from + i]);
        }
      }
    }
  }
  return Buffer.from(out);
}

function compressChunk(data, start, end) {
  const out = [];
  let cur = start;
  while (cur < end) {
    const flagIndex = out.length;
    out.push(0);
    let flags = 0;
    for (let bit = 0; bit < 8 && cur < end; bit++) {
      let bestLength = 0;
      let bestOffset = 0;
      if (cur > start) {
        const bitCount = copyTokenBitCount(cur - start);
        const maxLength = (0xffff >> bitCount) + 3;
        for (let candidate = cur - 1; candidate >= start; candidate--) {
          let length = 0;
          while (
            length < maxLength &&
            cur + length < end &&
            data[candidate + length] === data[cur + length]
          ) {
            length++;
          }
          if (length > bestLength) {
            bestLength = length;
            bestOffset = cur - candidate;
          }
        }
        if (bestLength >= 3) {
          const token = ((bestOffset - 1) << (16 - bitCount)) | (bestLength - 3);
          out.push(token & 0xff, token >> 8);
          flags |= 1 << bit;
          cur += bestLength;
          continue;
        }
      }
      out.push(data[cur++]);
    }
    out[flagIndex] = flags;
  }
  return out;
}

function compress(data) {
  const out = [0x01];
  for (let start = 0; start < data.length; start += 4096) {
    const end = Math.min(start + 4096, data.length);
    const chunk = compressChunk(data, start, end);
    if (chunk.length <= 4096) {
      const header = (chunk.length + 2 - 3) | 0x3000 | 0x8000;
      out.push(header & 0xff, header >> 8, ...chunk);
    } else if (end - start === 4096) {
      // Incompressible full chunk: store raw [MS-OVBA 2.4.1.3.10]
      out.push(0xff, 0x3f, ...data.subarray(start, end));
    } else {
      throw new Error("Final chunk is incompressible; cannot store without padding");
    }
  }
  return Buffer.from(out);
}

// ---------------------------------------------------------------------------
// dir stream [MS-OVBA 2.3.4.2]
// ---------------------------------------------------------------------------

/**
 * Walk dir records. Returns each module's name, stream, source offset, the
 * position of its MODULEOFFSET field and the byte range of its whole MODULE
 * record, plus the position of the PROJECTMODULES count.
 */
function parseDir(dir) {
  const modules = [];
  let moduleCountPos = -1;
  let current = null;
  let pos = 0;
  while (pos < dir.length) {
    const id = dir.readUInt16LE(pos);
    let size = dir.readUInt32LE(pos + 2);
    // PROJECTVERSION's size field says 4 but 6 bytes follow
    if (id === 0x0009) size += 2;
    const dataPos = pos + 6;
    const data = dir.subarray(dataPos, dataPos + size);
    if (id === 0x000f) {
      moduleCountPos = dataPos;
    } else if (id === 0x0019) {
      current = { name: data.toString("latin1"), start: pos };
      modules.push(current);
    } else if (id === 0x002b && current) {
      current.end = dataPos + size;
    } else if (id === 0x001a && current) {
      current.streamName = data.toString("latin1");
    } else if (id === 0x0031 && current) {
      current.offset = data.readUInt32LE(0);
      current.offsetFieldPos = dataPos;
    } else if (id === 0x0021 || id === 0x0022) {
      if (current) current.kind = id === 0x0021 ? "procedural" : "document/class";
    }
    pos = dataPos + size;
  }
  return { modules, moduleCountPos };
}

/** Cut the MODULE records of removed modules out of dir and lower the module count. */
function removeModulesFromDir(dir, moduleCountPos, removed) {
  let out = Buffer.from(dir);
  out.writeUInt16LE(out.readUInt16LE(moduleCountPos) - removed.length, moduleCountPos);
  for (const m of [...removed].sort((a, b) => b.start - a.start)) {
    out = Buffer.concat([out.subarray(0, m.start), out.subarray(m.end)]);
  }
  return out;
}

/**
 * Drop removed modules from the PROJECT stream [MS-OVBA 2.3.1]: their
 * "Module=", "Class=" or "BaseClass=" (form) line and their editor-window line
 * under [Workspace].
 */
function removeModulesFromProjectText(text, names) {
  let section = "";
  const kept = text.split("\r\n").filter((line) => {
    if (line.startsWith("[")) section = line;
    if (section === "" && names.some((n) => [`Module=${n}`, `Class=${n}`, `BaseClass=${n}`].includes(line))) return false;
    if (section === "[Workspace]" && names.some((n) => line.startsWith(`${n}=`))) return false;
    return true;
  });
  return kept.join("\r\n");
}

/**
 * Drop removed modules from the PROJECTwm name map [MS-OVBA 2.3.3]: pairs of
 * a null-terminated MBCS name and a null-terminated UTF-16 name, then 0x0000.
 */
function removeModulesFromNameMap(nameMap, names) {
  const kept = [];
  let pos = 0;
  while (pos + 1 < nameMap.length && !(nameMap[pos] === 0 && nameMap[pos + 1] === 0)) {
    const mbcsEnd = nameMap.indexOf(0, pos);
    const name = nameMap.toString("latin1", pos, mbcsEnd);
    let end = mbcsEnd + 1;
    while (!(nameMap[end] === 0 && nameMap[end + 1] === 0)) end += 2;
    end += 2;
    if (!names.includes(name)) kept.push(nameMap.subarray(pos, end));
    pos = end;
  }
  return Buffer.concat([...kept, Buffer.from([0, 0])]);
}

// ---------------------------------------------------------------------------
// Source handling
// ---------------------------------------------------------------------------

function listSourceFiles(dir) {
  const out = new Map();
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      for (const [name, file] of listSourceFiles(full)) out.set(name, file);
    } else if (/\.(bas|cls|frm)$/i.test(entry.name)) {
      out.set(entry.name.replace(/\.(bas|cls|frm)$/i, ""), full);
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Source checks. Excel compiles the VBA only when it loads the add-in, so
// structural mistakes that would fail there are caught before building.
// ---------------------------------------------------------------------------

const PROC_START = /^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+\w/i;
const PROC_END = /^\s*End\s+(?:Sub|Function|Property)\b/i;
const DECLARATION = /^\s*(?:(?:Public|Private|Global|Dim|Static)\s+\w|Const\s+\w|Declare\s|Type\s|Enum\s|Option\s)/i;

/**
 * Module-level declarations after the first procedure. VBA only allows them
 * in the declarations section at the top of a module; Excel reports "Only
 * comments may appear after End Sub, End Function, or End Property".
 * Returns "line N: text" for each.
 */
function misplacedDeclarations(text) {
  const found = [];
  let seenProcedure = false;
  let inProcedure = false;
  text
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .forEach((line, i) => {
      if (inProcedure) {
        if (PROC_END.test(line)) inProcedure = false;
      } else if (PROC_START.test(line)) {
        seenProcedure = inProcedure = true;
      } else if (seenProcedure && DECLARATION.test(line)) {
        found.push(`line ${i + 1}: ${line.trim()}`);
      }
    });
  return found;
}

/** Stop the build or check if any src file would not compile in Excel for a reason found above. */
function checkSources(sources) {
  const problems = [];
  for (const file of sources.values()) {
    for (const p of misplacedDeclarations(fs.readFileSync(file, "latin1"))) {
      problems.push(`${path.relative(REPO_ROOT, file)} ${p} (declarations must come before the first procedure)`);
    }
  }
  if (problems.length > 0) {
    console.error(problems.map((p) => `  - ${p}`).join("\n"));
    fail("src/ has VBA that Excel will not compile. Fix it and run the build again.");
  }
}

/** Strip the export header (VERSION, BEGIN..END, form designer block, Attribute lines). */
function sourceBody(text) {
  const lines = text.replace(/\r\n?/g, "\n").split("\n");
  let i = 0;
  if (/^VERSION\s/i.test(lines[i] ?? "")) i++;
  if (/^BEGIN\b/i.test(lines[i] ?? "") || /^Begin\s+\{/i.test(lines[i] ?? "")) {
    // BEGIN..END (class) or Begin {GUID} name ... End (form); both close at column 0
    i++;
    while (i < lines.length && !/^END\s*$/i.test(lines[i])) i++;
    i++;
  }
  while (i < lines.length && /^Attribute\s/i.test(lines[i])) i++;
  return lines.slice(i).join("\n");
}

/** Split stored module source into its Attribute header and code body. */
function splitStoredSource(text) {
  const lines = text.replace(/\r\n?/g, "\n").split("\n");
  let i = 0;
  while (i < lines.length && /^Attribute\s/i.test(lines[i])) i++;
  return { header: lines.slice(0, i), body: lines.slice(i).join("\n") };
}

function toCrlf(text) {
  return text.replace(/\n/g, "\r\n");
}

function readStoredSource(content, offset) {
  return decompress(content.subarray(offset)).toString("latin1");
}

function fail(message) {
  console.error(`xlam build: ${message}`);
  process.exit(1);
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

function loadTemplate() {
  const zip = unzipSync(fs.readFileSync(XLAM_PATH));
  if (!zip[VBA_BIN_ZIP_PATH]) fail(`${VBA_BIN_ZIP_PATH} missing from ${XLAM_PATH}`);
  const cfb = CFB.read(Buffer.from(zip[VBA_BIN_ZIP_PATH]), { type: "buffer" });
  // cfb only resolves nested paths when they start at the root ("/VBA/dir")
  const entry = (name) => {
    const found = cfbFind(cfb, `/${name}`);
    if (!found) fail(`stream ${name} missing from vbaProject.bin`);
    return found;
  };
  const dirEntry = entry("VBA/dir");
  const dir = decompress(Buffer.from(dirEntry.content));
  const { modules, moduleCountPos } = parseDir(dir);
  // PROJECT lists worksheet and workbook code modules as "Document=Name/&H..."
  const documentNames = Buffer.from(entry("PROJECT").content)
    .toString("latin1")
    .split("\r\n")
    .filter((line) => line.startsWith("Document="))
    .map((line) => line.slice("Document=".length).split("/")[0]);
  return { zip, cfb, entry, dirEntry, dir, modules, moduleCountPos, documentNames };
}

/**
 * Check src/ against the template's modules. Returns the modules (standard,
 * class or form) whose src file was deleted; those are removed by the build.
 * Document modules (Sheet1..3) without a src file keep their stored code.
 */
function matchModules(modules, sources, documentNames) {
  for (const name of sources.keys()) {
    if (!modules.some((m) => m.name === name)) {
      fail(`src module "${name}" has no counterpart in the template xlam (adding modules is not supported yet)`);
    }
  }
  return modules.filter((m) => !sources.has(m.name) && !documentNames.includes(m.name));
}

/**
 * Delete a form's designer storage [MS-OVBA 2.2.10]: the root-level storage
 * named after its module stream, holding the form's controls. Returns whether
 * the form had one.
 */
function removeDesignerStorage(cfb, streamName) {
  const storagePath = cfb.FullPaths[0] + streamName + "/";
  if (!cfb.FullPaths.includes(storagePath)) return false;
  for (const fullPath of cfb.FullPaths.filter((p) => p.startsWith(storagePath))) {
    const index = cfb.FullPaths.indexOf(fullPath);
    cfb.FullPaths.splice(index, 1);
    cfb.FileIndex.splice(index, 1);
  }
  return true;
}

function build() {
  const { zip, cfb, entry, dirEntry, dir, modules, moduleCountPos, documentNames } = loadTemplate();
  const sources = listSourceFiles(SRC_DIR);
  checkSources(sources);
  const removed = matchModules(modules, sources, documentNames);
  const removedNames = removed.map((m) => m.name);

  for (const m of modules) {
    if (removedNames.includes(m.name)) continue;
    const stream = entry(`VBA/${m.streamName}`);
    const stored = splitStoredSource(readStoredSource(Buffer.from(stream.content), m.offset));
    // Document modules without a src file (Sheet1..3) keep their stored code
    const body = sources.has(m.name)
      ? sourceBody(fs.readFileSync(sources.get(m.name), "latin1"))
      : stored.body;
    const text = toCrlf([...stored.header, body].join("\n"));
    const compressed = compress(Buffer.from(text, "latin1"));
    stream.content = compressed;
    stream.size = compressed.length;
    dir.writeUInt32LE(0, m.offsetFieldPos);
  }

  // Modules deleted from src/: drop them from dir, PROJECT and PROJECTwm, and
  // delete their code streams and, for forms, their designer storage
  let finalDir = dir;
  if (removed.length > 0) {
    finalDir = removeModulesFromDir(dir, moduleCountPos, removed);
    const project = entry("PROJECT");
    const projectText = removeModulesFromProjectText(Buffer.from(project.content).toString("latin1"), removedNames);
    project.content = Buffer.from(projectText, "latin1");
    project.size = project.content.length;
    const nameMap = entry("PROJECTwm");
    nameMap.content = removeModulesFromNameMap(Buffer.from(nameMap.content), removedNames);
    nameMap.size = nameMap.content.length;
    for (const m of removed) {
      CFB.utils.cfb_del(cfb, `/VBA/${m.streamName}`);
      removeDesignerStorage(cfb, m.streamName);
      console.log(`Removed module ${m.name} (its file is no longer in src/)`);
    }
  }

  const newDir = compress(finalDir);
  dirEntry.content = newDir;
  dirEntry.size = newDir.length;

  const vbaProject = entry("VBA/_VBA_PROJECT");
  vbaProject.content = Buffer.from(VBA_PROJECT_SOURCE_ONLY);
  vbaProject.size = VBA_PROJECT_SOURCE_ONLY.length;

  // __SRP_* streams are the compiled performance cache; stale once code changes
  for (const fullPath of [...cfb.FullPaths]) {
    if (/\/VBA\/__SRP_[^/]*$/.test(fullPath)) CFB.utils.cfb_del(cfb, fullPath);
  }
  CFB.utils.cfb_gc(cfb);
  relinkDirectoryTree(cfb);
  const treeProblems = validateDirectoryTree(cfb.FileIndex);
  if (treeProblems.length > 0) fail(`invalid directory tree:\n  ${treeProblems.join("\n  ")}`);

  zip[VBA_BIN_ZIP_PATH] = new Uint8Array(CFB.write(cfb, { type: "buffer" }));
  zip[RIBBON_ZIP_PATH] = new Uint8Array(fs.readFileSync(RIBBON_SRC));

  const entries = {};
  for (const [name, data] of Object.entries(zip)) {
    entries[name] = [data, { level: 6, mtime: ZIP_MTIME }];
  }
  fs.writeFileSync(XLAM_PATH, zipSync(entries));
  console.log(`Built ${path.relative(REPO_ROOT, XLAM_PATH)} from ${sources.size} src modules`);
}

function check() {
  const { zip, cfb, entry, modules, documentNames } = loadTemplate();
  const sources = listSourceFiles(SRC_DIR);
  checkSources(sources);
  const leftovers = matchModules(modules, sources, documentNames);
  const problems = validateDirectoryTree(cfb.FileIndex).map((p) => `vbaProject.bin: ${p}`);
  for (const m of leftovers) problems.push(`module ${m.name} was deleted from src/ but is still in the xlam`);

  for (const m of modules) {
    if (!sources.has(m.name)) continue;
    const stream = entry(`VBA/${m.streamName}`);
    const stored = splitStoredSource(readStoredSource(Buffer.from(stream.content), m.offset));
    const expected = sourceBody(fs.readFileSync(sources.get(m.name), "latin1"));
    if (stored.body !== expected) problems.push(`module ${m.name} differs from src`);
  }

  const ribbon = Buffer.from(zip[RIBBON_ZIP_PATH] ?? []);
  if (!ribbon.equals(fs.readFileSync(RIBBON_SRC))) problems.push("ribbon XML differs from src");

  if (problems.length > 0) {
    console.error(problems.map((p) => `  - ${p}`).join("\n"));
    fail("dist/Breakdown.xlam is out of date. Run `npm run build` in tools/xlam and commit the result.");
  }
  console.log(`dist/Breakdown.xlam matches src/ (${sources.size} modules + ribbon)`);
}

export {
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
};

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  if (process.argv.includes("--check")) check();
  else build();
}
