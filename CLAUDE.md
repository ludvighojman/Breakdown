# Breakdown — Working with this repo

## What this is

An Excel add-in for financial modelers. **The shipped product is the
VBA add-in `dist/Breakdown.xlam`**, installed locally through Excel's
Tools → Excel Add-ins (Mac) / File → Options → Add-ins (Windows). It must
work on both Mac and Windows. Its source is `src/` (modules, class modules,
forms, ribbon XML).

Changes are made one feature at a time, and every change must keep working
on both Mac and Windows Excel.

## Working on the add-in

1. Edit the VBA in `src/`. Never edit `dist/Breakdown.xlam` by hand, and never
   commit a copy that Excel saved; always commit the build output.
2. Build: `cd tools/xlam && npm install && npm run build`. This writes
   `dist/Breakdown.xlam` with the `src/` code as source only (no p-code), so
   Excel recompiles it natively on Mac and Windows at load.
3. `npm run check` (also in CI) fails if `dist/Breakdown.xlam` is out of date
   with `src/`; `npm test` tests the build tool. Build and check both refuse
   module-level declarations (`Private Const`, `Dim`, ...) placed after the
   first procedure, which Excel rejects when it compiles the module.
4. Commit `src/` and `dist/Breakdown.xlam` together.
5. The harness cannot compile VBA. Ask the user to install the built `.xlam`
   and test on Mac and on Windows.

Releases: pushing a tag `vX.Y.Z` runs `.github/workflows/release.yml`, which
checks the build and attaches `dist/Breakdown.xlam` to a GitHub Release. The
README links to `releases/latest/download/Breakdown.xlam`, so the release
asset must keep that exact name.

Build limits: form designer layouts (`.frx`, gitignored) live only inside
`dist/Breakdown.xlam`, which the build uses as its template. Removing a
standard module (`.bas`), class module (`.cls`) or form (`.frm`) is
automated: delete its file from `src/` and rebuild, and the build drops it
from the project, including a form's layout. Adding a module or form is not
automated yet. It requires creating it once in Excel's VBE, saving the
`.xlam`, exporting to `src/`, and then rebuilding.

Mac-compatibility rules for VBA:
- Any `Declare` (user32, gdi32, kernel32, ...) must be inside
  `#If Mac Then … #Else … #End If`, and every call site needs a Mac branch too,
  because VBA compiles all procedures. See `src/class modules/clResizer.cls`.
- No Windows-only COM: `CreateObject("VBScript.RegExp")`,
  `Scripting.Dictionary`, `Scripting.FileSystemObject`, `WScript.*`, and so on.
  Use plain VBA (`Like`, `InStr`, `Collection`). See the pattern helpers in
  `src/modules/AutoColorModule.bas`.
- No `SendKeys` or `Shell` for UI automation.
- Only use MSForms constants that exist in the type library. For example,
  `fmShiftMask` does not exist, so the KeyDown Shift bit needs your own
  constant (`1`). Check a name with
  `grep -ac <name> "/Applications/Microsoft Excel.app/Contents/SharedSupport/Type Libraries/fm20.tlb"`.
  VBA compiles each procedure the first time it runs, so a bad name only
  fails when that code path is used.
- Never name a form-level procedure or variable after a UserForm member,
  including the hidden designer properties `ClientWidth`, `ClientHeight`,
  `ClientLeft` and `ClientTop` (listed in every `.frm` header). The clash
  makes the whole form fail to compile, and Excel reports it as "Type mismatch"
  where another module creates the form (`Dim f As New frmX`).
- Settings persist via `ThisWorkbook.Save` into the add-in file, or via
  `SaveSetting` (the registry on Windows, a preferences plist on Mac).

Trace windows: `src/forms/frmDependents.frm` is generated from
`src/forms/frmPrecedents.frm`. Edit only `frmPrecedents.frm`, then regenerate:

```sh
cd src/forms
sed 's/Precedents/Dependents/g;s/precedents/dependents/g;s/Precedent/Dependent/g;s/precedent/dependent/g' frmPrecedents.frm \
  | sed 's/^Private Const FORMULA_FROM_ANCESTOR As Boolean = True$/Private Const FORMULA_FROM_ANCESTOR As Boolean = False/' > frmDependents.frm
```
