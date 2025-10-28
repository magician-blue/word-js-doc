import os
import shutil
import subprocess
import tempfile
import json
import re
from textwrap import dedent
from typing import Dict, List, Any

def parse_tsc_errors(stdout: str) -> List[Dict[str, Any]]:
    """Parse TypeScript compiler errors into structured format."""
    errors = []
    # Pattern: filename(line,col): error TSxxxx: message
    pattern = r'(.+?)\((\d+),(\d+)\):\s+(error|warning)\s+TS(\d+):\s+(.+)'
   
    for match in re.finditer(pattern, stdout):
        filename, line, col, severity, code, message = match.groups()
        errors.append({
            'file': os.path.basename(filename),
            'line': int(line),
            'column': int(col),
            'severity': severity,
            'code': f'TS{code}',
            'message': message.strip()
        })
   
    return errors


def format_error_summary(errors: List[Dict[str, Any]]) -> str:
    """Format errors into a readable summary."""
    if not errors:
        return "No errors found"
   
    summary = []
    for i, err in enumerate(errors, 1):
        summary.append(
            f"  [{i}] Line {err['line']}:{err['column']} - "
            f"{err['severity'].upper()} {err['code']}\n"
            f"      {err['message']}"
        )
   
    return "\n".join(summary)


def check_officejs_ts(code_ts: str, use_preview: bool = True) -> dict:
    """
    Type-check a TypeScript Office.js snippet in an isolated temp project.
    Installs typescript + @types/office-js (+ preview) into the temp folder,
    then runs the local tsc with --noEmit.
    """
    npm = "npm.cmd" if os.name == "nt" else "npm"
    if shutil.which(npm) is None:
        raise RuntimeError("npm not found on PATH. Install Node.js or add npm to PATH.")

    # 1) temp project
    td = tempfile.mkdtemp(prefix="ojs_tscheck_")
    ts_path = os.path.join(td, "snippet.ts")
    with open(ts_path, "w", encoding="utf-8") as f:
        f.write(dedent(code_ts).strip() + "\n")

    # 2) package.json
    p = subprocess.run([npm, "init", "-y"], cwd=td, capture_output=True, text=True)
    if p.returncode != 0:
        raise RuntimeError(f"npm init failed:\n{p.stdout}\n{p.stderr}")

    # 3) deps
    pkgs = ["typescript", "@types/office-js"]
    if use_preview:
        pkgs.append("@types/office-js-preview")
    p = subprocess.run([npm, "i", "-D", "--silent", "--no-audit", "--no-fund", *pkgs],
                       cwd=td, capture_output=True, text=True)
    if p.returncode != 0:
        raise RuntimeError(f"npm install failed:\n{p.stdout}\n{p.stderr}")

    # 4) tsconfig.json (make TS look in this folder for types)
    tsconfig = {
        "compilerOptions": {
            "target": "ES2018",
            "module": "ESNext",
            "lib": ["ES2018", "DOM"],
            "strict": False,
            "moduleResolution": "Node",
            "typeRoots": [os.path.join(td, "node_modules", "@types")],  # explicit
            "types": ["office-js"] + (["office-js-preview"] if use_preview else []),
            "skipLibCheck": True
        },
        "include": ["snippet.ts"]
    }
    with open(os.path.join(td, "tsconfig.json"), "w", encoding="utf-8") as f:
        json.dump(tsconfig, f, indent=2)

    # 5) run local tsc (no need for npx)
    tsc = os.path.join(td, "node_modules", ".bin", "tsc")
    if os.name == "nt":
        # prefer .cmd if present
        if os.path.exists(tsc + ".cmd"):
            tsc = tsc + ".cmd"
    cmd = [tsc, "--noEmit"]
    proc = subprocess.run(cmd, cwd=td, capture_output=True, text=True)

    # Parse errors for better readability
    errors = parse_tsc_errors(proc.stdout)
    error_summary = format_error_summary(errors)
   
    # Determine success/failure
    success = proc.returncode == 0
    status = "PASS" if success else "FAIL"

    return {
        # Status
        "success": success,
        "status": status,
        "returncode": proc.returncode,
       
        # Errors
        "errors": errors,
        "error_count": len(errors),
        "error_summary": error_summary,
       
        # Raw output
        "stdout": proc.stdout,
        "stderr": proc.stderr,
       
        # Metadata
        "temp_dir": td,
        "ts_file": ts_path,
        "tsconfig": os.path.join(td, "tsconfig.json"),
        "cmd": " ".join(cmd),
        "use_preview": use_preview
    }


# ---------------- example ----------------
if __name__ == "__main__":
    print("=" * 70)
    print("Office.js TypeScript Static Checker - Test Suite")
    print("=" * 70)
   
    # Test 1: Valid basic scenario
    print("\n[Test 1] Valid basic paragraph insertion:")
    print("-" * 70)
    snippet1 = """
    /// <reference types="office-js" />
    async function addTitle() {
      await Word.run(async (context) => {
        const p = context.document.body.insertParagraph("Hello from Office.js", Word.InsertLocation.start);
        p.styleBuiltIn = Word.BuiltInStyleName.title;
        await context.sync();
      });
    }
    """
    res1 = check_officejs_ts(snippet1, use_preview=True)
    print(f"{'✅ PASS' if res1['success'] else '❌ FAIL'} - {res1['error_count']} error(s)")
    if not res1['success']:
        print(f"\n{res1['error_summary']}")
   
    # Test 2: Invalid property (should fail)
    print("\n[Test 2] Invalid property (should fail):")
    print("-" * 70)
    snippet2 = """
    /// <reference types="office-js" />
    async function testInvalid() {
      await Word.run(async (context) => {
        const p = context.document.body.insertParagraph("Test", Word.InsertLocation.start);
        p.nonExistentProperty = "this should fail";
        await context.sync();
      });
    }
    """
    res2 = check_officejs_ts(snippet2, use_preview=True)
    print(f"{'❌ FAIL (expected)' if not res2['success'] else '⚠️  UNEXPECTED PASS'} - {res2['error_count']} error(s)")
    if not res2['success']:
        print(f"\n{res2['error_summary']}")
        print (res2)
   
    # Test 3: Table manipulation
    print("\n[Test 3] Table creation and formatting:")
    print("-" * 70)
    snippet3 = """
    /// <reference types="office-js" />
    async function createTable() {
      await Word.run(async (context) => {
        const body = context.document.body;
        const table = body.insertTable(3, 4, Word.InsertLocation.end, [
          ["Header 1", "Header 2", "Header 3", "Header 4"],
          ["Cell 1", "Cell 2", "Cell 3", "Cell 4"],
          ["Cell 5", "Cell 6", "Cell 7", "Cell 8"]
        ]);
        table.styleBuiltIn = Word.BuiltInStyleName.gridTable1Light;
        table.font.bold = true;
        await context.sync();
      });
    }
    """
    res3 = check_officejs_ts(snippet3, use_preview=True)
    print(f"{'✅ PASS' if res3['success'] else '❌ FAIL'} - {res3['error_count']} error(s)")
    if not res3['success']:
        print(f"\n{res3['error_summary']}")
   
    # Test 4: Font formatting
    print("\n[Test 4] Font and text formatting:")
    print("-" * 70)
    snippet4 = """
    /// <reference types="office-js" />
    async function formatText() {
      await Word.run(async (context) => {
        const range = context.document.getSelection();
        range.font.name = "Arial";
        range.font.size = 14;
        range.font.bold = true;
        range.font.italic = true;
        range.font.color = "#FF0000";
        range.font.underline = Word.UnderlineType.single;
        await context.sync();
      });
    }
    """
    res4 = check_officejs_ts(snippet4, use_preview=True)
    print(f"{'✅ PASS' if res4['success'] else '❌ FAIL'} - {res4['error_count']} error(s)")
    if not res4['success']:
        print(f"\n{res4['error_summary']}")
   
    # Test 5: List operations
    print("\n[Test 5] List creation:")
    print("-" * 70)
    snippet5 = """
    /// <reference types="office-js" />
    async function createList() {
      await Word.run(async (context) => {
        const body = context.document.body;
        const p1 = body.insertParagraph("Item 1", Word.InsertLocation.end);
        const p2 = body.insertParagraph("Item 2", Word.InsertLocation.end);
        const p3 = body.insertParagraph("Item 3", Word.InsertLocation.end);
       
        p1.listItem.listString = "1. ";
        p2.listItem.listString = "2. ";
        p3.listItem.listString = "3. ";
       
        await context.sync();
      });
    }
    """
    res5 = check_officejs_ts(snippet5, use_preview=True)
    print(f"✅ Result: {'PASS' if res5['returncode'] == 0 else 'FAIL'}")
    if res5['returncode'] != 0:
        print(f"Errors:\n{res5['stdout']}")
   
    # Test 6: Content control operations
    print("\n[Test 6] Content control creation:")
    print("-" * 70)
    snippet6 = """
    /// <reference types="office-js" />
    async function createContentControl() {
      await Word.run(async (context) => {
        const range = context.document.getSelection();
        const cc = range.insertContentControl();
        cc.title = "My Content Control";
        cc.tag = "test-tag";
        cc.appearance = Word.ContentControlAppearance.boundingBox;
        cc.cannotDelete = true;
        await context.sync();
      });
    }
    """
    res6 = check_officejs_ts(snippet6, use_preview=True)
    print(f"✅ Result: {'PASS' if res6['returncode'] == 0 else 'FAIL'}")
    if res6['returncode'] != 0:
        print(f"Errors:\n{res6['stdout']}")
   
    # Test 7: Search and replace
    print("\n[Test 7] Search and replace:")
    print("-" * 70)
    snippet7 = """
    /// <reference types="office-js" />
    async function searchAndReplace() {
      await Word.run(async (context) => {
        const results = context.document.body.search("old text", {matchCase: false});
        results.load("font");
        await context.sync();
       
        for (let i = 0; i < results.items.length; i++) {
          results.items[i].insertText("new text", Word.InsertLocation.replace);
        }
        await context.sync();
      });
    }
    """
    res7 = check_officejs_ts(snippet7, use_preview=True)
    print(f"✅ Result: {'PASS' if res7['returncode'] == 0 else 'FAIL'}")
    if res7['returncode'] != 0:
        print(f"Errors:\n{res7['stdout']}")
   
    # Test 8: Wrong InsertLocation value (should fail)
    print("\n[Test 8] Invalid enum value (should fail):")
    print("-" * 70)
    snippet8 = """
    /// <reference types="office-js" />
    async function testInvalidEnum() {
      await Word.run(async (context) => {
        const p = context.document.body.insertParagraph("Test", "invalidLocation" as any);
        await context.sync();
      });
    }
    """
    res8 = check_officejs_ts(snippet8, use_preview=True)
    print(f"⚠️  Result: {'PASS (using any type)' if res8['returncode'] == 0 else 'FAIL'}")
   
    # Test 9: Complex document structure
    print("\n[Test 9] Complex document structure:")
    print("-" * 70)
    snippet9 = """
    /// <reference types="office-js" />
    async function createComplexDoc() {
      await Word.run(async (context) => {
        const body = context.document.body;
        body.clear();
       
        // Add title
        const title = body.insertParagraph("Document Title", Word.InsertLocation.start);
        title.styleBuiltIn = Word.BuiltInStyleName.title;
       
        // Add heading
        const h1 = body.insertParagraph("Section 1", Word.InsertLocation.end);
        h1.styleBuiltIn = Word.BuiltInStyleName.heading1;
       
        // Add paragraph
        const p = body.insertParagraph("This is a regular paragraph.", Word.InsertLocation.end);
       
        // Insert page break
        const breakPara = body.insertParagraph("", Word.InsertLocation.end);
        breakPara.insertBreak(Word.BreakType.page, Word.InsertLocation.after);
       
        await context.sync();
      });
    }
    """
    res9 = check_officejs_ts(snippet9, use_preview=True)
    print(f"✅ Result: {'PASS' if res9['returncode'] == 0 else 'FAIL'}")
    if res9['returncode'] != 0:
        print(f"Errors:\n{res9['stdout']}")
   
    # Test 10: Missing await (syntax warning)
    print("\n[Test 10] Missing await (may not error in strict:false):")
    print("-" * 70)
    snippet10 = """
    /// <reference types="office-js" />
    async function missingAwait() {
      Word.run(async (context) => {
        const p = context.document.body.insertParagraph("Test", Word.InsertLocation.start);
        context.sync();  // Missing await
      });
    }
    """
    res10 = check_officejs_ts(snippet10, use_preview=True)
    print(f"ℹ️  Result: {'PASS (no strict mode)' if res10['returncode'] == 0 else 'FAIL'}")
   
    # Test 11: Table cell manipulation
    print("\n[Test 11] Table cell manipulation:")
    print("-" * 70)
    snippet11 = """
    /// <reference types="office-js" />
    async function formatTableCells() {
      await Word.run(async (context) => {
        const table = context.document.body.tables.getFirst();
        const cell = table.getCell(0, 0);
        cell.body.clear();
        cell.body.insertText("Modified Cell", Word.InsertLocation.end);
        cell.shadingColor = "#FFFF00";
        cell.verticalAlignment = Word.VerticalAlignment.center;
        await context.sync();
      });
    }
    """
    res11 = check_officejs_ts(snippet11, use_preview=True)
    print(f"✅ Result: {'PASS' if res11['returncode'] == 0 else 'FAIL'}")
    if res11['returncode'] != 0:
        print(f"Errors:\n{res11['stdout']}")
   
    # Test 12: Hyperlink operations
    print("\n[Test 12] Hyperlink insertion:")
    print("-" * 70)
    snippet12 = """
    /// <reference types="office-js" />
    async function insertHyperlink() {
      await Word.run(async (context) => {
        const range = context.document.getSelection();
        range.insertText("Visit Microsoft", Word.InsertLocation.replace);
        range.hyperlink = "https://www.microsoft.com";
        range.font.color = "#0000FF";
        range.font.underline = Word.UnderlineType.single;
        await context.sync();
      });
    }
    """
    res12 = check_officejs_ts(snippet12, use_preview=True)
    print(f"✅ Result: {'PASS' if res12['returncode'] == 0 else 'FAIL'}")
    if res12['returncode'] != 0:
        print(f"Errors:\n{res12['stdout']}")
   
    # Test 13: Image insertion
    print("\n[Test 13] Image insertion:")
    print("-" * 70)
    snippet13 = """
    /// <reference types="office-js" />
    async function insertImage() {
      await Word.run(async (context) => {
        const body = context.document.body;
        const base64Image = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";
        const pic = body.insertInlinePicture(base64Image, Word.InsertLocation.end);
        pic.width = 100;
        pic.height = 100;
        await context.sync();
      });
    }
    """
    res13 = check_officejs_ts(snippet13, use_preview=True)
    print(f"✅ Result: {'PASS' if res13['returncode'] == 0 else 'FAIL'}")
    if res13['returncode'] != 0:
        print(f"Errors:\n{res13['stdout']}")
   
    # Test 14: Paragraph alignment and spacing
    print("\n[Test 14] Paragraph alignment and spacing:")
    print("-" * 70)
    snippet14 = """
    /// <reference types="office-js" />
    async function formatParagraph() {
      await Word.run(async (context) => {
        const p = context.document.body.insertParagraph("Formatted paragraph", Word.InsertLocation.end);
        p.alignment = Word.Alignment.centered;
        p.lineSpacing = 1.5;
        p.spaceAfter = 12;
        p.spaceBefore = 6;
        p.leftIndent = 36;
        p.firstLineIndent = 18;
        await context.sync();
      });
    }
    """
    res14 = check_officejs_ts(snippet14, use_preview=True)
    print(f"✅ Result: {'PASS' if res14['returncode'] == 0 else 'FAIL'}")
    if res14['returncode'] != 0:
        print(f"Errors:\n{res14['stdout']}")
   
    # Test 15: Section and page setup
    print("\n[Test 15] Section and page setup:")
    print("-" * 70)
    snippet15 = """
    /// <reference types="office-js" />
    async function setupPage() {
      await Word.run(async (context) => {
        const section = context.document.sections.getFirst();
        section.load("body");
        await context.sync();
       
        const header = section.getHeader(Word.HeaderFooterType.primary);
        header.insertParagraph("Header Text", Word.InsertLocation.end);
       
        const footer = section.getFooter(Word.HeaderFooterType.primary);
        footer.insertParagraph("Footer Text", Word.InsertLocation.end);
       
        await context.sync();
      });
    }
    """
    res15 = check_officejs_ts(snippet15, use_preview=True)
    print(f"✅ Result: {'PASS' if res15['returncode'] == 0 else 'FAIL'}")
    if res15['returncode'] != 0:
        print(f"Errors:\n{res15['stdout']}")
   
    # Test 16: Range operations
    print("\n[Test 16] Range manipulation:")
    print("-" * 70)
    snippet16 = """
    /// <reference types="office-js" />
    async function manipulateRange() {
      await Word.run(async (context) => {
        const range = context.document.body.getRange();
        range.select(Word.SelectionMode.end);
        const newRange = range.expandTo(range.paragraphs.getFirst().getRange());
        newRange.font.highlightColor = "#FFFF00";
        await context.sync();
      });
    }
    """
    res16 = check_officejs_ts(snippet16, use_preview=True)
    print(f"✅ Result: {'PASS' if res16['returncode'] == 0 else 'FAIL'}")
    if res16['returncode'] != 0:
        print(f"Errors:\n{res16['stdout']}")
   
    # Test 17: Document properties
    print("\n[Test 17] Document properties access:")
    print("-" * 70)
    snippet17 = """
    /// <reference types="office-js" />
    async function accessDocProperties() {
      await Word.run(async (context) => {
        const props = context.document.properties;
        props.load("title,author,subject");
        await context.sync();
       
        console.log("Title:", props.title);
        console.log("Author:", props.author);
       
        props.title = "New Document Title";
        props.author = "Test Author";
        await context.sync();
      });
    }
    """
    res17 = check_officejs_ts(snippet17, use_preview=True)
    print(f"✅ Result: {'PASS' if res17['returncode'] == 0 else 'FAIL'}")
    if res17['returncode'] != 0:
        print(f"Errors:\n{res17['stdout']}")
   
    # Test 18: Track changes and comments
    print("\n[Test 18] Comments (preview API):")
    print("-" * 70)
    snippet18 = """
    /// <reference types="office-js-preview" />
    async function addComment() {
      await Word.run(async (context) => {
        const range = context.document.getSelection();
        const comment = range.insertComment("This is a test comment");
        comment.authorName = "Test User";
        await context.sync();
      });
    }
    """
    res18 = check_officejs_ts(snippet18, use_preview=True)
    print(f"✅ Result: {'PASS' if res18['returncode'] == 0 else 'FAIL'}")
    if res18['returncode'] != 0:
        print(f"Errors:\n{res18['stdout']}")
   
    # Test 19: Incorrect API version (should fail)
    print("\n[Test 19] Incorrect API method (should fail):")
    print("-" * 70)
    snippet19 = """
    /// <reference types="office-js" />
    async function useNonExistentAPI() {
      await Word.run(async (context) => {
        const body = context.document.body;
        body.nonExistentMethod();
        await context.sync();
      });
    }
    """
    res19 = check_officejs_ts(snippet19, use_preview=True)
    print(f"❌ Result: {'FAIL (expected)' if res19['returncode'] != 0 else 'UNEXPECTED PASS'}")
    if res19['returncode'] != 0:
        print(f"Expected errors found:\n{res19['stdout'][:200]}...")
   
    # Test 20: Multiple sync operations
    print("\n[Test 20] Multiple sync operations:")
    print("-" * 70)
    snippet20 = """
    /// <reference types="office-js" />
    async function multipleSyncs() {
      await Word.run(async (context) => {
        const body = context.document.body;
       
        const p1 = body.insertParagraph("First paragraph", Word.InsertLocation.end);
        await context.sync();
       
        const p2 = body.insertParagraph("Second paragraph", Word.InsertLocation.end);
        p2.font.bold = true;
        await context.sync();
       
        const p3 = body.insertParagraph("Third paragraph", Word.InsertLocation.end);
        p3.font.italic = true;
        await context.sync();
      });
    }
    """
    res20 = check_officejs_ts(snippet20, use_preview=True)
    print(f"✅ Result: {'PASS' if res20['returncode'] == 0 else 'FAIL'}")
    if res20['returncode'] != 0:
        print(f"Errors:\n{res20['stdout']}")
   
    # Summary
    print("\n" + "=" * 70)
    print("Test Summary:")
    print("=" * 70)
    tests = [
        ("Basic paragraph", res1),
        ("Invalid property", res2),
        ("Table creation", res3),
        ("Font formatting", res4),
        ("List operations", res5),
        ("Content controls", res6),
        ("Search & replace", res7),
        ("Invalid enum", res8),
        ("Complex structure", res9),
        ("Missing await", res10),
        ("Table cells", res11),
        ("Hyperlinks", res12),
        ("Image insertion", res13),
        ("Para alignment", res14),
        ("Headers/footers", res15),
        ("Range operations", res16),
        ("Doc properties", res17),
        ("Comments (preview)", res18),
        ("Non-existent API", res19),
        ("Multiple syncs", res20)
    ]
   
    for i, (name, result) in enumerate(tests, 1):
        status = "✅ PASS" if result.get('success', result['returncode'] == 0) else "❌ FAIL"
        error_info = f" ({result.get('error_count', 0)} errors)" if not result.get('success', result['returncode'] == 0) else ""
        print(f"Test {i:2d}: {name:20s} {status}{error_info}")
   
    passed = sum(1 for _, r in tests if r.get('success', r['returncode'] == 0))
    failed = len(tests) - passed
   
    print(f"\n{'='*70}")
    print(f"Total Tests: {len(tests)} | Passed: {passed} | Failed: {failed}")
    print(f"{'='*70}")
   
    print("\nNote: Temp directories created (will need manual cleanup):")
    print(f"  Last test temp dir: {res20['temp_dir']}")